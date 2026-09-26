using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OneNoteCodeHelper.Services.Agent
{
    internal sealed class AgentOptions
    {
        internal int MaxTurns { get; set; } = 12;
        internal int MaxToolCalls { get; set; } = 48;
        internal int TimeoutSeconds { get; set; } = 600;
        internal int MaxPageChars { get; set; } = 40000;
        internal int MaxRequestChars { get; set; } = 120000;
        internal bool SendThinking { get; set; } = true;
        internal bool ReplayReasoning { get; set; } = true;
        internal bool StreamUsage { get; set; } = true;
        // 已经通过 Office16 往返探针；可为其他 Office 构建单独关闭。
        internal bool EnableParagraphSpacing { get; set; } = true;
        internal bool EnableNativeHeadings { get; set; } = true;
        internal bool EnableMixedOutlines { get; set; } = true;
        internal string FontFamily { get; set; } = "Microsoft YaHei";

        internal static AgentOptions Parse(XElement element)
        {
            var value = new AgentOptions();
            if (element == null) return value;
            value.MaxTurns = Number(element, "MaxTurns", 12, 2, 30);
            value.MaxToolCalls = Number(element, "MaxToolCalls", 48, 6, 100);
            value.TimeoutSeconds = Number(element, "TimeoutSeconds", 600, 30, 1800);
            value.MaxPageChars = Number(element, "MaxPageChars", 40000, 1000, 100000);
            value.MaxRequestChars = Number(element, "MaxRequestChars", 120000, 16000, 500000);
            value.SendThinking = Boolean(element, "SendThinking", true);
            value.ReplayReasoning = Boolean(element, "ReplayReasoning", true);
            value.StreamUsage = Boolean(element, "StreamUsage", true);
            value.EnableParagraphSpacing = Boolean(element, "EnableParagraphSpacing", true);
            value.EnableNativeHeadings = Boolean(element, "EnableNativeHeadings", true);
            value.EnableMixedOutlines = Boolean(element, "EnableMixedOutlines", true);
            var font = (string)element.Element("FontFamily");
            if (ParagraphStyles.Fonts.Contains(font)) value.FontFamily = font;
            return value;
        }
        private static int Number(XElement e, string key, int fallback, int min, int max) =>
            int.TryParse((string)e.Element(key), out var value) ? Math.Max(min, Math.Min(max, value)) : fallback;
        private static bool Boolean(XElement e, string key, bool fallback) =>
            bool.TryParse((string)e.Element(key), out var value) ? value : fallback;
    }

    internal sealed class AgentBlock
    {
        internal string Id;
        internal string ObjectId;
        internal XElement Original;
        internal XElement Draft;
        internal string Fingerprint;
        internal string Text;
        internal string ProtectedReason;
        internal string ContainerId;
        internal string ParentId;
        internal int Depth;
        internal bool Read;
        internal bool Editable => ProtectedReason == null;
        internal bool Changed => !XNode.DeepEquals(Original, Draft);
    }

    internal sealed class AgentPageSnapshot
    {
        internal readonly string SnapshotId = Guid.NewGuid().ToString("N");
        internal readonly List<AgentBlock> Blocks = new List<AgentBlock>();
        internal readonly AgentOptions Options;
        internal readonly XElement Page;
        internal readonly XElement DraftStyles;
        internal string PageId => (string)Page.Attribute("ID");
        internal string Title => (string)Page.Attribute("name") ?? "当前页面";
        internal int Revision;
        internal bool Frozen;
        internal bool SelectionOnly;
        internal static XNamespace One => OneNoteApi.One;

        internal AgentPageSnapshot(string xml, ISet<string> selection, AgentOptions options)
        {
            Page = ParsePage(xml);
            DraftStyles = new XElement("styles", Page.Elements(One + "QuickStyleDef").Select(e => new XElement(e)));
            Options = options;
            SelectionOnly = selection != null;
            var objects = Page.Descendants(One + "OE").Where(e => e.Elements(One + "T").Any()).ToList();
            var duplicate = new HashSet<string>(objects.GroupBy(e => (string)e.Attribute("objectID")).Where(g => g.Count() > 1).Select(g => g.Key));
            foreach (var oe in objects)
            {
                var objectId = (string)oe.Attribute("objectID");
                if (selection != null && !selection.Contains(objectId ?? "")) continue;
                var container = oe.Ancestors().FirstOrDefault(e => e.Parent == Page);
                var reason = string.IsNullOrEmpty(objectId) || duplicate.Contains(objectId) ? "missing_or_duplicate_id" : null;
                if (container == null || !(container.Name == One + "Title" || container.Name == One + "Outline")) reason = "unsupported_container";
                if (!options.EnableMixedOutlines && container != null && container.Descendants().Any(IsBinary)) reason = "mixed_outline_not_verified";
                if (container != null && container.Descendants().Any(e => IsBinary(e) && e.Name != One + "Image")) reason = "unsupported_outline_objects";
                if (oe.Elements().Any(IsBinary) || oe.Elements(One + "InkWord").Any()) reason = "unsupported_content";
                var text = "";
                try
                {
                    var rich = new AgentRichText(oe);
                    text = rich.Text;
                    if (string.IsNullOrWhiteSpace(text)) reason = "empty";
                    if (IsCode(oe, Page)) reason = "protected_code";
                }
                catch (Exception ex) when (ex is XmlException || ex is AiException || ex is ArgumentException)
                { reason = "unsupported_html"; }
                var original = new XElement(oe);
                Blocks.Add(new AgentBlock
                {
                    Id = "p" + (Blocks.Count + 1), ObjectId = objectId, Original = original, Draft = new XElement(original),
                    Text = text, ProtectedReason = reason, Fingerprint = Fingerprint(oe, Page),
                    ContainerId = container == null ? "" : (string)container.Attribute("objectID") ?? container.Name.LocalName,
                    ParentId = (string)oe.Ancestors(One + "OE").FirstOrDefault()?.Attribute("objectID"),
                    Depth = oe.Ancestors(One + "OE").Count()
                });
            }
            if (Blocks.Sum(b => b.Editable ? b.Text.Length : 0) > options.MaxPageChars || Blocks.Count > 1000)
                throw new AiException("页面内容超过 Agent 限额，请选择较小范围后重试。");
        }

        internal static XElement ParsePage(string xml)
        {
            using (var reader = XmlReader.Create(new System.IO.StringReader(xml), new XmlReaderSettings
            { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 16000000 }))
            {
                var page = XElement.Load(reader, LoadOptions.PreserveWhitespace);
                if (page.Name != One + "Page" || string.IsNullOrEmpty((string)page.Attribute("ID"))) throw new AiException("无效的 OneNote 页面。");
                return page;
            }
        }

        internal static HashSet<string> SelectedIds(XElement page) => new HashSet<string>(PageEditor.FindSelectedParagraphs(page)
            .Where(e => e.Elements(One + "T").Any(t => (string)t.Attribute("selected") == "all" && !string.IsNullOrWhiteSpace(OneNoteHtmlEncoder.DecodeToPlainText(t.Value)))
                || (string)e.Attribute("selected") == "all")
            .Select(e => (string)e.Attribute("objectID")).Where(id => !string.IsNullOrEmpty(id)));

        internal XElement CreateDraftPage()
        {
            var page = new XElement(Page);
            page.Elements(One + "QuickStyleDef").Remove();
            page.AddFirst(DraftStyles.Elements().Select(e => new XElement(e)));
            foreach (var b in Blocks.Where(b => b.Changed)) CopyFormat(b.Draft, AgentCommitter.Find(page, b.ObjectId));
            return page;
        }

        private static bool IsBinary(XElement e) => new[] { "Image", "InkDrawing", "InkWord", "InkParagraph", "InsertedFile", "MediaFile", "FutureObject", "HTMLBlock" }.Contains(e.Name.LocalName);
        private static bool IsCode(XElement oe, XElement page)
        {
            if (oe.AncestorsAndSelf(One + "OE").Any(PageEditor.IsCodeParagraph)) return true;
            var style = Css.Effective(oe, page);
            if (style.TryGetValue("font-family", out var font) && new[] { "consolas", "nsimsun", "courier new", "courier", "lucida console", "cascadia code", "cascadia mono" }.Contains(font)) return true;
            // 单个等宽片段也保守保护整段，避免覆盖行内代码。
            return oe.Elements(One + "T").Any(t => System.Text.RegularExpressions.Regex.IsMatch(t.Value,
                @"font-family\s*:\s*['"" ]*(Consolas|NSimSun|Courier|Cascadia|Lucida Console)", System.Text.RegularExpressions.RegexOptions.IgnoreCase));
        }

        internal static string Fingerprint(XElement oe, XElement page)
        {
            var copy = new XElement(oe);
            foreach (var a in copy.DescendantsAndSelf().Attributes().Where(a => a.Name.LocalName == "selected" || a.Name.LocalName == "lastModifiedTime").ToList()) a.Remove();
            var data = new StringBuilder(copy.ToString(SaveOptions.DisableFormatting));
            foreach (var parent in oe.AncestorsAndSelf().Where(e => e != page).Reverse())
            {
                data.Append('|').Append(parent.Name).Append(':').Append(parent.ElementsBeforeSelf(parent.Name).Count());
                data.Append(parent.Element(One + "Position"));
                foreach (var a in parent.Attributes().Where(a => a.Name.LocalName != "selected" && a.Name.LocalName != "lastModifiedTime")) data.Append(a.ToString());
                var index = (string)parent.Attribute("quickStyleIndex");
                if (index != null) data.Append(page.Elements(One + "QuickStyleDef").FirstOrDefault(d => (string)d.Attribute("index") == index));
            }
            // T 也可能直接引用 quick style。
            foreach (var index in oe.DescendantsAndSelf().Attributes("quickStyleIndex").Select(a => a.Value).Distinct())
                data.Append(page.Elements(One + "QuickStyleDef").FirstOrDefault(d => (string)d.Attribute("index") == index));
            using (var sha = SHA256.Create()) return Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(data.ToString())));
        }

        internal static DateTime Modified(XElement page)
        {
            if (!DateTime.TryParse((string)page.Attribute("lastModifiedTime"), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var time) || time == DateTime.MinValue)
                throw new AiException("无法读取页面修改时间，没有写入。");
            return time;
        }

        internal static void CopyFormat(XElement source, XElement target)
        {
            foreach (var name in new[] { "style", "alignment", "spaceBefore", "spaceAfter", "quickStyleIndex" }) target.SetAttributeValue(name, (string)source.Attribute(name));
            var src = source.Elements(One + "T").ToList();
            var dst = target.Elements(One + "T").ToList();
            if (src.Count != dst.Count)
            {
                if (dst.Count == 0) throw new AiException("目标段落没有文字片段。");
                dst[0].ReplaceWith(src.Select(t => new XElement(t)));
                foreach (var t in dst.Skip(1)) t.Remove();
                return;
            }
            for (var i = 0; i < src.Count; i++)
            {
                dst[i].SetAttributeValue("style", (string)src[i].Attribute("style"));
                dst[i].ReplaceNodes(src[i].Nodes().Select(n => n is XCData c ? (XNode)new XCData(c.Value) : new XText(((XText)n).Value)));
            }
        }

        internal static string SemanticFormat(XElement oe, XElement page)
        {
            var rich = new AgentRichText(oe);
            var index = (string)oe.Attribute("quickStyleIndex");
            var role = (string)page.Elements(One + "QuickStyleDef").FirstOrDefault(d => (string)d.Attribute("index") == index)?.Attribute("name") ?? "p";
            return rich.Signature(page, true) + "|" + ((string)oe.Attribute("alignment") ?? "left") + "|" +
                NormalizeSpacing(oe, "spaceBefore") + "|" + NormalizeSpacing(oe, "spaceAfter") + "|" + role;
        }
        private static string NormalizeSpacing(XElement e, string key) => double.TryParse((string)e.Attribute(key), NumberStyles.Float, CultureInfo.InvariantCulture, out var n)
            ? n.ToString("0.###", CultureInfo.InvariantCulture) : "0";
    }
}
