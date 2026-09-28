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
        // 按 DeepSeek 的 1M token 上下文估算（约 0.6 token/汉字），留出输出余量；上下文较短的模型在配置里调小。
        internal int MaxPageChars { get; set; } = 200000;
        internal int MaxRequestChars { get; set; } = 1000000;
        internal bool SendThinking { get; set; } = true;
        internal bool ReplayReasoning { get; set; } = true;
        internal bool StreamUsage { get; set; } = true;
        // 已经通过 Office16 往返探针；可为其他 Office 构建单独关闭。
        internal bool EnableParagraphSpacing { get; set; } = true;
        internal bool EnableNativeHeadings { get; set; } = true;
        internal bool EnableMixedOutlines { get; set; } = true;
        // 把未高亮的代码转换为插件代码框；关闭时整段等宽代码仍只保护。
        internal bool EnableCodeHighlight { get; set; } = true;
        // 列表符号、待办等标记和表格外观；关闭时这些只保护、不修改。
        internal bool EnableLists { get; set; } = true;
        internal bool EnableTags { get; set; } = true;
        internal bool EnableTableStyles { get; set; } = true;
        // 改变段落结构的工具：删空行、调整缩进、移动段落、插入段落、把分隔的文字转成表格。
        internal bool EnableBlankLineRemoval { get; set; } = true;
        internal bool EnableIndent { get; set; } = true;
        internal bool EnableMoves { get; set; } = true;
        internal bool EnableInsert { get; set; } = true;
        internal bool EnableTextTables { get; set; } = true;
        internal string FontFamily { get; set; } = "Microsoft YaHei";

        internal static AgentOptions Parse(XElement element)
        {
            var value = new AgentOptions();
            if (element == null) return value;
            value.MaxTurns = Number(element, "MaxTurns", 12, 2, 30);
            value.MaxToolCalls = Number(element, "MaxToolCalls", 48, 6, 100);
            value.TimeoutSeconds = Number(element, "TimeoutSeconds", 600, 30, 1800);
            value.MaxPageChars = Number(element, "MaxPageChars", 200000, 1000, 1000000);
            value.MaxRequestChars = Number(element, "MaxRequestChars", 1000000, 16000, 3000000);
            value.SendThinking = Boolean(element, "SendThinking", true);
            value.ReplayReasoning = Boolean(element, "ReplayReasoning", true);
            value.StreamUsage = Boolean(element, "StreamUsage", true);
            value.EnableParagraphSpacing = Boolean(element, "EnableParagraphSpacing", true);
            value.EnableNativeHeadings = Boolean(element, "EnableNativeHeadings", true);
            value.EnableMixedOutlines = Boolean(element, "EnableMixedOutlines", true);
            value.EnableCodeHighlight = Boolean(element, "EnableCodeHighlight", true);
            value.EnableLists = Boolean(element, "EnableLists", true);
            value.EnableTags = Boolean(element, "EnableTags", true);
            value.EnableTableStyles = Boolean(element, "EnableTableStyles", true);
            value.EnableBlankLineRemoval = Boolean(element, "EnableBlankLineRemoval", true);
            value.EnableIndent = Boolean(element, "EnableIndent", true);
            value.EnableMoves = Boolean(element, "EnableMoves", true);
            value.EnableInsert = Boolean(element, "EnableInsert", true);
            value.EnableTextTables = Boolean(element, "EnableTextTables", true);
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
        /// <summary>快照时的文字。修正过错别字后，草稿里的文字见 <see cref="CurrentText"/>。</summary>
        internal string Text;
        /// <summary>草稿里已排的文字修正，每项形如「原文」→「改后」。写回核验通过后进入结果，撤销时一起还原。</summary>
        internal readonly List<string> TextFixes = new List<string>();
        internal string CurrentText => TextFixes.Count == 0 ? Text : new AgentRichText(Draft).Text;
        internal string ProtectedReason;
        internal string ContainerId;
        internal string ParentId;
        /// <summary>段落所在表格的短 ID（t1…），不在已登记的表格里为 null。</summary>
        internal string TableId;
        internal int Depth;
        internal bool Read;
        /// <summary>已排入草稿的代码框转换；这些段落不再接受样式修改。</summary>
        internal AgentCodeConversion Conversion;
        internal bool Editable => ProtectedReason == null;
        /// <summary>整段等宽、还没放进代码框的代码：可以读取并转换为代码框，不能设置样式。</summary>
        internal bool CodeCandidate => ProtectedReason == "unhighlighted_code";
        internal bool Changed => !XNode.DeepEquals(Original, Draft);
    }

    /// <summary>OneNote 已识别出文字的图片。只读，短 ID 为 i1、i2…</summary>
    internal sealed class AgentImage
    {
        internal const int MaxChars = 4000;
        internal string Id;
        /// <summary>图片自己的 objectID，段落里的图片没有，取所在段落的；用来在结构草稿里找当前所在的文本框。</summary>
        internal string ObjectId;
        internal string ContainerId;
        internal string Text;
        internal bool Truncated;
    }

    internal sealed class AgentPageSnapshot
    {
        internal readonly string SnapshotId = Guid.NewGuid().ToString("N");
        internal readonly List<AgentBlock> Blocks = new List<AgentBlock>();
        internal readonly List<AgentCodeConversion> CodeConversions = new List<AgentCodeConversion>();
        /// <summary>撤销时要换回原段落的代码框。</summary>
        internal readonly List<AgentCodeUndoItem> CodeRestores = new List<AgentCodeUndoItem>();
        internal readonly List<AgentTable> Tables = new List<AgentTable>();
        internal readonly List<AgentImage> Images = new List<AgentImage>();
        internal readonly AgentOptions Options;
        internal readonly XElement Page;
        internal readonly XElement DraftStyles;
        /// <summary>草稿里的 TagDef：页面原有的加上 set_tag 新增的。提交时按内容重新对应到页面上的编号。</summary>
        internal readonly XElement DraftTags;
        /// <summary>
        /// 结构草稿：页面副本，删空行、缩进、移动、插入只改它，格式草稿仍在各段落的 Draft 里。
        /// 段落用 <see cref="AgentLayout.Key"/> 记短 ID。工具在副本上改好、校验通过后整个换掉。
        /// </summary>
        internal XElement Layout;
        /// <summary>可以调整结构的文本框（objectID）：不是标题，不含墨迹、附件等不支持的对象。</summary>
        internal readonly HashSet<string> EditableOutlines = new HashSet<string>();
        internal readonly List<AgentLayoutChange> LayoutChanges = new List<AgentLayoutChange>();
        internal readonly List<AgentInserted> Inserted = new List<AgentInserted>();
        /// <summary>撤销时整框换回的文本框。</summary>
        internal readonly List<AgentOutlineUndoItem> OutlineRestores = new List<AgentOutlineUndoItem>();
        /// <summary>
        /// 结构草稿改过的文本框，提交时整框替换。跨框移动、合并连起来的文本框在同一组里，整组写入或整组跳过，
        /// 不会一边写进、一边没删。组内按页面顺序。
        /// </summary>
        internal List<List<string>> LayoutGroups
        {
            get
            {
                var root = new Dictionary<string, string>();
                string Find(string id) { while (root[id] != id) id = root[id] = root[root[id]]; return id; }
                foreach (var c in LayoutChanges)
                {
                    foreach (var id in new[] { c.OutlineId, c.From }.Where(x => x != null)) if (!root.ContainsKey(id)) root[id] = id;
                    if (c.From != null) root[Find(c.From)] = Find(c.OutlineId);
                }
                var order = Page.Elements(One + "Outline").Select(o => (string)o.Attribute("objectID")).ToList();
                return root.Keys.ToList().GroupBy(Find).Select(g => g.OrderBy(order.IndexOf).ToList()).ToList();
            }
        }
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
            DraftTags = new XElement("tags", Page.Elements(One + "TagDef").Select(e => new XElement(e)));
            Options = options;
            SelectionOnly = selection != null;
            var tableIds = CollectTables(selection);
            CollectImages(selection);
            foreach (var outline in Page.Elements(One + "Outline"))
                if (!string.IsNullOrEmpty((string)outline.Attribute("objectID")) && ContainerReason(outline) == null) EditableOutlines.Add((string)outline.Attribute("objectID"));
            Layout = new XElement(Page);
            var objects = Page.Descendants(One + "OE").Where(e => e.Elements(One + "T").Any()).ToList();
            var layoutObjects = Layout.Descendants(One + "OE").Where(e => e.Elements(One + "T").Any()).ToList();
            var duplicate = new HashSet<string>(objects.GroupBy(e => (string)e.Attribute("objectID")).Where(g => g.Count() > 1).Select(g => g.Key));
            for (var index = 0; index < objects.Count; index++)
            {
                var oe = objects[index];
                var objectId = (string)oe.Attribute("objectID");
                if (selection != null && !selection.Contains(objectId ?? "")) continue;
                var container = oe.Ancestors().FirstOrDefault(e => e.Parent == Page);
                var reason = ContainerReason(container) ?? (string.IsNullOrEmpty(objectId) || duplicate.Contains(objectId) ? "missing_or_duplicate_id" : null);
                if (oe.Elements().Any(IsBinary) || oe.Elements(One + "InkWord").Any()) reason = "unsupported_content";
                var text = "";
                try
                {
                    var rich = new AgentRichText(oe);
                    text = rich.Text;
                    if (string.IsNullOrWhiteSpace(text)) reason = "empty";
                    var code = CodeKind(oe, rich, Page);
                    if (code != null && (reason == null || code != "unhighlighted_code")) reason = code;
                }
                catch (Exception ex) when (ex is XmlException || ex is AiException || ex is ArgumentException)
                { reason = "unsupported_html"; }
                var original = new XElement(oe);
                layoutObjects[index].SetAttributeValue(AgentLayout.Key, "p" + (Blocks.Count + 1));
                Blocks.Add(new AgentBlock
                {
                    Id = "p" + (Blocks.Count + 1), ObjectId = objectId, Original = original, Draft = new XElement(original),
                    Text = text, ProtectedReason = reason, Fingerprint = Fingerprint(oe, Page),
                    ContainerId = container == null ? "" : (string)container.Attribute("objectID") ?? container.Name.LocalName,
                    ParentId = (string)oe.Ancestors(One + "OE").FirstOrDefault()?.Attribute("objectID"),
                    TableId = oe.Ancestors(One + "Table").FirstOrDefault() is XElement table && tableIds.TryGetValue(table, out var tableId) ? tableId : null,
                    Depth = oe.Ancestors(One + "OE").Count()
                });
            }
            if (Blocks.Sum(b => b.Editable ? b.Text.Length : 0) > options.MaxPageChars || Blocks.Count > 1000)
                throw new AiException("页面内容超过 Agent 限额，请选择较小范围后重试。");
            // 待高亮代码也要发给模型读。放不下时按原来的方式只保护，不让整页失败。
            if (!options.EnableCodeHighlight || Blocks.Sum(b => b.Editable || b.CodeCandidate ? b.Text.Length : 0) > options.MaxPageChars)
                foreach (var b in Blocks.Where(b => b.CodeCandidate)) b.ProtectedReason = "protected_code";
        }

        /// <summary>容器级的保护原因：只支持标题和文本框；图文混排要开关允许，墨迹、附件等对象所在的整个文本框跳过。</summary>
        private string ContainerReason(XElement container)
        {
            string reason = null;
            if (container == null || !(container.Name == One + "Title" || container.Name == One + "Outline")) reason = "unsupported_container";
            if (!Options.EnableMixedOutlines && container != null && container.Descendants().Any(IsBinary)) reason = "mixed_outline_not_verified";
            if (container != null && container.Descendants().Any(e => IsBinary(e) && e.Name != One + "Image")) reason = "unsupported_outline_objects";
            return reason;
        }

        /// <summary>
        /// 登记表格，返回表格元素到短 ID 的对应。选中范围时，表格里有文字的段落都选中了才算在范围内。
        /// 代码框和不支持的文本框里的表格只保护。
        /// </summary>
        private Dictionary<XElement, string> CollectTables(ISet<string> selection)
        {
            var ids = new Dictionary<XElement, string>();
            foreach (var table in Page.Descendants(One + "Table"))
            {
                var objectId = (string)table.Attribute("objectID");
                if (string.IsNullOrEmpty(objectId)) continue;
                if (selection != null)
                {
                    var lines = table.Descendants(One + "OE").Where(e => e.Elements(One + "T").Any() && !string.IsNullOrWhiteSpace(PageEditor.ExtractPlainText(e))).ToList();
                    if (lines.Count == 0 || lines.Any(e => !selection.Contains((string)e.Attribute("objectID") ?? ""))) continue;
                }
                var container = table.Ancestors().FirstOrDefault(e => e.Parent == Page);
                var reason = IsCodeBox(table, Page) ? "highlighted_code" : ContainerReason(container);
                var cells = table.Element(One + "Row")?.Elements(One + "Cell").Select(c =>
                    PageEditor.ExtractPlainText(c.Descendants(One + "OE").FirstOrDefault(e => e.Elements(One + "T").Any()) ?? new XElement(One + "OE")).Trim());
                var summary = reason == "highlighted_code" || cells == null ? null : string.Join(" | ", cells);
                var item = new AgentTable
                {
                    Id = "t" + (Tables.Count + 1), ObjectId = objectId, ProtectedReason = reason,
                    ContainerId = container == null ? "" : (string)container.Attribute("objectID") ?? container.Name.LocalName,
                    Fingerprint = AgentTable.TakeFingerprint(table), Original = TableLook.Read(table), Draft = TableLook.Read(table),
                    Rows = table.Elements(One + "Row").Count(), Columns = table.Element(One + "Columns")?.Elements(One + "Column").Count() ?? 0,
                    Summary = summary == null || summary.Length <= 40 ? summary : summary.Substring(0, 40)
                };
                Tables.Add(item);
                ids[table] = item.Id;
            }
            return ids;
        }

        /// <summary>登记 OneNote 已识别出文字的图片。选中范围时只收图片本身或所在段落被整体选中的。</summary>
        private void CollectImages(ISet<string> selection)
        {
            foreach (var image in Page.Descendants(One + "Image"))
            {
                var text = image.Element(One + "OCRData")?.Element(One + "OCRText")?.Value?.Trim();
                if (string.IsNullOrEmpty(text)) continue;
                if (selection != null && (string)image.Attribute("selected") != "all" && (string)image.Parent?.Attribute("selected") != "all") continue;
                var container = image.Parent == Page ? image : image.Ancestors().FirstOrDefault(e => e.Parent == Page);
                var length = Math.Min(text.Length, AgentImage.MaxChars);
                if (length < text.Length && char.IsHighSurrogate(text[length - 1])) length--;
                Images.Add(new AgentImage
                {
                    Id = "i" + (Images.Count + 1), ObjectId = (string)image.Attribute("objectID") ?? (string)image.Parent?.Attribute("objectID"),
                    ContainerId = container == null ? "" : (string)container.Attribute("objectID") ?? container.Name.LocalName,
                    Text = text.Substring(0, length), Truncated = length < text.Length
                });
            }
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

        /// <summary>草稿页面：结构草稿（默认当前的 <see cref="Layout"/>）加上格式草稿和草稿里的样式、标记定义。段落仍带 <see cref="AgentLayout.Key"/>。</summary>
        internal XElement CreateDraftPage(XElement layout = null)
        {
            var page = new XElement(layout ?? Layout);
            page.Elements(One + "QuickStyleDef").Remove();
            page.Elements(One + "TagDef").Remove();
            page.AddFirst(DraftTags.Elements().Concat(DraftStyles.Elements()).Select(e => new XElement(e)));
            foreach (var b in Blocks.Where(b => b.Changed))
            {
                var target = AgentLayout.Find(page, b.Id);
                if (target != null) CopyFormat(b.Draft, target);
            }
            return page;
        }

        /// <summary>
        /// 结构操作的两条不变量，candidate 是改好的结构草稿：已排入的代码框和表格转换仍然连续、内容不变；
        /// 已有段落的格式不变（挂到带样式的段落下面会继承它的样式，这种调整不做）。
        /// </summary>
        internal void CheckLayout(XElement candidate)
        {
            foreach (var conversion in CodeConversions)
            {
                CodeSelection selection = null;
                try { selection = AgentCode.Select(candidate, conversion.Blocks.Select(b => b.ObjectId).ToList()); } catch (AiException) { }
                if (selection == null || selection.Code != conversion.Code) throw new AiException("这样调整会打断已排入的代码框或表格转换，没有应用。");
            }
            var before = CreateDraftPage();
            var after = CreateDraftPage(candidate);
            foreach (var b in Blocks)
            {
                var left = AgentLayout.Find(before, b.Id);
                var right = AgentLayout.Find(after, b.Id);
                if (left == null || right == null) continue;
                string expected;
                try { expected = SemanticFormat(left, before); } catch (Exception) { continue; }
                string actual;
                try { actual = SemanticFormat(right, after); } catch (Exception) { actual = null; }
                if (expected != actual) throw new AiException($"这样调整会改变段落 {b.Id} 的格式（会继承上级段落或文本框的样式），没有应用。");
            }
        }

        private static bool IsBinary(XElement e) => new[] { "Image", "InkDrawing", "InkWord", "InkParagraph", "InsertedFile", "MediaFile", "FutureObject", "HTMLBlock" }.Contains(e.Name.LocalName);
        /// <summary>
        /// 代码段落分三种：已在代码框里（单格表格、格内全是等宽段落）的保护不动；整段等宽的待转换为代码框；
        /// 只有部分文字等宽的是行内代码，保守保护整段。不是代码返回 null。
        /// </summary>
        private static string CodeKind(XElement oe, AgentRichText rich, XElement page)
        {
            var table = oe.Ancestors(One + "Cell").FirstOrDefault()?.Ancestors(One + "Table").FirstOrDefault();
            if (table != null && IsCodeBox(table, page)) return "highlighted_code";
            var (any, all) = rich.Monospace(page);
            if (all) return "unhighlighted_code";
            if (any || (!string.IsNullOrWhiteSpace(rich.Text) && oe.AncestorsAndSelf(One + "OE").Any(PageEditor.IsCodeParagraph))) return "protected_code";
            return null;
        }
        /// <summary>已有代码框：单行单格的表格，格内全是等宽段落。</summary>
        internal static bool IsCodeBox(XElement table, XElement page)
        {
            var rows = table.Elements(One + "Row").ToList();
            if (rows.Count != 1 || rows[0].Elements(One + "Cell").Count() != 1) return false;
            return rows[0].Element(One + "Cell").Descendants(One + "OE").Where(e => e.Elements(One + "T").Any()).All(e => IsMonospaceParagraph(e, page));
        }

        private static bool IsMonospaceParagraph(XElement oe, XElement page)
        {
            if (PageEditor.IsCodeParagraph(oe)) return true;
            try
            {
                var rich = new AgentRichText(oe);
                return string.IsNullOrWhiteSpace(rich.Text)
                    ? Css.Effective(oe, page).TryGetValue("font-family", out var font) && Css.IsMonospace(font)
                    : rich.Monospace(page).All;
            }
            catch (Exception ex) when (ex is XmlException || ex is AiException || ex is ArgumentException) { return false; }
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
            foreach (var tag in oe.Elements(One + "Tag")) data.Append(AgentMarks.Definition(page, tag));
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
            AgentMarks.CopyMarks(source, target);
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
                NormalizeSpacing(oe, "spaceBefore") + "|" + NormalizeSpacing(oe, "spaceAfter") + "|" + role + "|" + AgentMarks.Projection(oe, page);
        }
        private static string NormalizeSpacing(XElement e, string key) => double.TryParse((string)e.Attribute(key), NumberStyles.Float, CultureInfo.InvariantCulture, out var n)
            ? n.ToString("0.###", CultureInfo.InvariantCulture) : "0";
    }
}
