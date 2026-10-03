using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace OneNoteCodeHelper.Services.Agent
{
    /// <summary>标题整容器的写入计划。草稿文字和格式仍共用 Layout、AgentBlock.Draft。</summary>
    internal sealed class AgentTitleEdit
    {
        internal XElement Before;
        internal string Fingerprint;
        internal bool Restore;
        internal XElement Draft;
        internal List<XElement> Styles = new List<XElement>();
        internal List<XElement> Tags = new List<XElement>();
    }

    internal sealed class AgentTitleUndoItem
    {
        internal XElement Before;
        internal string AfterFingerprint;
        internal List<XElement> Styles = new List<XElement>();
        internal List<XElement> Tags = new List<XElement>();
    }

    internal static class AgentPageTitle
    {
        private static XNamespace One => OneNoteApi.One;

        internal static string Text(XElement title)
        {
            var text = title == null ? "" : string.Join("\n", title.Elements(One + "OE").Select(oe => RichParagraph.Parse(oe).Text));
            return string.IsNullOrWhiteSpace(text) ? "" : text;
        }

        internal static void Validate(XElement page)
        {
            var titles = page.Elements(One + "Title").ToList();
            if (titles.Count > 1) throw new AiException("页面标题结构异常，没有修改。");
            var title = titles.SingleOrDefault();
            if (title == null) return;
            var lines = title.Elements(One + "OE").ToList();
            if (lines.Count > 1 || title.Descendants(One + "OE").Count() != lines.Count ||
                title.Descendants().Any(AgentPageSnapshot.IsBinary) ||
                lines.Any(oe => oe.Elements().Any(e => e.Name != One + "T" && e.Name != One + "Meta")))
                throw new AiException("页面标题包含不支持的内容，没有修改。");
            foreach (var oe in lines)
            {
                var id = (string)oe.Attribute("objectID");
                if (id != null && page.Descendants(One + "OE").Count(e => (string)e.Attribute("objectID") == id) != 1)
                    throw new AiException("页面标题段落 ID 重复，没有修改。");
                if (!RichParagraph.Parse(oe).IsLossless) throw new AiException("页面标题包含不支持的 HTML，没有修改。");
            }
        }

        internal static XElement Empty() => new XElement(One + "Title", new XElement(One + "OE", new XElement(One + "T", new XCData(""))));

        internal static void Put(XElement page, XElement title)
        {
            var current = page.Element(One + "Title");
            if (current != null) current.ReplaceWith(title);
            else
            {
                var following = page.Elements().FirstOrDefault(e => e.Name == One + "Outline" || e.Name == One + "Image" || AgentPageSnapshot.IsBinary(e));
                if (following != null) following.AddBeforeSelf(title); else page.Add(title);
            }
        }

        /// <summary>只比较标题及其引用的样式，不让正文编辑、选区和时间戳引起标题冲突。</summary>
        internal static string Fingerprint(XElement page)
        {
            var title = page.Element(One + "Title");
            var data = new StringBuilder(title == null ? "missing-title" : Stable(title));
            if (title != null)
            {
                foreach (var index in title.DescendantsAndSelf().Attributes("quickStyleIndex").Select(a => a.Value).Distinct().OrderBy(x => x))
                    data.Append('|').Append(Stable(page.Elements(One + "QuickStyleDef").FirstOrDefault(d => (string)d.Attribute("index") == index)));
                foreach (var tag in title.Descendants(One + "Tag")) data.Append('|').Append(Stable(AgentMarks.Definition(page, tag)));
            }
            using (var sha = SHA256.Create()) return Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(data.ToString())));
        }

        private static string Stable(XElement element)
        {
            if (element == null) return "missing";
            var copy = new XElement(element);
            AgentLayout.Strip(copy);
            foreach (var a in copy.DescendantsAndSelf().Attributes().Where(a => a.Name.LocalName == "selected" || a.Name.LocalName == "lastModifiedTime").ToList()) a.Remove();
            foreach (var t in copy.DescendantNodes().OfType<XText>().Where(t => t.Parent.HasElements && string.IsNullOrWhiteSpace(t.Value)).ToList()) t.Remove();
            foreach (var e in copy.DescendantsAndSelf()) e.ReplaceAttributes(e.Attributes().OrderBy(a => a.Name.ToString(), StringComparer.Ordinal).ToArray());
            return copy.ToString(SaveOptions.DisableFormatting);
        }

        internal static void Remap(XElement title, XElement page, IEnumerable<XElement> styles, IEnumerable<XElement> tags)
        {
            var definitions = styles.ToDictionary(d => (string)d.Attribute("index"));
            foreach (var a in title.DescendantsAndSelf().Attributes("quickStyleIndex").ToList())
                if (definitions.TryGetValue(a.Value, out var definition)) a.Value = ParagraphStyles.EnsureDefinition(page, definition);
            var tagDefinitions = tags.ToDictionary(d => (string)d.Attribute("index"));
            foreach (var tag in title.Descendants(One + "Tag"))
                if (tagDefinitions.TryGetValue((string)tag.Attribute("index"), out var definition)) tag.SetAttributeValue("index", AgentMarks.EnsureTagDefinition(page, definition));
        }

        internal static bool Verified(XElement expected, XElement actual, bool preserveFormat)
        {
            Validate(actual);
            var wanted = Text(expected.Element(One + "Title"));
            var written = actual.Element(One + "Title");
            if (Text(written) != wanted || (wanted.Length > 0 && (string)actual.Attribute("name") != wanted)) return false;
            // Office 可以把空标题回存成没有 OE/T 的占位容器，或者省略整个 Title。
            if (wanted.Length == 0)
            {
                var beforeLine = expected.Element(One + "Title")?.Element(One + "OE");
                var afterLine = written?.Element(One + "OE");
                return !preserveFormat || beforeLine?.Elements(One + "T").Any() != true || afterLine?.Elements(One + "T").Any() != true ||
                    AgentPageSnapshot.BlankFormat(beforeLine, expected) == AgentPageSnapshot.BlankFormat(afterLine, actual);
            }
            return !preserveFormat || AgentPageSnapshot.TitleFormat(expected.Element(One + "Title").Element(One + "OE"), expected) ==
                AgentPageSnapshot.TitleFormat(written.Element(One + "OE"), actual);
        }
    }
}
