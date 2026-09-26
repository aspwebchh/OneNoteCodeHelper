using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace OneNoteCodeHelper.Services.Agent
{
    /// <summary>草稿里的一处「连续段落 → 高亮代码框」。提交时在重新读取的页面上重放。</summary>
    internal sealed class AgentCodeConversion
    {
        internal List<AgentBlock> Blocks;
        internal string LanguageId;
        internal string Code;
        internal XElement Table;
    }

    /// <summary>代码框的撤销记录：代码框没被改过时换回原来的段落（段落会得到新的 objectID）。</summary>
    internal sealed class AgentCodeUndoItem
    {
        internal string TableId;
        internal string Fingerprint;
        internal List<XElement> Originals;
        internal List<XElement> Styles;
    }

    /// <summary>
    /// Agent 复用「高亮选中」的 <see cref="CodeSelection"/> 和 <see cref="CodeBlockBuilder"/>，
    /// 但只接受能原样撤销的范围：同一文本块里连续的段落，第一段在最外层，每段的下级段落都在范围内。
    /// 这样替换时不会搬动任何没选中的段落，撤销时把原段落放回代码框的位置即可。
    /// </summary>
    internal static class AgentCode
    {
        private static XNamespace One => OneNoteApi.One;

        internal static CodeSelection Select(XElement page, ICollection<string> objectIds)
        {
            var ids = new HashSet<string>(objectIds);
            var paragraphs = page.Descendants(One + "OE").Where(e => ids.Contains((string)e.Attribute("objectID") ?? "")).ToList();
            if (paragraphs.Count != ids.Count) throw new AiException("找不到目标段落。");
            var block = PageEditor.TextBlockOf(paragraphs[0]);
            if (block == null || block.Name == One + "Title") throw new AiException("页面标题不能转换为代码框。");
            if (paragraphs.Any(oe => PageEditor.TextBlockOf(oe) != block)) throw new AiException("代码段落必须在同一个文本框或表格单元格里。");
            var outline = block.AncestorsAndSelf(One + "Outline").FirstOrDefault();
            if (outline == null || string.IsNullOrEmpty((string)outline.Attribute("objectID"))) throw new AiException("代码段落不在可编辑的文本框里。");
            var lines = block.Descendants(One + "OE").Where(oe => PageEditor.TextBlockOf(oe) == block).ToList();
            if (lines.IndexOf(paragraphs[paragraphs.Count - 1]) - lines.IndexOf(paragraphs[0]) + 1 != paragraphs.Count)
                throw new AiException("代码段落必须连续，中间不能夹着其他段落或对象。");
            var selected = new HashSet<XElement>(paragraphs);
            if (paragraphs.Any(oe => oe.Descendants(One + "OE").Any(child => !selected.Contains(child))))
                throw new AiException("段落的下级段落必须一起转换为代码框。");
            if (paragraphs.Any(oe => oe.Elements(One + "List").Any() || oe.Elements(One + "Tag").Any()))
                throw new AiException("带项目符号、编号或标记的段落不能转换为代码框。");
            var depth = Depth(paragraphs[0], block);
            if (paragraphs.Any(oe => Depth(oe, block) < depth)) throw new AiException("代码的第一段不能比后面的段落缩进更深。");
            var selection = new CodeSelection(paragraphs, block, outline);
            if (string.IsNullOrWhiteSpace(selection.Code)) throw new AiException("目标段落里没有代码文字。");
            return selection;
        }

        private static int Depth(XElement oe, XElement block) => oe.Ancestors().TakeWhile(e => e != block).Count(e => e.Name == One + "OE");

        /// <summary>替换前记下最外层的原段落（连同下级段落）和它们引用的样式定义。</summary>
        internal static AgentCodeUndoItem CaptureOriginals(CodeSelection selection, XElement page)
        {
            var selected = new HashSet<XElement>(selection.Paragraphs);
            var tops = selection.Paragraphs.Where(oe => !selected.Contains(oe.Ancestors(One + "OE").FirstOrDefault())).Select(oe => new XElement(oe)).ToList();
            var indexes = new HashSet<string>(tops.SelectMany(t => t.DescendantsAndSelf()).Select(e => (string)e.Attribute("quickStyleIndex")).Where(i => i != null));
            return new AgentCodeUndoItem
            {
                Originals = tops,
                Styles = page.Elements(One + "QuickStyleDef").Where(d => indexes.Contains((string)d.Attribute("index"))).Select(d => new XElement(d)).ToList()
            };
        }

        /// <summary>按 Table 的 objectID 找代码框所在的段落。OneNote 回存时会换掉外层段落的 ID，Table 的 ID 不变。</summary>
        internal static XElement FindCodeBox(XElement page, string tableId)
        {
            var wrapper = page.Descendants(One + "Table").FirstOrDefault(t => (string)t.Attribute("objectID") == tableId)?.Parent;
            return wrapper != null && wrapper.Name == One + "OE" && !wrapper.Elements(One + "T").Any() ? wrapper : null;
        }

        /// <summary>把代码框换回原段落。去掉 ID 和编辑记录让 OneNote 当新段落建，样式定义按当前页面重新对应。</summary>
        internal static List<XElement> Restore(XElement wrapper, AgentCodeUndoItem item, XElement page)
        {
            var map = item.Styles.ToDictionary(d => (string)d.Attribute("index"), d => ParagraphStyles.EnsureDefinition(page, d));
            var restored = item.Originals.Select(o => new XElement(o)).ToList();
            foreach (var e in restored.SelectMany(r => r.DescendantsAndSelf()))
            {
                foreach (var a in e.Attributes().Where(a => a.Name.LocalName == "objectID" || a.Name.LocalName == "selected" ||
                    a.Name.LocalName == "creationTime" || a.Name.LocalName.StartsWith("lastModified", StringComparison.Ordinal) ||
                    a.Name.LocalName.StartsWith("author", StringComparison.Ordinal)).ToList()) a.Remove();
                var index = (string)e.Attribute("quickStyleIndex");
                if (index != null && map.TryGetValue(index, out var mapped)) e.SetAttributeValue("quickStyleIndex", mapped);
            }
            wrapper.ReplaceWith(restored);
            return restored;
        }

        /// <summary>代码框的指纹：Table ID、底色和各行文字。不含列宽和外层段落 ID，这两样 OneNote 会自己改。</summary>
        internal static string Fingerprint(XElement wrapper)
        {
            var data = new StringBuilder((string)wrapper.Element(One + "Table")?.Attribute("objectID") ?? "");
            foreach (var cell in wrapper.Descendants(One + "Cell")) data.Append('|').Append(((string)cell.Attribute("shadingColor") ?? "").ToLowerInvariant());
            foreach (var oe in wrapper.Descendants(One + "OE").Where(e => e.Elements(One + "T").Any())) data.Append('\n').Append(PlainText(oe));
            using (var sha = SHA256.Create()) return Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(data.ToString())));
        }

        /// <summary>新建段落的回读核验只比文字：OneNote 会改写代码行的 span 和硬空格。</summary>
        internal static string PlainText(XElement oe) => PageEditor.ExtractPlainText(oe).Replace(' ', ' ').TrimEnd();
    }
}
