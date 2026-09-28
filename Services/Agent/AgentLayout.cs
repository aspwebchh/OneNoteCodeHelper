using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace OneNoteCodeHelper.Services.Agent
{
    /// <summary>结构草稿里的一次改动，按短 ID 记录。Kind 为 removed、moved、indented、inserted。</summary>
    internal sealed class AgentLayoutChange
    {
        internal string Kind;
        internal string OutlineId;
        internal string[] Ids;
    }

    /// <summary>insert_blocks 新插入的段落，短 ID 为 n1、n2…</summary>
    internal sealed class AgentInserted
    {
        internal string Id;
        internal string OutlineId;
        internal string Text;
    }

    /// <summary>整框撤销记录：文本框写入后没被改过时，换回写入前的样子。</summary>
    internal sealed class AgentOutlineUndoItem
    {
        internal string OutlineId;
        internal XElement Before;
        /// <summary>Before 引用的样式和标记定义，恢复时按当前页面重新对应编号。</summary>
        internal List<XElement> Styles;
        internal List<XElement> Tags;
        internal string AfterFingerprint;
    }

    /// <summary>
    /// 结构草稿（<see cref="AgentPageSnapshot.Layout"/>）上的操作：删空行、调整缩进、移动和插入段落。
    /// 草稿里的段落用私有属性 <see cref="Key"/> 记短 ID（p1…、新插入的 n1…），不依赖 objectID，提交前统一去掉。
    /// 本机实测（OneNote 16.0.20326.20158）：整框回传时，移动、挂到别的段落下、提到上一级的段落都保留 objectID；
    /// 省略的段落被删除；无 ID 的段落和表格按新对象建立。
    /// </summary>
    internal static class AgentLayout
    {
        private static XNamespace One => OneNoteApi.One;
        internal static readonly XName Key = XNamespace.Get("urn:onenote-code-helper:agent") + "id";

        internal static XElement Find(XElement layout, string id) => layout.Descendants(One + "OE").FirstOrDefault(e => (string)e.Attribute(Key) == id);
        internal static string KeyOf(XElement oe) => (string)oe?.Attribute(Key);

        /// <summary>去掉结构草稿的私有属性，得到能交给 OneNote 的 XML。</summary>
        internal static void Strip(XElement root)
        {
            foreach (var a in root.DescendantsAndSelf().Attributes(Key).ToList()) a.Remove();
        }

        /// <summary>
        /// 删空行，返回删掉的段落。collapse 同「排版优化」：连续空行留一行，文本框（表格单元格）首尾的空行删掉；
        /// all 删掉全部，但一摞段落不会被删空。canRemove 决定哪些空行在范围内。
        /// </summary>
        internal static List<XElement> RemoveBlankLines(XElement layout, Func<XElement, bool> canRemove, bool all)
        {
            var candidates = layout.Descendants(One + "OE").Where(e => BlankLines.IsBlankLine(e) && canRemove(e)).ToList();
            var removable = new HashSet<XElement>(candidates);
            if (!all) BlankLines.RemoveFromPage(layout, canRemove, new HashSet<XElement>());
            else
                foreach (var flow in BlankLines.FlowsOf(layout))
                {
                    var lines = LinesOf(flow).ToList();
                    var doomed = lines.Where(removable.Contains).ToList();
                    foreach (var oe in doomed.Skip(doomed.Count == lines.Count ? 1 : 0)) Detach(oe);
                }
            return candidates.Where(e => e.Parent == null).ToList();
        }

        /// <summary>把段落（连同下级段落）按原来的先后顺序移到目标段落前面或后面，成为目标的兄弟段落。</summary>
        internal static void Move(IList<XElement> nodes, XElement target, bool before)
        {
            foreach (var oe in nodes) Detach(oe);
            if (before) target.AddBeforeSelf(nodes); else target.AddAfterSelf(nodes);
        }

        /// <summary>
        /// in：挂到前一个兄弟段落的下级末尾；out：移到上级段落之后，原来排在它后面的兄弟段落改挂到它下面。
        /// 两种都不改变段落在页面上的上下顺序，只改层级。
        /// </summary>
        internal static void Indent(XElement oe, bool deeper, string id)
        {
            if (deeper)
            {
                var previous = oe.ElementsBeforeSelf(One + "OE").LastOrDefault();
                if (previous == null || !previous.Elements(One + "T").Any()) throw new AiException($"段落 {id} 前面没有可以作为上级的文字段落。");
                Detach(oe);
                Children(previous).Add(oe);
                return;
            }
            var parent = oe.Parent?.Parent;
            if (parent == null || parent.Name != One + "OE") throw new AiException($"段落 {id} 已经在最外层。");
            var following = oe.ElementsAfterSelf(One + "OE").ToList();
            foreach (var f in following) f.Remove();
            if (following.Count > 0) Children(oe).Add(following);
            Detach(oe);
            parent.AddAfterSelf(oe);
        }

        /// <summary>insert_blocks 的新段落：纯文字按 HTML 转义，套用预设样式，可带列表。</summary>
        internal static XElement NewParagraph(string id, string text, string preset, string list, AgentOptions options, XElement styles)
        {
            var oe = new XElement(One + "OE", new XAttribute(Key, id), new XElement(One + "T", new XCData(OneNoteHtmlEncoder.EncodePlainText(text))));
            ParagraphStyles.Apply(oe, preset, new Dictionary<string, object>(), options);
            if (options.EnableNativeHeadings) oe.SetAttributeValue("quickStyleIndex", ParagraphStyles.EnsureDefinition(styles, ParagraphStyles.Definition(preset, options)));
            if (list != null) AgentMarks.SetList(oe, list);
            return oe;
        }

        /// <summary>
        /// 整框指纹：文本框的全部 XML（不含选中状态、修改时间和 OneNote 后台补上的图片识别文字），加上框内引用的样式和标记定义。
        /// 提交时比较，处理期间文本框里有任何改动都不整框替换；撤销时比较写入后的指纹。
        /// </summary>
        internal static string OutlineFingerprint(XElement outline, XElement page)
        {
            var copy = new XElement(outline);
            copy.Descendants(One + "OCRData").Remove();
            foreach (var a in copy.DescendantsAndSelf().Attributes().Where(a => a.Name.LocalName == "selected" || a.Name.LocalName == "lastModifiedTime").ToList()) a.Remove();
            var data = new StringBuilder(copy.ToString(SaveOptions.DisableFormatting));
            foreach (var d in Styles(outline, page).Concat(Tags(outline, page))) data.Append('|').Append(d.ToString(SaveOptions.DisableFormatting));
            using (var sha = SHA256.Create()) return Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(data.ToString())));
        }

        /// <summary>框内引用的 QuickStyleDef 副本。</summary>
        internal static List<XElement> Styles(XElement outline, XElement page)
        {
            var indexes = new HashSet<string>(outline.DescendantsAndSelf().Attributes("quickStyleIndex").Select(a => a.Value));
            return page.Elements(One + "QuickStyleDef").Where(d => indexes.Contains((string)d.Attribute("index"))).Select(d => new XElement(d)).ToList();
        }

        /// <summary>框内标记引用的 TagDef 副本。</summary>
        internal static List<XElement> Tags(XElement outline, XElement page)
        {
            var indexes = new HashSet<string>(outline.Descendants(One + "Tag").Select(t => (string)t.Attribute("index")));
            return page.Elements(One + "TagDef").Where(d => indexes.Contains((string)d.Attribute("index"))).Select(d => new XElement(d)).ToList();
        }

        /// <summary>
        /// OneNote 不会因为 OE 里少了 List 就去掉列表（本机实测），原来有列表、要写的没有时让它把这一段当新段落重建。
        /// current 是页面上现有的对象，用来判断原来有没有列表。
        /// </summary>
        internal static void RebuildDroppedLists(XElement written, XElement current)
        {
            var listed = new HashSet<string>(current.Descendants(One + "OE").Where(e => e.Element(One + "List") != null)
                .Select(e => (string)e.Attribute("objectID")).Where(id => id != null));
            foreach (var oe in written.DescendantsAndSelf(One + "OE").Where(e => e.Element(One + "List") == null && listed.Contains((string)e.Attribute("objectID") ?? "")).ToList())
                AgentCode.StripIdentity(oe);
        }

        /// <summary>页面上从上到下的一摞段落，含缩进的下级段落，不进表格。</summary>
        private static IEnumerable<XElement> LinesOf(XElement children)
        {
            foreach (var oe in children.Elements(One + "OE"))
            {
                yield return oe;
                foreach (var line in oe.Elements(One + "OEChildren").SelectMany(LinesOf)) yield return line;
            }
        }

        private static XElement Children(XElement oe)
        {
            var children = oe.Element(One + "OEChildren");
            if (children == null) oe.Add(children = new XElement(One + "OEChildren"));
            return children;
        }

        /// <summary>从原位置取下段落；缩进的下级段落取光了，空的 OEChildren 也去掉。</summary>
        private static void Detach(XElement oe)
        {
            var parent = oe.Parent;
            oe.Remove();
            if (parent != null && parent.Name == One + "OEChildren" && parent.Parent?.Name == One + "OE" && !parent.Elements(One + "OE").Any()) parent.Remove();
        }
    }
}
