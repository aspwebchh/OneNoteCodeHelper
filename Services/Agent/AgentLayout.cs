using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace OneNoteCodeHelper.Services.Agent
{
    /// <summary>
    /// 结构草稿里的一次改动，按短 ID 记录。Kind 为 removed、moved、indented、inserted、merged、unwrapped。
    /// 跨文本框的移动和合并记下源文本框 From，提交时两个文本框一起写入、一起跳过。
    /// </summary>
    internal sealed class AgentLayoutChange
    {
        internal string Kind;
        internal string OutlineId;
        internal string From;
        internal string[] Ids;
        /// <summary>unwrapped：拆开的代码框的短 ID（t1…），Ids 是拆出来的新段落 u1…。</summary>
        internal string TableId;
        /// <summary>unwrapped：拆框时删除的链接数，对应整组代码行核验通过后才计入结果。</summary>
        internal int LinksRemoved;
    }

    /// <summary>insert_blocks 新插入的段落，短 ID 为 n1、n2…</summary>
    internal sealed class AgentInserted
    {
        internal string Id;
        internal string OutlineId;
        internal string Text;
        /// <summary>unwrap_code 拆出来的段落（u1…）：只在概况里列出，不登记在 <see cref="AgentPageSnapshot.Inserted"/>，也不接受其他工具。</summary>
        internal bool Unwrapped;
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
        /// <summary>一起写入的文本框（跨框移动、合并）同一个值，撤销时整组恢复或整组跳过。</summary>
        internal string Group;
        /// <summary>文本框已被合并删掉，撤销时按 Before 重新建立（得到新的 ID）。</summary>
        internal bool Deleted;
        /// <summary>已核验的文字编辑计数，整框撤销成功后恢复到结果。</summary>
        internal int MarkdownMarks;
        internal int LinksRemoved;
        internal List<string> TextFixes = new List<string>();
    }

    /// <summary>
    /// 结构草稿（<see cref="AgentPageSnapshot.Layout"/>）上的操作：删空行、调整缩进、移动和插入段落、合并文本框。
    /// 草稿里的段落用私有属性 <see cref="Key"/> 记短 ID（p1…、新插入的 n1…），不依赖 objectID，提交前统一去掉。
    /// 本机实测（OneNote 16.0.20326.20158）：整框回传时，同一文本框里移动、挂到别的段落下、提到上一级的段落都保留 objectID；
    /// 省略的段落被删除；无 ID 的段落和表格按新对象建立。移到另一个文本框的对象一律按新对象建立（带着原 ID 也一样），
    /// 图片只带 CallbackID 会建成坏图，必须带 one:Data；Outline 不能没有段落，写成只剩一行空白时 OneNote 会直接删掉这个文本框。
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
            foreach (var a in root.DescendantsAndSelf().Attributes().Where(a => a.Name.Namespace == Key.Namespace).ToList()) a.Remove();
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

        /// <summary>合并文本框：把源文本框的全部顶层段落（含表格、图片、空行）按顺序移到目标段落前后，再删掉源文本框。</summary>
        internal static void Merge(XElement source, XElement target, bool before)
        {
            var nodes = source.Elements(One + "OEChildren").Elements(One + "OE").ToList();
            foreach (var oe in nodes) oe.Remove();
            if (before) target.AddBeforeSelf(nodes); else target.AddAfterSelf(nodes);
            source.Remove();
        }

        /// <summary>页面上每个带 objectID 的对象所在的顶层容器（文本框、标题）的 ID。</summary>
        internal static Dictionary<string, string> Homes(XElement page)
        {
            var homes = new Dictionary<string, string>();
            foreach (var container in page.Elements())
                foreach (var id in container.DescendantsAndSelf().Attributes("objectID").Select(a => a.Value))
                    if (!homes.ContainsKey(id)) homes[id] = (string)container.Attribute("objectID") ?? container.Name.LocalName;
            return homes;
        }

        /// <summary>
        /// 要写入 outline 这个文本框的对象里，原来不在这个框里的（跨框移动过来的）去掉 ID，让 OneNote 按新对象建立。
        /// 返回去掉了 ID 的段落原来的 ID，用来给其中的图片补上数据。
        /// </summary>
        internal static Dictionary<XElement, string> StripForeign(XElement outline, string outlineId, Dictionary<string, string> homes)
        {
            var stripped = new Dictionary<XElement, string>();
            foreach (var e in outline.Descendants().Where(e => e.Attribute("objectID") != null).ToList())
            {
                var id = (string)e.Attribute("objectID");
                if (homes.TryGetValue(id, out var home) && home == outlineId) continue;
                AgentCode.StripIdentity(e);
                if (e.Name == One + "OE") stripped[e] = id;
            }
            return stripped;
        }

        /// <summary>
        /// 给图片补上 one:Data（取自 piBinaryData 读到的页面，按所在段落的 ID 对应），已经有数据的跳过。
        /// 给了 originalIds 时只补去掉了 ID 的段落里的图片，ID 取原来的；否则补全部图片。有要补的图片才读取 binary。
        /// </summary>
        internal static void FillImageData(XElement outline, Func<XElement> binary, IDictionary<XElement, string> originalIds = null)
        {
            foreach (var image in outline.Descendants(One + "Image").Where(i => i.Element(One + "Data") == null).ToList())
            {
                var parent = image.Parent;
                string id;
                if (originalIds == null) id = (string)parent?.Attribute("objectID");
                else if (parent == null || !originalIds.TryGetValue(parent, out id)) continue;
                var data = id == null ? null : binary().Descendants(One + "OE").FirstOrDefault(e => (string)e.Attribute("objectID") == id)?.Element(One + "Image")?.Element(One + "Data");
                if (data == null) throw new AiException("读不到要移动的图片数据，已阻止写入。");
                image.Add(new XElement(data));
            }
        }

        /// <summary>写入前整理图片：所在段落保留原 ID 的只带 CallbackID，重建的段落里的图片只带数据。</summary>
        internal static void PrepareImages(XElement outline)
        {
            foreach (var image in outline.Descendants(One + "Image"))
            {
                if (image.Parent?.Attribute("objectID") != null || image.Element(One + "Data") == null) image.Elements(One + "Data").Remove();
                else image.Elements(One + "CallbackID").Remove();
            }
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
        /// 新补的空行（insert_blocks 的 blank 项、代码框间隔补入）：只取正文字体外观，不引用标题、列表或原生样式的段间距，用 &amp;nbsp; 占位。
        /// styles 只用来算外观，不往里加定义。
        /// </summary>
        internal static XElement NewBlankLine(string id, AgentOptions options, XElement styles)
        {
            var blank = NewParagraph(id, "\u00a0", "body", null, options, new XElement(styles));
            blank.Attribute("quickStyleIndex")?.Remove();
            var appearance = Css.Read((string)blank.Attribute("style"));
            Css.Merge(appearance, Css.Emphasis(ParagraphStyles.Definition("body", options)));
            blank.SetAttributeValue("style", Css.Write(appearance));
            blank.SetAttributeValue("spaceBefore", "0"); blank.SetAttributeValue("spaceAfter", "0");
            return blank;
        }

        /// <summary>
        /// unwrap_code：把代码框（外层 OE）换成正文段落，每个代码行一段。代码行复制后去掉 ID、清除行内格式、套正文预设，
        /// 文字、硬空格和段内换行原样保留；代码框留白用的段前段后间距一并去掉。调用方已校验代码框只有平铺的文字段落。
        /// </summary>
        internal static List<XElement> Unwrap(XElement wrapper, Func<string> nextId, AgentOptions options, XElement styles, bool removeLinks = true)
        {
            var lines = wrapper.Element(One + "Table").Element(One + "Row").Element(One + "Cell").Element(One + "OEChildren").Elements(One + "OE").Select(line =>
            {
                var oe = new XElement(line);
                new AgentRichText(oe).ClearInline(!removeLinks, line);
                AgentCode.StripIdentity(oe);
                foreach (var a in oe.Attributes().Where(a => a.Name.Namespace == Key.Namespace || new[] { "style", "quickStyleIndex", "spaceBefore", "spaceAfter" }.Contains(a.Name.LocalName)).ToList())
                    a.Remove();
                ParagraphStyles.Apply(oe, "body", new Dictionary<string, object>(), options, false);
                if (options.EnableNativeHeadings) oe.SetAttributeValue("quickStyleIndex", ParagraphStyles.EnsureDefinition(styles, ParagraphStyles.Definition("body", options)));
                ParagraphStyles.ResetBodyText(oe, options);
                oe.SetAttributeValue(Key, nextId());
                return oe;
            }).ToList();
            wrapper.ReplaceWith(lines);
            return lines;
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
        internal static void Detach(XElement oe)
        {
            var parent = oe.Parent;
            oe.Remove();
            if (parent != null && parent.Name == One + "OEChildren" && parent.Parent?.Name == One + "OE" && !parent.Elements(One + "OE").Any()) parent.Remove();
        }
    }
}
