using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using OneNoteCodeHelper.Services.Agent;

namespace OneNoteCodeHelper.Services.Markdown
{
    /// <summary>
    /// 把解析好的 Markdown 写成 OneNote 段落：标题、正文、引用套 Agent 的段落预设，列表和待办用原生列表、标记，
    /// 围栏代码用和「插入代码」相同的代码框，表格用原生表格。整份内容放进页面末尾新建的文本框，不改页面上已有的内容。
    /// </summary>
    internal static class MarkdownWriter
    {
        private static XNamespace One => OneNoteApi.One;

        /// <summary>表格标题行的底色，同 Agent 表格样式里的浅灰。</summary>
        private static string HeaderShading => TableLook.Shadings[1];

        /// <summary>
        /// 生成段落。styles、tags 是页面上 QuickStyleDef、TagDef 的副本：用到的样式、标记定义没有时加进去，
        /// 提交时连同它们一起回传。返回顶层的 one:OE，下级段落已挂在各自上一层的 OEChildren 里。
        /// </summary>
        internal static List<XElement> Write(MarkdownDocument doc, AddInSettings settings, AgentOptions options, XElement styles, XElement tags)
        {
            var top = new List<XElement>();
            // 每一层最近的一段，下一层的块挂在它下面。
            var path = new List<XElement>();
            foreach (var block in doc.Blocks)
            {
                var node = Build(block, settings, options, styles, tags);
                var level = Math.Min(block.Level, path.Count);
                // 代码框、表格下面不挂段落，挂不上就提到同一层。空行有 one:T，照常可以挂：
                // 「> 甲」「>」「>> 乙」里的乙挂在空行下面，才能既排在空行后面、又多缩进一级。
                while (level > 0 && path[level - 1].Element(One + "T") == null) level--;
                if (level == 0)
                {
                    top.Add(node);
                }
                else
                {
                    var parent = path[level - 1];
                    var children = parent.Element(One + "OEChildren");
                    if (children == null) parent.Add(children = new XElement(One + "OEChildren"));
                    children.Add(node);
                }

                path.RemoveRange(level, path.Count - level);
                path.Add(node);
            }

            return top;
        }

        private static XElement Build(MarkdownBlock block, AddInSettings settings, AgentOptions options, XElement styles, XElement tags)
        {
            switch (block.Kind)
            {
                case MarkdownBlockKind.Blank:
                    var blank = AgentLayout.NewBlankLine("md", options, styles);
                    blank.SetAttributeValue(AgentLayout.Key, null);
                    return blank;
                case MarkdownBlockKind.Code:
                    return new XElement(One + "OE", CodeBlockBuilder.BuildTable(block.Code, block.Language, settings.Theme, settings));
                case MarkdownBlockKind.Table:
                    return new XElement(One + "OE", Table(block, settings, options, styles));
                default:
                    var oe = Paragraph(block.Inline, block.Preset, block.BoldText, settings, options, styles);
                    if (block.ListKind != null) AgentMarks.SetList(oe, block.ListKind, block.Number);
                    if (block.Todo != null) AgentMarks.SetTag(oe, "todo", block.Todo, tags);
                    return oe;
            }
        }

        /// <summary>
        /// 先按纯文字套预设，再换成带格式的 HTML：反过来的话套预设会去掉行内代码的字体。
        /// 标题的加粗来自样式定义，和 MCP 导入的做法相同。
        /// </summary>
        private static XElement Paragraph(List<MarkdownRun> runs, string preset, bool bold, AddInSettings settings, AgentOptions options, XElement styles)
        {
            var oe = AgentLayout.NewParagraph("md", MarkdownBlock.PlainText(runs), preset, null, options, styles);
            oe.SetAttributeValue(AgentLayout.Key, null);
            oe.Element(One + "T").ReplaceNodes(new XCData(MarkdownInline.ToHtml(runs, settings.FontFamily, bold)));
            return oe;
        }

        /// <summary>
        /// 原生表格：显示边框，第一行是标题行（加粗、浅灰底）。列宽只是架构要求的占位值，
        /// 未锁定的列宽由 OneNote 按内容重新计算（本机实测，同 Agent 的转表格）。
        /// </summary>
        private static XElement Table(MarkdownBlock block, AddInSettings settings, AgentOptions options, XElement styles)
        {
            var columns = block.Rows.Max(r => r.Count);
            var table = new XElement(One + "Table", new XAttribute("bordersVisible", "true"), new XAttribute("hasHeaderRow", "true"),
                new XElement(One + "Columns", Enumerable.Range(0, columns).Select(i =>
                    new XElement(One + "Column", new XAttribute("index", i), new XAttribute("width", 100)))));
            for (var r = 0; r < block.Rows.Count; r++)
            {
                var row = new XElement(One + "Row");
                for (var c = 0; c < columns; c++)
                {
                    var paragraph = Paragraph(block.Rows[r][c], "body", r == 0, settings, options, styles);
                    var alignment = block.Alignments[c];
                    if (alignment == "center" || alignment == "right") paragraph.SetAttributeValue("alignment", alignment);
                    var cell = new XElement(One + "Cell", new XElement(One + "OEChildren", paragraph));
                    if (r == 0) cell.SetAttributeValue("shadingColor", HeaderShading);
                    row.Add(cell);
                }

                table.Add(row);
            }

            return table;
        }

        /// <summary>
        /// 把 Markdown 插到页面末尾的新文本框里。始终带读取时的修改时间提交；页面在这期间被改过（0x80042010）时
        /// 重新读页面、重新生成再交，最多三次。只新增文本框，重新生成不会覆盖别人的改动。
        /// verified 表示回读时找到了文字一致的新文本框。
        /// </summary>
        internal static EditResult Insert(IOneNotePageAccess api, string pageId, string markdown, AddInSettings settings,
            AgentOptions options, out bool verified)
        {
            verified = false;
            var watch = Stopwatch.StartNew();
            var doc = MarkdownParser.Parse(markdown);
            if (doc.LimitError != null) return EditResult.Fail(doc.LimitError);
            if (doc.Blocks.Count == 0) return EditResult.Fail("没有要插入的内容。");

            string expected = null;
            for (var attempt = 1; ; attempt++)
            {
                var page = XDocument.Parse(api.GetPageContent(pageId, PageInfo.piBasic)).Root;
                if (page == null) return EditResult.Fail("读取当前页面内容失败。");
                DateTime lastModified;
                try { lastModified = AgentPageSnapshot.Modified(page); }
                catch (AiException ex) { return EditResult.Fail(ex.Message); }

                var styles = new XElement("styles", page.Elements(One + "QuickStyleDef").Select(e => new XElement(e)));
                var tags = new XElement("tags", page.Elements(One + "TagDef").Select(e => new XElement(e)));
                var (x, y) = PageEditor.NextFreePosition(page);
                var outline = new XElement(One + "Outline",
                    new XElement(One + "Position",
                        new XAttribute("x", x.ToString("0.#", CultureInfo.InvariantCulture)),
                        new XAttribute("y", y.ToString("0.#", CultureInfo.InvariantCulture)),
                        new XAttribute("z", 0)),
                    new XElement(One + "Size",
                        new XAttribute("width", settings.CodeBlockWidth.ToString("0.#", CultureInfo.InvariantCulture)),
                        new XAttribute("height", "100")),
                    new XElement(One + "OEChildren", Write(doc, settings, options, styles, tags)));
                expected = Projection(outline);

                // TagDef、QuickStyleDef 的编号只在同一份提交 XML 里有效，按架构顺序排在文本框前面。
                var definitions = (outline.Descendants(One + "Tag").Any() ? tags.Elements() : Enumerable.Empty<XElement>())
                    .Concat(styles.Elements());
                var xml = PageEditor.BuildPageChanges(pageId, definitions.Concat(new[] { outline }).ToArray());
                try
                {
                    lock (PageEditCoordinator.ForPage(pageId)) api.UpdatePageContent(xml, lastModified);
                    break;
                }
                catch (COMException ex) when (ex.ErrorCode == unchecked((int)0x80042010))
                {
                    if (attempt < 3) continue;
                    return EditResult.Fail("页面一直在变化，没有插入。请稍后重试。");
                }
                catch (Exception ex)
                {
                    AddInLog.Error("插入 Markdown 失败。pageId=" + pageId, ex);
                    return EditResult.Fail("写回 OneNote 失败：" + ex.Message);
                }
            }

            try
            {
                var saved = XDocument.Parse(api.GetPageContent(pageId, PageInfo.piBasic)).Root;
                verified = saved != null && saved.Elements(One + "Outline").Any(o => Projection(o) == expected);
            }
            catch (Exception ex)
            {
                AddInLog.Warn("插入 Markdown 后回读失败。pageId=" + pageId, ex);
            }

            AddInLog.Info($"插入 Markdown：{doc.TextCount} 段文字，{doc.TableCount} 个表格，{doc.CodeCount} 个代码框，" +
                          $"{doc.WarningCount} 处按原文保留，回读{(verified ? "一致" : "未确认")}，用时 {watch.ElapsedMilliseconds} ms，pageId={pageId}");
            return EditResult.Ok(verified
                ? $"已插入 Markdown：{doc.TextCount} 段文字，{doc.TableCount} 个表格，{doc.CodeCount} 个代码框。"
                : "已提交，但回读时没有找到插入的内容，请检查当前页面。");
        }

        /// <summary>回读核对用：文本框里全部文字按顺序拼起来，去掉所有空白。OneNote 回存会改写空格、硬空格和标签写法，只比文字。</summary>
        private static string Projection(XElement outline) =>
            new string(string.Concat(outline.Descendants(One + "OE").Select(PageEditor.ExtractPlainText))
                .Where(c => !char.IsWhiteSpace(c)).ToArray());
    }
}
