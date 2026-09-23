using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OneNoteCodeHelper.Services
{
    /// <summary>
    /// 删多余的空行：连续的空行只留一行，文本框（表格单元格）开头、结尾的空行删掉。
    ///
    /// OneNote 里按一次回车就是一个新段落（one:OE），空行就是没有字的段落；Shift+回车是段内换行（&lt;br&gt;）。
    /// 两种都要管：前者是删段落，AI 按约定不能动段落，所以由插件直接改页面 XML，不经过 AI；
    /// 后者只是段内文字，和 AI 的改动合在一起写回。
    /// </summary>
    internal static class BlankLines
    {
        private static XNamespace One => OneNoteApi.One;

        /// <summary>段内换行：去掉开头、结尾的空行，中间连续的空行只留第一行。</summary>
        internal static string CollapseInText(string text)
        {
            if (string.IsNullOrEmpty(text) || text.IndexOf('\n') < 0)
            {
                return text;
            }

            var lines = text.Split('\n');
            var first = 0;
            while (first < lines.Length && string.IsNullOrWhiteSpace(lines[first]))
            {
                first++;
            }

            // 整段都是空白：这是个空行段落，归 RemoveFromPage 管。
            if (first == lines.Length)
            {
                return text;
            }

            var last = lines.Length - 1;
            while (string.IsNullOrWhiteSpace(lines[last]))
            {
                last--;
            }

            var kept = new List<string>();
            for (var i = first; i <= last; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i]) && string.IsNullOrWhiteSpace(lines[i - 1]))
                {
                    continue;
                }

                kept.Add(lines[i]);
            }

            return string.Join("\n", kept);
        }

        /// <summary>
        /// 删掉页面上多余的空行段落，返回删了几行，改过的文本框（one:Outline）加进 changed。
        /// canRemove 决定哪一行空行可以删（只处理选中文字时只删选区里的）；删不了的空行照样算进「连着几行」里。
        /// </summary>
        internal static int RemoveFromPage(XElement page, Func<XElement, bool> canRemove, ISet<XElement> changed)
        {
            var removed = 0;

            foreach (var flow in FlowsOf(page))
            {
                var count = RemoveFromFlow(flow, canRemove);
                if (count > 0)
                {
                    removed += count;
                    changed.Add(flow.Ancestors(One + "Outline").First());
                }
            }

            return removed;
        }

        /// <summary>
        /// 页面上各自独立的一摞摞段落：每个文本框一摞，表格的每个单元格一摞。
        /// 分开算，免得把上一格末尾和下一格开头的空行当成连在一起的。标题不算。
        /// </summary>
        internal static List<XElement> FlowsOf(XElement page)
        {
            var outlines = page.Elements(One + "Outline").ToList();

            return outlines.Elements(One + "OEChildren")
                .Concat(outlines.Descendants(One + "Cell").Elements(One + "OEChildren"))
                .ToList();
        }

        /// <summary>
        /// 空行：只有空白文字的段落。带待办标记、项目符号、编号、图片、表格、下级段落的都不算，
        /// 那些在页面上看得见；代码行（等宽字体）也不算，代码里的空行是有意义的。
        /// </summary>
        internal static bool IsBlankLine(XElement oe)
        {
            if (!oe.Elements(One + "T").Any()
                || oe.Elements().Any(e => e.Name != One + "T" && e.Name != One + "Meta")
                || PageEditor.IsCodeParagraph(oe))
            {
                return false;
            }

            var rich = RichParagraph.Parse(oe);
            return rich.IsLossless && string.IsNullOrWhiteSpace(rich.Text);
        }

        private static int RemoveFromFlow(XElement flow, Func<XElement, bool> canRemove)
        {
            var lines = LinesOf(flow).ToList();
            var blank = lines.Select(IsBlankLine).ToList();
            var doomed = new List<XElement>();

            var i = 0;
            while (i < lines.Count)
            {
                if (!blank[i])
                {
                    i++;
                    continue;
                }

                var start = i;
                while (i < lines.Count && blank[i])
                {
                    i++;
                }

                // 夹在两行内容之间的留一行；文本框开头、结尾的不留；整摞都是空行时留一行。
                var atStart = start == 0;
                var atEnd = i == lines.Count;
                var keep = (atStart || atEnd) && !(atStart && atEnd) ? 0 : 1;

                // 删不了的空行先顶上「要留的那一行」，都能删时留第一行。
                var run = lines.GetRange(start, i - start);
                var keepRemovable = Math.Max(0, keep - run.Count(oe => !canRemove(oe)));

                foreach (var oe in run.Where(canRemove))
                {
                    if (keepRemovable > 0)
                    {
                        keepRemovable--;
                        continue;
                    }

                    doomed.Add(oe);
                }
            }

            foreach (var oe in doomed)
            {
                var parent = oe.Parent;
                oe.Remove();

                // 缩进的下级段落删光了，空的 OEChildren 也去掉。
                if (parent != null && parent != flow && !parent.Elements(One + "OE").Any())
                {
                    parent.Remove();
                }
            }

            return doomed.Count;
        }

        /// <summary>按页面上从上到下的顺序列出一摞里的段落，含缩进的下级段落，不进表格。</summary>
        private static IEnumerable<XElement> LinesOf(XElement oeChildren)
        {
            foreach (var oe in oeChildren.Elements(One + "OE"))
            {
                yield return oe;

                foreach (var line in oe.Elements(One + "OEChildren").SelectMany(LinesOf))
                {
                    yield return line;
                }
            }
        }
    }
}
