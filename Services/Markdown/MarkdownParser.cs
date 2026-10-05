using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using OneNoteCodeHelper.Highlighting;

namespace OneNoteCodeHelper.Services.Markdown
{
    /// <summary>
    /// Markdown 文本 → 块。按 CommonMark/GFM 的常见写法逐行解析，只认插入笔记用得上的部分。
    ///
    /// 和 CommonMark 不同的地方：每个源行一段，不把相邻的行合并成一段（合并要在行间补空格，中文里会多出空格）；
    /// 空行连续几个只留一个，开头结尾和列表内部的空行去掉；嵌套列表只要比上一项缩进深就算下一层。
    /// </summary>
    internal static class MarkdownParser
    {
        internal const int MaxChars = 200000;
        internal const int MaxParagraphs = 2000;

        private static readonly Regex FenceOpen = new Regex(@"^(?<mark>`{3,}|~{3,})(?<info>.*)$");
        private static readonly Regex Atx = new Regex(@"^#{1,6}(?=[ \t]|$)");
        private static readonly Regex ClosingHashes = new Regex(@"(?:^|[ \t]+)#+[ \t]*$");
        private static readonly Regex Rule = new Regex(@"^(?:(?:\*[ \t]*){3,}|(?:-[ \t]*){3,}|(?:_[ \t]*){3,})$");
        private static readonly Regex SetextH1 = new Regex(@"^=+[ \t]*$");
        private static readonly Regex SetextH2 = new Regex(@"^-+[ \t]*$");
        private static readonly Regex Quote = new Regex(@"^(?:>[ \t]?)+");
        private static readonly Regex ListItem = new Regex(@"^(?<marker>[-*+]|(?<number>[0-9]{1,9})[.)])(?:[ \t]+|$)(?:\[(?<todo>[ xX])\](?:[ \t]+|$))?");
        private static readonly Regex DelimiterRow = new Regex(@"^\|?[ \t]*:?-+:?[ \t]*(?:\|[ \t]*:?-+:?[ \t]*)*\|?[ \t]*$");
        private static readonly Regex CellSplit = new Regex(@"(?<!\\)\|");
        /// <summary>元数据里的一行：缩进的续行、- 列表、# 注释，或者 key: 值（冒号后面是空白或行尾）。</summary>
        private static readonly Regex YamlLine = new Regex(@"^(?:[ \t]|-(?:[ \t]|$)|#|[^\s:#][^:]*:(?:[ \t]|$))");

        /// <summary>一层列表：列表符号所在的列、这一项的层级和列表种类（bullet、number）。</summary>
        private sealed class ListEntry
        {
            internal int MarkerColumn;
            internal int Level;
            internal string Kind;
        }

        private sealed class State
        {
            internal readonly MarkdownDocument Doc = new MarkdownDocument();
            internal readonly List<ListEntry> Lists = new List<ListEntry>();
            internal bool PendingBlank;
            /// <summary>上一行是列表项或它的续行，没有隔着空行：不缩进的下一行也算这一项的续行。</summary>
            internal bool LazyOpen;

            internal void Add(MarkdownBlock block, bool inList)
            {
                // 列表里的空行不要；列表外连续的空行只留一个，开头的不要。
                if (PendingBlank && !inList && Doc.Blocks.Count > 0)
                    Doc.Blocks.Add(new MarkdownBlock { Kind = MarkdownBlockKind.Blank });
                PendingBlank = false;
                Doc.Blocks.Add(block);
            }
        }

        internal static MarkdownDocument Parse(string text)
        {
            text = text ?? string.Empty;
            if (text.Length > MaxChars)
            {
                return new MarkdownDocument { LimitError = $"内容太长：最多 {MaxChars / 10000} 万字，现在约 {text.Length / 10000.0:0.#} 万字。" };
            }

            var lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            var state = new State();
            for (var i = FrontMatter(lines, state); i < lines.Length;)
            {
                i = ParseLine(lines, i, state);
            }

            var doc = state.Doc;
            var paragraphs = doc.ParagraphCount;
            if (paragraphs > MaxParagraphs)
            {
                doc.LimitError = $"内容太长：最多 {MaxParagraphs} 段（表格单元格和代码行都算），现在 {paragraphs} 段。";
            }

            return doc;
        }

        /// <summary>
        /// 文档开头的 YAML 元数据（Obsidian、Hugo、Jekyll 的笔记常见）：第一个非空行是 ---，到下一个 --- 或 ... 为止，
        /// 中间每行都像 YAML（key: 值、- 列表、# 注释或缩进的续行）。整块写成 YAML 代码框，
        /// 不然两条 --- 会变成空行，最后一行还会和结尾的 --- 凑成 Setext 二级标题。返回接着解析的行号，不是元数据时返回 0。
        /// </summary>
        private static int FrontMatter(string[] lines, State state)
        {
            var start = 0;
            while (start < lines.Length && string.IsNullOrWhiteSpace(lines[start])) start++;
            if (start >= lines.Length || lines[start].TrimEnd() != "---") return 0;

            var end = start + 1;
            while (end < lines.Length && lines[end].TrimEnd() != "---" && lines[end].TrimEnd() != "...") end++;
            if (end >= lines.Length) return 0;

            var body = lines.Skip(start + 1).Take(end - start - 1).ToList();
            if (!body.Any(l => l.Trim().Length > 0) || !body.All(l => l.Trim().Length == 0 || YamlLine.IsMatch(l))) return 0;

            state.Add(new MarkdownBlock { Kind = MarkdownBlockKind.Code, Code = string.Join("\n", body), Language = LanguageRegistry.Find("yaml") }, false);
            return end + 1;
        }

        /// <summary>处理从 lines[i] 开始的一行或几行（围栏、表格），返回下一行的下标。</summary>
        private static int ParseLine(string[] lines, int i, State state)
        {
            var line = lines[i];
            if (string.IsNullOrWhiteSpace(line))
            {
                state.PendingBlank = true;
                state.LazyOpen = false;
                return i + 1;
            }

            var indent = Indent(line, out var contentStart);
            var content = line.Substring(contentStart);

            // 这一行在列表里的位置：比哪一层列表符号缩进深，就挂在那一项下面。
            var parent = state.Lists.LastOrDefault(l => l.MarkerColumn < indent);
            var item = ListItem.Match(content);
            var isRule = Rule.IsMatch(content);

            if (!isRule && item.Success)
            {
                // 退回到这一项所在的那层；退出来的最后一层就是同一层的上一项。
                ListEntry sibling = null;
                while (state.Lists.Count > 0 && state.Lists[state.Lists.Count - 1].MarkerColumn >= indent)
                {
                    sibling = state.Lists[state.Lists.Count - 1];
                    state.Lists.RemoveAt(state.Lists.Count - 1);
                }

                var level = state.Lists.Count == 0 ? 0 : state.Lists[state.Lists.Count - 1].Level + 1;
                var numbered = item.Groups["number"].Success;
                var kind = numbered ? "number" : "bullet";
                var todo = item.Groups["todo"].Success ? (bool?)(item.Groups["todo"].Value != " ") : null;
                // 同一层上一项是同种列表才接着编号（中间隔着空行、下级项也算）；新起的一组记下原文的起始编号，
                // 这样全写成 1. 的列表也按 1、2、3 编号。
                var continues = sibling != null && sibling.Level == level && sibling.Kind == kind;
                var block = new MarkdownBlock
                {
                    Kind = MarkdownBlockKind.ListItem,
                    Level = level,
                    // 「- [ ]」只加待办标记，不再加圆点。
                    ListKind = todo != null && !numbered ? null : kind,
                    Todo = todo,
                    Number = numbered && !continues ? int.Parse(item.Groups["number"].Value, CultureInfo.InvariantCulture) : (int?)null,
                    Inline = MarkdownInline.Parse(HardBreak(TrimEnd(content.Substring(item.Length)), lines, i), state.Doc)
                };
                // 同一组或下一层的列表项前面的空行不要；换了一种列表时保留，看得出是两组。
                state.Add(block, continues || state.Lists.Count > 0);
                state.Lists.Add(new ListEntry { MarkerColumn = indent, Level = level, Kind = kind });
                state.LazyOpen = true;
                return i + 1;
            }

            // 列表项下面的内容：缩进到列表符号右边，或者紧跟着列表项、没有隔空行的续行。
            var listLevel = -1;
            if (state.Lists.Count > 0)
            {
                if (parent != null)
                {
                    while (state.Lists[state.Lists.Count - 1] != parent) state.Lists.RemoveAt(state.Lists.Count - 1);
                    listLevel = parent.Level + 1;
                }
                else if (state.LazyOpen && !state.PendingBlank && !IsBlockStart(content))
                {
                    listLevel = state.Lists[state.Lists.Count - 1].Level + 1;
                }
                else
                {
                    state.Lists.Clear();
                }
            }

            var inside = listLevel >= 0;
            var baseLevel = inside ? listLevel : 0;

            var fence = FenceOpen.Match(content);
            if ((inside || indent <= 3) && fence.Success && !(fence.Groups["mark"].Value[0] == '`' && fence.Groups["info"].Value.IndexOf('`') >= 0))
            {
                return ParseFence(lines, i, indent, fence, baseLevel, inside, state);
            }

            if (!inside && indent >= 4 && CanStartIndentedCode(lines, i, state))
            {
                return ParseIndentedCode(lines, i, state);
            }

            if ((inside || indent <= 3) && content.Contains("|") && i + 1 < lines.Length && IsTableStart(content, lines[i + 1]))
            {
                return ParseTable(lines, i, baseLevel, inside, state);
            }

            if (isRule && (inside || indent <= 3))
            {
                // OneNote 没有分隔线：换成一个空行。
                if (!inside) state.PendingBlank = true;
                state.LazyOpen = false;
                return i + 1;
            }

            var atx = Atx.Match(content);
            if (atx.Success && (inside || indent <= 3))
            {
                var title = content.Substring(atx.Length).Trim();
                var closing = ClosingHashes.Match(title);
                if (closing.Success) title = title.Substring(0, closing.Index).TrimEnd();
                state.Add(new MarkdownBlock
                {
                    Kind = MarkdownBlockKind.Heading,
                    HeadingLevel = atx.Length,
                    Level = baseLevel,
                    Inline = MarkdownInline.Parse(title, state.Doc)
                }, inside);
                state.LazyOpen = false;
                return i + 1;
            }

            var quote = Quote.Match(content);
            if (quote.Success && (inside || indent <= 3))
            {
                var depth = quote.Value.Count(c => c == '>');
                var body = TrimEnd(content.Substring(quote.Length));
                if (body.Trim().Length == 0)
                {
                    // 引用里单独一个 > 是段落间隔。
                    if (!inside) state.PendingBlank = true;
                    return i + 1;
                }

                state.Add(new MarkdownBlock
                {
                    Kind = MarkdownBlockKind.Quote,
                    Level = baseLevel + depth - 1,
                    Inline = MarkdownInline.Parse(HardBreak(body, lines, i), state.Doc)
                }, inside);
                state.LazyOpen = inside && state.LazyOpen;
                return i + 1;
            }

            // Setext 标题：正文下一行全是 = 或 -（列表外）。
            if (!inside && indent <= 3 && i + 1 < lines.Length)
            {
                var next = lines[i + 1];
                var nextIndent = Indent(next, out var nextStart);
                var underline = nextIndent <= 3 ? next.Substring(nextStart) : string.Empty;
                var level = SetextH1.IsMatch(underline) ? 1 : SetextH2.IsMatch(underline) ? 2 : 0;
                if (level > 0)
                {
                    state.Add(new MarkdownBlock
                    {
                        Kind = MarkdownBlockKind.Heading,
                        HeadingLevel = level,
                        Inline = MarkdownInline.Parse(TrimEnd(content), state.Doc)
                    }, false);
                    state.LazyOpen = false;
                    return i + 2;
                }
            }

            state.Add(new MarkdownBlock
            {
                Kind = MarkdownBlockKind.Paragraph,
                Level = baseLevel,
                Inline = MarkdownInline.Parse(HardBreak(TrimEnd(content), lines, i), state.Doc)
            }, inside);
            state.LazyOpen = inside;
            return i + 1;
        }

        /// <summary>
        /// 围栏代码：收到同一种符号、不短于开头的结束行为止，没有结束行就到文末（CommonMark 的规定）。
        /// 内容去掉开头那一行的缩进；语言按围栏上写的名字或别名找。写了但没有对应高亮的（diff、ruby）用纯文本，
        /// 不去猜：自动识别会把 diff 认成 C#、ruby 认成 Lua，猜错比不猜更糟。没写语言才自动识别，再不行用纯文本。
        /// </summary>
        private static int ParseFence(string[] lines, int i, int indent, Match fence, int level, bool inList, State state)
        {
            var mark = fence.Groups["mark"].Value;
            var info = fence.Groups["info"].Value.Trim();
            var close = new Regex("^[ \\t]*" + Regex.Escape(mark[0].ToString()) + "{" + mark.Length + @",}[ \t]*$");
            var end = i + 1;
            while (end < lines.Length && !close.IsMatch(lines[end])) end++;

            var code = string.Join("\n", lines.Skip(i + 1).Take(end - i - 1).Select(l => Dedent(l, indent)));
            // js {1,3}、{.python} 这类写法只取语言名。
            var name = info.TrimStart('{').Split(' ', '\t', '{', '}', ',')[0].TrimStart('.');
            var language = name.Length > 0
                ? LanguageRegistry.FindByName(name) ?? LanguageRegistry.Find("text")
                : LanguageRegistry.Detect(code) ?? LanguageRegistry.Find("text");
            state.Add(new MarkdownBlock { Kind = MarkdownBlockKind.Code, Level = level, Code = code, Language = language }, inList);
            state.LazyOpen = false;
            return Math.Min(end + 1, lines.Length);
        }

        /// <summary>缩进 4 格的代码块：一直到不再缩进 4 格的非空行；末尾的空行不算。</summary>
        private static int ParseIndentedCode(string[] lines, int i, State state)
        {
            var end = i;
            var last = i;
            while (end < lines.Length && (string.IsNullOrWhiteSpace(lines[end]) || Indent(lines[end], out _) >= 4))
            {
                if (!string.IsNullOrWhiteSpace(lines[end])) last = end;
                end++;
            }

            var code = string.Join("\n", lines.Skip(i).Take(last - i + 1).Select(l => Dedent(l, 4)));
            var language = LanguageRegistry.Detect(code) ?? LanguageRegistry.Find("text");
            state.Add(new MarkdownBlock { Kind = MarkdownBlockKind.Code, Code = code, Language = language }, false);
            state.LazyOpen = false;
            return last + 1;
        }

        /// <summary>缩进代码块不能打断正文：前一行要是空行、文档开头，或者前面是标题、代码、表格。</summary>
        private static bool CanStartIndentedCode(string[] lines, int i, State state)
        {
            if (i == 0 || string.IsNullOrWhiteSpace(lines[i - 1])) return true;
            var previous = state.Doc.Blocks.LastOrDefault();
            return previous == null || previous.Kind == MarkdownBlockKind.Heading || previous.Kind == MarkdownBlockKind.Code ||
                   previous.Kind == MarkdownBlockKind.Table;
        }

        private static bool IsTableStart(string header, string delimiter)
        {
            Indent(delimiter, out var start);
            var row = delimiter.Substring(start);
            return row.Contains("|") && DelimiterRow.IsMatch(row) && Cells(header).Count == Cells(row).Count;
        }

        /// <summary>GFM 表格：标题行、分隔行，再往下每个含 | 的非空行一行。列数按最多的那行，缺的格留空。</summary>
        private static int ParseTable(string[] lines, int i, int level, bool inList, State state)
        {
            var alignments = Cells(lines[i + 1]).Select(c =>
            {
                var cell = c.Trim();
                var left = cell.StartsWith(":", StringComparison.Ordinal);
                var right = cell.EndsWith(":", StringComparison.Ordinal);
                return left && right ? "center" : right ? "right" : left ? "left" : null;
            }).ToList();

            var rows = new List<List<List<MarkdownRun>>> { Row(lines[i], state.Doc) };
            var end = i + 2;
            while (end < lines.Length && !string.IsNullOrWhiteSpace(lines[end]) && lines[end].Contains("|"))
            {
                rows.Add(Row(lines[end], state.Doc));
                end++;
            }

            var columns = rows.Max(r => r.Count);
            foreach (var row in rows)
            {
                while (row.Count < columns) row.Add(new List<MarkdownRun>());
            }

            while (alignments.Count < columns) alignments.Add(null);
            state.Add(new MarkdownBlock { Kind = MarkdownBlockKind.Table, Level = level, Rows = rows, Alignments = alignments }, inList);
            state.LazyOpen = false;
            return end;
        }

        private static List<List<MarkdownRun>> Row(string line, MarkdownDocument doc) =>
            Cells(line).Select(c => MarkdownInline.Parse(c.Trim().Replace("\\|", "|"), doc)).ToList();

        /// <summary>按没有转义的 | 切开，去掉行首行尾各一个 |。</summary>
        private static List<string> Cells(string line)
        {
            var trimmed = line.Trim();
            if (trimmed.StartsWith("|", StringComparison.Ordinal)) trimmed = trimmed.Substring(1);
            if (trimmed.EndsWith("|", StringComparison.Ordinal) && !trimmed.EndsWith("\\|", StringComparison.Ordinal))
                trimmed = trimmed.Substring(0, trimmed.Length - 1);
            return CellSplit.Split(trimmed).ToList();
        }

        /// <summary>会另起一块的行：不能当成上一个列表项的续行。</summary>
        private static bool IsBlockStart(string content) =>
            Atx.IsMatch(content) || Quote.IsMatch(content) || FenceOpen.IsMatch(content) || Rule.IsMatch(content);

        /// <summary>行首空白的列数（Tab 按 4 列的制表位），contentStart 是第一个非空白字符的下标。</summary>
        private static int Indent(string line, out int contentStart)
        {
            var column = 0;
            var i = 0;
            for (; i < line.Length; i++)
            {
                if (line[i] == ' ') column++;
                else if (line[i] == '\t') column += 4 - column % 4;
                else break;
            }

            contentStart = i;
            return column;
        }

        /// <summary>去掉最多 columns 列的行首空白，多出来的缩进保留。</summary>
        private static string Dedent(string line, int columns)
        {
            var column = 0;
            var i = 0;
            for (; i < line.Length && column < columns; i++)
            {
                if (line[i] == ' ') column++;
                else if (line[i] == '\t')
                {
                    var width = 4 - column % 4;
                    if (column + width > columns) return new string(' ', column + width - columns) + line.Substring(i + 1);
                    column += width;
                }
                else break;
            }

            return line.Substring(i);
        }

        private static string TrimEnd(string text) => text.TrimEnd(' ', '\t');

        /// <summary>
        /// 行尾单个反斜杠在 CommonMark 里是段内换行。这里每行本来就单独一段，下一行还有字（引用里去掉 &gt; 后看）时去掉它；
        /// 段落最后一行（后面是空行或文末）的反斜杠按规定是原文，留着。连着两个是转义的反斜杠，也留着。
        /// </summary>
        private static string HardBreak(string content, string[] lines, int i)
        {
            var trailing = content.Length - content.TrimEnd('\\').Length;
            if (trailing % 2 == 0 || i + 1 >= lines.Length) return content;
            var next = lines[i + 1].Trim();
            var quote = Quote.Match(next);
            if (next.Substring(quote.Length).Trim().Length == 0) return content;
            return TrimEnd(content.Substring(0, content.Length - 1));
        }

        /// <summary>
        /// 粘贴的内容像不像 Markdown，只用来在插入窗口里提示。有围栏或表格分隔行就算；链接、粗体、## 标题要再加上别的一类；
        /// 单靠 # 注释和 - 列表不算，YAML、Python 也常这么写。
        /// </summary>
        internal static bool LooksLikeMarkdown(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            var strong = new HashSet<string>();
            var weak = new HashSet<string>();
            foreach (var raw in text.Replace("\r\n", "\n").Split('\n').Take(400))
            {
                var line = raw.Trim();
                if (line.Length == 0) continue;
                if (FenceOpen.IsMatch(line)) strong.Add("fence");
                else if (Regex.IsMatch(line, @"^#{2,6}[ \t]+\S")) strong.Add("heading");
                else if (Regex.IsMatch(line, @"^#[ \t]+\S")) weak.Add("heading");
                else if (line.Contains("|") && line.Contains("-") && DelimiterRow.IsMatch(line)) strong.Add("table");
                else if (ListItem.IsMatch(line) && line.Length > 2) weak.Add("list");
                else if (line.StartsWith(">", StringComparison.Ordinal)) weak.Add("quote");
                if (Regex.IsMatch(line, @"\[[^\]\n]+\]\([^)\s]+\)")) strong.Add("link");
                if (Regex.IsMatch(line, @"\*\*[^*\s][^*]*\*\*")) strong.Add("bold");
            }

            return strong.Contains("fence") || strong.Contains("table") || strong.Count > 0 && strong.Count + weak.Count >= 2;
        }
    }
}
