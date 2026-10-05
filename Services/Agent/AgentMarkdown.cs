using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace OneNoteCodeHelper.Services.Agent
{
    /// <summary>
    /// strip_markdown 的文字分析：找出段落里要删掉的 Markdown 标记字符，其余的字一个不动。
    /// 真正改段落由 <see cref="AgentTools"/> 按这里给的范围逐处删除，没删的字保留原有格式和链接。
    /// 链接、图片、表格和反斜杠转义不处理：链接不能改，表格交给 text_to_table，转义和 Windows 路径分不清。
    /// </summary>
    internal static class AgentMarkdown
    {
        internal static readonly string[] Kinds = { "heading", "quote", "list", "emphasis", "inline_code", "fence", "rule" };

        internal enum LineRole { Text, Fence, Code }

        /// <summary>一个段落的分析结果。位置都是段落文字（段内换行算一个字）里的下标。</summary>
        internal sealed class Result
        {
            /// <summary>要删掉的标记，互不重叠。</summary>
            internal readonly List<(int Start, int Length)> Marks = new List<(int, int)>();
            /// <summary>emphasis=format 时，去掉符号后要加格式的文字范围（原文下标）和 CSS。</summary>
            internal readonly List<(int Start, int Length, Dictionary<string, string> Css)> Formats = new List<(int, int, Dictionary<string, string>)>();
            /// <summary>第一行原来的标记：标题级别、列表种类、待办是否勾选、是否引用、列表前的缩进（空格数，Tab 算 4 个）。</summary>
            internal int? Heading;
            internal string List;
            /// <summary>原始数字列表的编号，用于恢复独立列表的起点。</summary>
            internal int? Number;
            internal bool? Todo;
            internal bool Quote;
            internal int Indent;
            /// <summary>整段只有一个围栏或分隔线标记，可以整段删掉。</summary>
            internal bool Separator;
            /// <summary>在非 Markdown 围栏里、没有处理的代码行数。</summary>
            internal int CodeLines;
        }

        /// <summary>
        /// 逐行跟踪 ``` 或 ~~~ 围栏，按文本框（或单元格）各用一个。围栏语言是 markdown/md 时里面仍是 Markdown，照常处理，
        /// 这种围栏里带语言的围栏算嵌套的代码块；其他围栏（包括不写语言的）里是代码，不动。
        /// </summary>
        internal sealed class Fences
        {
            private readonly List<(char Mark, int Length, bool Markdown)> _open = new List<(char, int, bool)>();

            internal LineRole Next(string line)
            {
                var m = FenceLine.Match(MatchingText(line));
                var inCode = _open.Count > 0 && !_open[_open.Count - 1].Markdown;
                if (!m.Success || (m.Groups[1].Value[0] == '`' && m.Groups[2].Value.IndexOf('`') >= 0))
                    return inCode ? LineRole.Code : LineRole.Text;
                var mark = m.Groups[1].Value[0];
                var length = m.Groups[1].Length;
                var info = m.Groups[2].Value.Trim();
                if (_open.Count > 0 && info.Length == 0 && _open[_open.Count - 1].Mark == mark && length >= _open[_open.Count - 1].Length)
                {
                    _open.RemoveAt(_open.Count - 1);
                    return LineRole.Fence;
                }
                // 代码块里的围栏样的行是代码内容。
                if (inCode) return LineRole.Code;
                var language = info.Split(' ', '\t')[0].ToLowerInvariant();
                _open.Add((mark, length, language == "markdown" || language == "md"));
                return LineRole.Fence;
            }
        }

        private static readonly Regex FenceLine = new Regex(@"^[ \t]{0,3}(`{3,}|~{3,})(.*)$");
        private static readonly Regex Rule = new Regex(@"^[ \t]{0,3}(?:(?:\*[ \t]*){3,}|(?:-[ \t]*){3,}|(?:_[ \t]*){3,}|={3,}[ \t]*)$");
        private static readonly Regex QuotePrefix = new Regex(@"\G(?:[ \t]{0,3}>[ \t]?)+");
        private static readonly Regex HeadingPrefix = new Regex(@"\G[ \t]{0,3}(#{1,6})(?=[ \t]|$)[ \t]*");
        private static readonly Regex ClosingHashes = new Regex(@"[ \t]+#+[ \t]*$");
        private static readonly Regex ListPrefix = new Regex(@"\G([ \t]*)(?:([-*+])|(?<number>[0-9]{1,9})[.)])[ \t]+(?:\[([ xX])\](?:[ \t]+|$))?");
        private static readonly Regex InlineCode = new Regex(@"(?<!`)(`+)(?!`)(.+?)(?<!`)\1(?!`)");
        /// <summary>仅匹配时把 OneNote 的硬空格当空格，长度和下标不变，写回仍用原文。</summary>
        private static string MatchingText(string text) => text.Replace('\u00a0', ' ');

        /// <param name="roles">每一行的角色，行数与 text 按 '\n' 切开的相同。</param>
        internal static Result Analyze(string text, IList<LineRole> roles, ISet<string> kinds, bool format)
        {
            var result = new Result();
            var lines = text.Split('\n');
            if (roles.Count != lines.Length) throw new ArgumentException("行角色数与段落行数不一致。");
            var offset = 0;
            var separators = 0;
            for (var i = 0; i < lines.Length; i++)
            {
                var line = MatchingText(lines[i]);
                switch (roles[i])
                {
                    case LineRole.Fence:
                        if (kinds.Contains("fence") && line.Length > 0) { result.Marks.Add((offset, line.Length)); separators++; }
                        break;
                    case LineRole.Code:
                        result.CodeLines++;
                        break;
                    default:
                        if (Rule.IsMatch(line))
                        {
                            if (kinds.Contains("rule")) { result.Marks.Add((offset, line.Length)); separators++; }
                        }
                        else Line(line, offset, i == 0, kinds, format, result);
                        break;
                }
                offset += line.Length + 1;
            }
            result.Separator = lines.Length == 1 && separators == 1;
            return result;
        }

        /// <summary>把 marks 从 text 里删掉，得到期望的文字；用来核对改段落的结果。</summary>
        internal static string Remove(string text, IEnumerable<(int Start, int Length)> marks)
        {
            foreach (var (start, length) in marks.OrderByDescending(m => m.Start)) text = text.Remove(start, length);
            return text;
        }

        private static void Line(string line, int offset, bool first, ISet<string> kinds, bool format, Result result)
        {
            // 标记范围里的字换成 \0，后面的规则不再把它们当成标记或内容的边界。
            var masked = line.ToCharArray();
            void Take(int start, int length, string kind)
            {
                for (var i = start; i < start + length; i++) masked[i] = '\0';
                if (kinds.Contains(kind) && length > 0) result.Marks.Add((offset + start, length));
            }
            var p = 0;
            var quote = QuotePrefix.Match(line);
            if (quote.Success && quote.Length > 0)
            {
                Take(0, quote.Length, "quote");
                if (first && kinds.Contains("quote")) result.Quote = true;
                p = quote.Length;
            }
            var heading = HeadingPrefix.Match(line, p);
            if (heading.Success)
            {
                Take(p, heading.Length, "heading");
                if (first && kinds.Contains("heading")) result.Heading = heading.Groups[1].Length;
                p += heading.Length;
                var closing = ClosingHashes.Match(line, p);
                if (closing.Success && closing.Index > p) Take(closing.Index, closing.Length, "heading");
            }
            else
            {
                var item = ListPrefix.Match(line, p);
                if (item.Success)
                {
                    Take(p, item.Length, "list");
                    if (first && kinds.Contains("list"))
                    {
                        result.List = item.Groups[2].Success ? "bullet" : "number";
                        if (item.Groups["number"].Success) result.Number = int.Parse(item.Groups["number"].Value, System.Globalization.CultureInfo.InvariantCulture);
                        if (item.Groups[3].Success) result.Todo = item.Groups[3].Value != " ";
                        result.Indent = item.Groups[1].Value.Sum(c => c == '\t' ? 4 : 1);
                    }
                    p += item.Length;
                }
            }
            for (var i = 0; i < p; i++) masked[i] = '\0';
            // 行内代码优先：里面的星号、下划线是代码，不是强调。
            foreach (Match m in InlineCode.Matches(new string(masked), p))
            {
                var ticks = m.Groups[1].Length;
                Take(m.Index, ticks, "inline_code");
                Take(m.Index + m.Length - ticks, ticks, "inline_code");
                for (var i = m.Index; i < m.Index + m.Length; i++) masked[i] = '\0';
            }
            if (kinds.Contains("emphasis")) Emphasis(masked, p, offset, format, result);
        }

        /// <summary>
        /// 插入 Markdown 用的行内强调分析：masked 是一行行内文字，行内代码、链接等不参与配对的字已换成 \0。
        /// 配对规则与 strip_markdown 相同，结果里只有强调符号（Marks）和要加的格式（Formats）。
        /// </summary>
        internal static Result InlineEmphasis(char[] masked)
        {
            var result = new Result();
            Emphasis(masked, 0, 0, true, result);
            return result;
        }

        private static void Emphasis(char[] masked, int p, int offset, bool format, Result result)
        {
            // 用栈配对，连续的结尾符号可以同时关闭内外层（**粗体里有 *斜体***）。
            var open = "*_~".ToDictionary(c => c, c => new EmphasisOpenings());
            for (var i = p; i < masked.Length;)
            {
                var mark = masked[i];
                if (mark != '*' && mark != '_' && mark != '~') { i++; continue; }
                var end = i + 1;
                while (end < masked.Length && masked[end] == mark) end++;
                var length = end - i;
                if (mark == '~' && length != 2) { i = end; continue; }
                var canOpen = end < masked.Length && !char.IsWhiteSpace(masked[end]) &&
                    (mark == '~' || i == p || !AsciiWord(masked[i - 1]));
                var canClose = i > p && !char.IsWhiteSpace(masked[i - 1]) &&
                    (mark == '~' || end == masked.Length || !AsciiWord(masked[end]));
                var openings = open[mark];
                var match = canClose ? openings.Match(length) : (First: -1, Last: -1);
                if (match.First >= 0)
                {
                    var close = i;
                    var start = openings.Items[match.First].Start;
                    for (var j = match.Last; j >= match.First; j--)
                    {
                        var opening = openings.Items[j];
                        result.Marks.Add((offset + opening.Start, opening.Length));
                        result.Marks.Add((offset + close, opening.Length));
                        if (format) result.Formats.Add((offset + opening.Start + opening.Length,
                            close - opening.Start - opening.Length, Css(new string(mark, opening.Length))));
                        close += opening.Length;
                    }
                    // 被跨过的未配对开启符号仍是正文，但不能再和这个外层之后的符号交叉配对。
                    foreach (var stack in open.Values) stack.RemoveFrom(start);
                }
                else if (canOpen && length <= 3) openings.Add(i, length);
                i = end;
            }
        }

        /// <summary>
        /// 同类开启符号按顺序累积长度，找到恰好消耗闭合串的最近一组；不部分消耗，也不删除未配对符号。
        /// 按闭合长度缓存配对区间，新增只检查新结尾，避免反复向前扫描大量未闭合符号。
        /// </summary>
        private sealed class EmphasisOpenings
        {
            internal readonly List<(int Start, int Length)> Items = new List<(int, int)>();
            private readonly List<int> _sums = new List<int> { 0 };
            private readonly Dictionary<int, int> _indices = new Dictionary<int, int> { [0] = 0 };
            private readonly Dictionary<int, List<(int First, int Last)>> _matches = new Dictionary<int, List<(int, int)>>();

            internal void Add(int start, int length)
            {
                Items.Add((start, length));
                var sum = _sums[_sums.Count - 1] + length;
                _sums.Add(sum); _indices.Add(sum, Items.Count);
                foreach (var pair in _matches)
                    if (_indices.TryGetValue(sum - pair.Key, out var first)) pair.Value.Add((first, Items.Count - 1));
            }

            internal (int First, int Last) Match(int length)
            {
                if (!_matches.TryGetValue(length, out var matches))
                {
                    matches = new List<(int, int)>();
                    for (var last = 0; last < Items.Count; last++)
                        if (_indices.TryGetValue(_sums[last + 1] - length, out var first)) matches.Add((first, last));
                    _matches.Add(length, matches);
                }
                return matches.Count == 0 ? (-1, -1) : matches[matches.Count - 1];
            }

            internal void RemoveFrom(int start)
            {
                while (Items.Count > 0 && Items[Items.Count - 1].Start >= start)
                {
                    _indices.Remove(_sums[_sums.Count - 1]);
                    _sums.RemoveAt(_sums.Count - 1); Items.RemoveAt(Items.Count - 1);
                }
                foreach (var matches in _matches.Values)
                    while (matches.Count > 0 && matches[matches.Count - 1].Last >= Items.Count) matches.RemoveAt(matches.Count - 1);
            }
        }

        private static bool AsciiWord(char c) => c >= 'a' && c <= 'z' || c >= 'A' && c <= 'Z' || c >= '0' && c <= '9';

        private static Dictionary<string, string> Css(string marker)
        {
            if (marker == "~~") return new Dictionary<string, string> { ["text-decoration"] = "line-through" };
            var css = new Dictionary<string, string>();
            if (marker.Length != 2) css["font-style"] = "italic";
            if (marker.Length >= 2) css["font-weight"] = "bold";
            return css;
        }
    }
}
