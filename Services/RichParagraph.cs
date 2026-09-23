using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace OneNoteCodeHelper.Services
{
    /// <summary>
    /// 一个 one:OE 段落里的文字，拆成逐个字符，每个字符记着它在哪个 one:T 里、外面包着哪些标签。
    ///
    /// AI 只看得到纯文本、还回来的也是纯文本。要是直接拿新文本整段覆盖，段落里的加粗、颜色、链接就全没了。
    /// 这里把新旧文本逐字符比对：没动的字符原样保留（连同原来的实体写法和标签），新插入的字符
    /// 沿用被替换的那个字符、或者前一个字符的格式，再按标签重新拼回 one:T 的 HTML。
    /// 段落自身的样式、列表符号、缩进都在 OE 上，这里不碰。
    /// </summary>
    internal sealed class RichParagraph
    {
        /// <summary>一层开标签。Path 是从最外层到这一层的完整链，用来比较两个字符的标签是否相同。</summary>
        private sealed class TagScope
        {
            internal TagScope(TagScope parent, string openTag, string name)
            {
                OpenTag = openTag;
                Name = name;
                Path = parent == null
                    ? new List<TagScope> { this }
                    : new List<TagScope>(parent.Path) { this };
                Parent = parent;
            }

            internal TagScope Parent { get; }

            /// <summary>原样的开标签文本，比如 &lt;span style='font-weight:bold'&gt;。</summary>
            internal string OpenTag { get; }

            internal string Name { get; }

            internal List<TagScope> Path { get; }
        }

        private readonly struct RichChar
        {
            internal RichChar(int run, TagScope scope, char text, string html)
            {
                Run = run;
                Scope = scope;
                Text = text;
                Html = html;
            }

            /// <summary>在第几个 one:T 里。</summary>
            internal int Run { get; }

            internal TagScope Scope { get; }

            /// <summary>给 AI 看的字符：&amp;nbsp; 归一成普通空格，&lt;br&gt; 是 \n。</summary>
            internal char Text { get; }

            /// <summary>写回时用的 HTML，没改动的字符保持原来的写法。</summary>
            internal string Html { get; }
        }

        private static readonly HashSet<string> VoidTags = new HashSet<string>(
            new[] { "img", "hr", "wbr", "input", "meta", "link", "col", "area", "base", "embed", "source", "track", "param" },
            System.StringComparer.OrdinalIgnoreCase);

        private readonly List<XElement> _runs;

        private readonly List<RichChar> _chars;

        private RichParagraph(List<XElement> runs, List<RichChar> chars, bool isLossless)
        {
            _runs = runs;
            _chars = chars;
            IsLossless = isLossless;
            Text = new string(chars.Select(c => c.Text).ToArray());
        }

        /// <summary>段落的纯文本。发给 AI 的是它，写回前核对「这段期间有没有被改过」比的也是它。</summary>
        internal string Text { get; }

        /// <summary>
        /// 能不能原样拼回去。碰到注释、img 之类的空标签、对不上的闭标签时为 false，
        /// 这种段落不交给 AI，免得改完丢东西。
        /// </summary>
        internal bool IsLossless { get; }

        private static XNamespace One => OneNoteApi.One;

        internal static RichParagraph Parse(XElement oe)
        {
            var runs = oe.Elements(One + "T").ToList();
            var chars = new List<RichChar>();
            var lossless = true;

            for (var run = 0; run < runs.Count; run++)
            {
                lossless &= ParseRun(runs[run].Value, run, chars);
            }

            return new RichParagraph(runs, chars, lossless);
        }

        /// <summary>
        /// 把段落文字改成 newText，只动有差异的字符。调用前应确认 <see cref="IsLossless"/>。
        /// </summary>
        internal void Apply(string newText)
        {
            var ops = TextDiff.Compute(Text, newText);
            var output = new List<RichChar>(newText.Length);

            // 段首就插入、前面没有字符可沿用格式时，用第一个保留下来的字符的格式。
            var firstKept = ops.Where(op => op.Kind == DiffKind.Keep).Select(op => (RichChar?)_chars[op.OldIndex]).FirstOrDefault()
                            ?? (_chars.Count > 0 ? _chars[0] : new RichChar(0, null, ' ', " "));

            // 每个位置之后第一个保留下来的字符，插入空白时要和前一个字符比一比。
            var nextKept = new RichChar?[ops.Count + 1];
            for (var k = ops.Count - 1; k >= 0; k--)
            {
                nextKept[k] = ops[k].Kind == DiffKind.Keep ? _chars[ops[k].OldIndex] : nextKept[k + 1];
            }

            // 紧挨着的一串删除里的第一个字符。随后的插入就是在「替换」它，沿用它的格式最自然：
            // 把加粗的错字改掉，改出来的字也还是加粗的。
            RichChar? replaced = null;

            for (var k = 0; k < ops.Count; k++)
            {
                var op = ops[k];
                switch (op.Kind)
                {
                    case DiffKind.Keep:
                        output.Add(_chars[op.OldIndex]);
                        replaced = null;
                        break;

                    case DiffKind.Delete:
                        replaced = replaced ?? _chars[op.OldIndex];
                        break;

                    default:
                        var template = replaced
                                       ?? PickTemplate(output.Count > 0 ? output[output.Count - 1] : (RichChar?)null,
                                           nextKept[k], op.NewChar)
                                       ?? firstKept;
                        output.Add(new RichChar(template.Run, template.Scope, op.NewChar, Encode(op.NewChar)));
                        break;
                }
            }

            WriteBack(output);
        }

        /// <summary>
        /// 新插入的字符默认沿用前一个字符的格式，和在编辑器里接着打字一样。
        /// 空白例外：排版优化会在「我用**Git**管理」的 Git 两边加空格，空格要是跟着进了加粗或链接，
        /// 下划线、背景色会多出一截，所以取前后两个字符里标签层数少的那个。
        /// </summary>
        private static RichChar? PickTemplate(RichChar? previous, RichChar? next, char inserted)
        {
            if (previous == null || next == null || !char.IsWhiteSpace(inserted))
            {
                return previous ?? next;
            }

            var previousDepth = previous.Value.Scope?.Path.Count ?? 0;
            var nextDepth = next.Value.Scope?.Path.Count ?? 0;
            return nextDepth < previousDepth ? next : previous;
        }

        private void WriteBack(List<RichChar> output)
        {
            var byRun = output.ToLookup(c => c.Run);
            var emptyRuns = new List<XElement>();

            for (var run = 0; run < _runs.Count; run++)
            {
                var chars = byRun[run].ToList();
                if (chars.Count == 0)
                {
                    emptyRuns.Add(_runs[run]);
                    continue;
                }

                _runs[run].ReplaceNodes(new XCData(Serialize(chars)));
            }

            // 字全删光了的 one:T 去掉，但一个段落至少留一个。
            if (emptyRuns.Count == _runs.Count && emptyRuns.Count > 0)
            {
                emptyRuns[0].ReplaceNodes(new XCData(string.Empty));
                emptyRuns.RemoveAt(0);
            }

            foreach (var run in emptyRuns)
            {
                run.Remove();
            }
        }

        /// <summary>按字符的标签链重新拼 HTML：相邻字符标签相同就不断开，不同时只关、开有差别的那几层。</summary>
        private static string Serialize(List<RichChar> chars)
        {
            var builder = new StringBuilder();
            var open = new List<TagScope>();

            foreach (var ch in chars)
            {
                var path = ch.Scope?.Path ?? new List<TagScope>();

                var common = 0;
                while (common < open.Count && common < path.Count && open[common] == path[common])
                {
                    common++;
                }

                for (var k = open.Count - 1; k >= common; k--)
                {
                    builder.Append("</").Append(open[k].Name).Append('>');
                }

                for (var k = common; k < path.Count; k++)
                {
                    builder.Append(path[k].OpenTag);
                }

                open = path;
                builder.Append(ch.Html);
            }

            for (var k = open.Count - 1; k >= 0; k--)
            {
                builder.Append("</").Append(open[k].Name).Append('>');
            }

            return builder.ToString();
        }

        /// <summary>解析一个 one:T 的 HTML，字符追加到 chars。返回 false 表示里面有拼不回去的东西。</summary>
        private static bool ParseRun(string html, int run, List<RichChar> chars)
        {
            var lossless = true;
            TagScope scope = null;
            var i = 0;

            while (i < html.Length)
            {
                var ch = html[i];

                if (ch == '<')
                {
                    var close = html.IndexOf('>', i);
                    if (close < 0)
                    {
                        return false;
                    }

                    var tag = html.Substring(i, close - i + 1);
                    var inner = tag.Substring(1, tag.Length - 2).Trim();
                    i = close + 1;

                    if (inner.StartsWith("/", System.StringComparison.Ordinal))
                    {
                        var closing = TagName(inner.Substring(1));
                        var match = scope;
                        while (match != null && !string.Equals(match.Name, closing, System.StringComparison.OrdinalIgnoreCase))
                        {
                            match = match.Parent;
                        }

                        // 对不上的闭标签，或者先关了外层：重新拼出来的结构会和原来不一样。
                        if (match == null || match != scope)
                        {
                            lossless = false;
                        }

                        if (match != null)
                        {
                            scope = match.Parent;
                        }

                        continue;
                    }

                    var name = TagName(inner);
                    if (string.Equals(name, "br", System.StringComparison.OrdinalIgnoreCase))
                    {
                        chars.Add(new RichChar(run, scope, '\n', tag));
                        continue;
                    }

                    if (inner.Length == 0 || inner[0] == '!' || inner[0] == '?' || inner.EndsWith("/", System.StringComparison.Ordinal)
                        || VoidTags.Contains(name))
                    {
                        lossless = false;
                        continue;
                    }

                    scope = new TagScope(scope, tag, name);
                    continue;
                }

                if (ch == '&')
                {
                    var semicolon = html.IndexOf(';', i);
                    if (semicolon > i && semicolon - i <= 10)
                    {
                        var entity = html.Substring(i, semicolon - i + 1);
                        var decoded = OneNoteHtmlEncoder.DecodeEntity(entity.Substring(1, entity.Length - 2));

                        if (decoded.Length == 1)
                        {
                            chars.Add(new RichChar(run, scope, Normalize(decoded[0]), entity));
                        }
                        else
                        {
                            // 超出 BMP 的字符（代理对）或者不认识的实体：拆成单个字符各自编码。
                            foreach (var c in decoded)
                            {
                                chars.Add(new RichChar(run, scope, Normalize(c), Encode(c)));
                            }
                        }

                        i = semicolon + 1;
                        continue;
                    }
                }

                chars.Add(new RichChar(run, scope, Normalize(ch), ch.ToString()));
                i++;
            }

            return lossless;
        }

        private static string TagName(string inner)
        {
            var end = 0;
            while (end < inner.Length && !char.IsWhiteSpace(inner[end]) && inner[end] != '/' && inner[end] != '>')
            {
                end++;
            }

            return inner.Substring(0, end);
        }

        /// <summary>HTML 源码里的硬空格、换行、制表符在页面上都显示成空格，给 AI 看的也就是空格。</summary>
        private static char Normalize(char ch)
        {
            return ch == ' ' || ch == '\r' || ch == '\n' || ch == '\t' ? ' ' : ch;
        }

        private static string Encode(char ch)
        {
            switch (ch)
            {
                case '&':
                    return "&amp;";
                case '<':
                    return "&lt;";
                case '>':
                    return "&gt;";
                case '\n':
                    return "<br>";
                case ' ':
                    return "&nbsp;";
                default:
                    return ch.ToString();
            }
        }
    }
}
