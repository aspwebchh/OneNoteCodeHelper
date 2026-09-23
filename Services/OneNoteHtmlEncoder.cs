using System.Collections.Generic;
using System.Text;
using OneNoteCodeHelper.Highlighting;
using OneNoteCodeHelper.Highlighting.Themes;

namespace OneNoteCodeHelper.Services
{
    /// <summary>把 token 流渲染成 OneNote one:T 里那一小撮受支持的 HTML。</summary>
    internal static class OneNoteHtmlEncoder
    {
        /// <summary>一行里的一段同类文本。</summary>
        private readonly struct Segment
        {
            internal Segment(TokenKind kind, string text)
            {
                Kind = kind;
                Text = text;
            }

            internal TokenKind Kind { get; }

            internal string Text { get; }
        }

        /// <summary>
        /// 渲染成逐行的 HTML 片段，每行对应 OneNote 里的一个 one:OE。
        /// 按行拆而不是用 &lt;br&gt;，是因为 OneNote 对段落的处理比对 br 可靠得多。
        /// </summary>
        internal static IReadOnlyList<string> BuildLines(
            string source, IEnumerable<Token> tokens, CodeTheme theme, int tabWidth)
        {
            var result = new List<string>();

            foreach (var line in SplitIntoLines(source ?? string.Empty, tokens))
            {
                result.Add(RenderLine(line, theme, tabWidth));
            }

            return result;
        }

        /// <summary>
        /// 按换行把 token 切成一行行的片段。块注释和 Lua 长字符串这类 token 会跨行，
        /// 必须在这里拆开，否则没法逐行包 span。\r\n、\n、\r 三种换行都要认。
        /// </summary>
        private static List<List<Segment>> SplitIntoLines(string source, IEnumerable<Token> tokens)
        {
            var lines = new List<List<Segment>>();
            var current = new List<Segment>();

            foreach (var token in tokens)
            {
                var segmentStart = token.Start;

                for (var i = token.Start; i < token.End; i++)
                {
                    var ch = source[i];
                    if (ch != '\n' && ch != '\r')
                    {
                        continue;
                    }

                    if (i > segmentStart)
                    {
                        current.Add(new Segment(token.Kind, source.Substring(segmentStart, i - segmentStart)));
                    }

                    lines.Add(current);
                    current = new List<Segment>();

                    if (ch == '\r' && i + 1 < source.Length && source[i + 1] == '\n')
                    {
                        i++;
                    }

                    segmentStart = i + 1;
                }

                if (token.End > segmentStart)
                {
                    current.Add(new Segment(token.Kind, source.Substring(segmentStart, token.End - segmentStart)));
                }
            }

            lines.Add(current);
            return lines;
        }

        private static string RenderLine(List<Segment> segments, CodeTheme theme, int tabWidth)
        {
            var column = 0;
            var atLineStart = true;

            // 先算出每段的 CSS 与 HTML，再把相邻的同款样式合并。
            // 不合并的话每个 token 都要包一层 span，一行几十个字符能膨胀到一两 KB，
            // 几百行代码就会把页面 XML 撑得 OneNote 都处理不动。
            var parts = new List<KeyValuePair<string, string>>();

            foreach (var segment in segments)
            {
                var expanded = ExpandTabs(segment.Text, ref column, tabWidth);
                var html = EncodeText(expanded, ref atLineStart);
                if (html.Length == 0)
                {
                    continue;
                }

                // 纯空白没有字形，不必上色；默认色的文字在浅色主题下也不必包 span。
                // 缩进和标点占了输出的大头，这两条省下来体积差一倍。
                var css = IsBlankHtml(html) || theme.IsOmittable(segment.Kind)
                    ? null
                    : BuildCss(theme.StyleFor(segment.Kind));

                if (parts.Count > 0 && parts[parts.Count - 1].Key == css)
                {
                    parts[parts.Count - 1] = new KeyValuePair<string, string>(
                        css, parts[parts.Count - 1].Value + html);
                }
                else
                {
                    parts.Add(new KeyValuePair<string, string>(css, html));
                }
            }

            var builder = new StringBuilder();
            foreach (var part in parts)
            {
                if (part.Key == null)
                {
                    builder.Append(part.Value);
                }
                else
                {
                    builder.Append("<span style='").Append(part.Key).Append("'>")
                           .Append(part.Value).Append("</span>");
                }
            }

            // 空行必须给个硬空格占位，否则 OneNote 会把这一段直接塌掉，代码的空行就没了。
            return builder.Length == 0 ? "&nbsp;" : builder.ToString();
        }

        /// <summary>判断一段已编码的 HTML 是不是只有空白（普通空格与硬空格）。</summary>
        private static bool IsBlankHtml(string html)
        {
            var i = 0;
            while (i < html.Length)
            {
                if (html[i] == ' ')
                {
                    i++;
                    continue;
                }

                if (string.CompareOrdinal(html, i, "&nbsp;", 0, 6) == 0)
                {
                    i += 6;
                    continue;
                }

                return false;
            }

            return true;
        }

        /// <summary>
        /// 把制表符按真实制表位展开成空格。column 跨片段累加，
        /// 所以一行里前面有多少字符会正确影响后面制表符的宽度。
        /// </summary>
        private static string ExpandTabs(string text, ref int column, int tabWidth)
        {
            if (text.IndexOf('\t') < 0)
            {
                column += text.Length;
                return text;
            }

            var builder = new StringBuilder(text.Length + tabWidth);
            foreach (var ch in text)
            {
                if (ch == '\t')
                {
                    var spaces = tabWidth - (column % tabWidth);
                    builder.Append(' ', spaces);
                    column += spaces;
                }
                else
                {
                    builder.Append(ch);
                    column++;
                }
            }

            return builder.ToString();
        }

        /// <summary>
        /// 转义 + 空白处理。
        ///
        /// 两个关键点：
        /// 1. one:T 虽然是 CDATA，但 OneNote 会把里面的内容当 HTML 解析，所以 &amp; &lt; &gt;
        ///    仍然必须转义——否则 Java 泛型 List&lt;String&gt; 会被当成标签整段吞掉。
        /// 2. HTML 会折叠连续空白：行首缩进全部转成 &amp;nbsp;，行内连续空格保留 n-1 个硬空格
        ///    再跟一个普通空格，这样既保住对齐又给长行留了换行机会。
        /// </summary>
        private static string EncodeText(string text, ref bool atLineStart)
        {
            var builder = new StringBuilder(text.Length);
            var i = 0;

            while (i < text.Length)
            {
                var ch = text[i];

                if (ch == ' ')
                {
                    var run = 0;
                    while (i + run < text.Length && text[i + run] == ' ')
                    {
                        run++;
                    }

                    if (atLineStart)
                    {
                        Repeat(builder, "&nbsp;", run);
                    }
                    else if (run >= 2)
                    {
                        Repeat(builder, "&nbsp;", run - 1);
                        builder.Append(' ');
                    }
                    else
                    {
                        builder.Append(' ');
                    }

                    i += run;
                    continue;
                }

                atLineStart = false;

                switch (ch)
                {
                    case '&':
                        builder.Append("&amp;");
                        break;
                    case '<':
                        builder.Append("&lt;");
                        break;
                    case '>':
                        builder.Append("&gt;");
                        break;
                    default:
                        builder.Append(ch);
                        break;
                }

                i++;
            }

            return builder.ToString();
        }

        private static void Repeat(StringBuilder builder, string text, int count)
        {
            for (var i = 0; i < count; i++)
            {
                builder.Append(text);
            }
        }

        private static string BuildCss(TokenStyle style)
        {
            var builder = new StringBuilder("color:", 48).Append(style.Color);

            if (style.Bold)
            {
                builder.Append(";font-weight:bold");
            }

            if (style.Italic)
            {
                builder.Append(";font-style:italic");
            }

            return builder.ToString();
        }

        /// <summary>
        /// 从 one:T 的内容里还原纯文本，用于读回 OneNote 上已有的代码。
        /// 只需处理我们自己可能写出去的那点标签和实体，外加 OneNote 自己会插的 br。
        /// </summary>
        internal static string DecodeToPlainText(string html)
        {
            if (string.IsNullOrEmpty(html))
            {
                return string.Empty;
            }

            var builder = new StringBuilder(html.Length);
            var i = 0;

            while (i < html.Length)
            {
                var ch = html[i];

                if (ch == '<')
                {
                    var close = html.IndexOf('>', i);
                    if (close < 0)
                    {
                        break;
                    }

                    // <br> 当换行，其余标签（span/b/i/font 等）直接丢掉只留文字
                    var tag = html.Substring(i + 1, close - i - 1).TrimStart('/').TrimEnd('/').Trim();
                    if (tag.StartsWith("br", System.StringComparison.OrdinalIgnoreCase))
                    {
                        builder.Append('\n');
                    }

                    i = close + 1;
                    continue;
                }

                if (ch == '&')
                {
                    var semicolon = html.IndexOf(';', i);
                    if (semicolon > i && semicolon - i <= 10)
                    {
                        var entity = html.Substring(i + 1, semicolon - i - 1);
                        builder.Append(DecodeEntity(entity));
                        i = semicolon + 1;
                        continue;
                    }
                }

                builder.Append(ch);
                i++;
            }

            return builder.ToString();
        }

        private static string DecodeEntity(string entity)
        {
            switch (entity.ToLowerInvariant())
            {
                case "nbsp":
                    return " ";
                case "amp":
                    return "&";
                case "lt":
                    return "<";
                case "gt":
                    return ">";
                case "quot":
                    return "\"";
                case "apos":
                    return "'";
                default:
                    if (entity.StartsWith("#", System.StringComparison.Ordinal))
                    {
                        var digits = entity.Substring(1);
                        var isHex = digits.StartsWith("x", System.StringComparison.OrdinalIgnoreCase);
                        var parsed = isHex
                            ? int.TryParse(digits.Substring(1), System.Globalization.NumberStyles.HexNumber,
                                System.Globalization.CultureInfo.InvariantCulture, out var hex) ? hex : -1
                            : int.TryParse(digits, out var dec) ? dec : -1;

                        if (parsed > 0)
                        {
                            return char.ConvertFromUtf32(parsed);
                        }
                    }

                    // 不认识的实体原样留着，总比吃掉强
                    return "&" + entity + ";";
            }
        }
    }
}
