using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OneNoteCodeHelper.Highlighting.Languages
{
    /// <summary>
    /// YAML 词法着色。YAML 没有关键字，难点是分清键和值：标量后面紧跟「冒号 + 空白」就是键，否则是值，
    /// 值再按字面量、数字、字符串细分。块标量（| 和 &gt;）的正文靠缩进界定，要等当前行结束后才开始，
    /// 处理方式和 Bash 的 heredoc 类似。
    /// </summary>
    internal sealed class YamlLanguage : ILanguage
    {
        /// <summary>YAML 1.1 的布尔与空值写法。1.2 只认 true/false/null，但 yes/no/on/off 在配置文件里仍然常见。</summary>
        private static readonly HashSet<string> Literals = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "true", "false", "yes", "no", "on", "off", "null", "~"
        };

        /// <summary>整数、小数、指数、0x/0o、.inf/.nan。1.2.3 这种版本号和 12:30 这种时间都不算数字。</summary>
        private static readonly Regex NumberPattern = new Regex(
            @"^[-+]?([0-9][0-9_]*(\.[0-9_]*)?([eE][-+]?[0-9]+)?|\.[0-9][0-9_]*([eE][-+]?[0-9]+)?|0x[0-9a-fA-F_]+|0o[0-7_]+|\.(inf|Inf|INF))$" +
            @"|^\.(nan|NaN|NAN)$");

        /// <summary>flow 集合（[a, b]、{a: 1}）里会打断 plain scalar 的字符。</summary>
        private const string FlowIndicators = ",[]{}";

        public string Id => "yaml";

        public string DisplayName => "YAML";

        public IEnumerable<Token> Tokenize(string source)
        {
            var c = new LexerCursor(source);
            var lineStart = 0;
            var flowDepth = 0;

            // 当前位置能不能是键的开头：行首缩进之后、"- " 之后、flow 集合里的 { [ , 之后
            var nodeStart = true;

            // 本行的「属主」所在列：行首缩进，遇到键或 "- " 时更新为它们的列。块标量正文必须比它缩进更深
            var ownerIndent = 0;

            // 本行出现了块标量头（| 或 >）时记下当时的 ownerIndent，换行后开始吃正文
            int? pendingBlockScalar = null;

            while (!c.AtEnd)
            {
                var start = c.Position;
                var ch = c.Current;

                if (ch == '\n' || ch == '\r')
                {
                    c.Advance(ch == '\r' && c.Peek() == '\n' ? 2 : 1);
                    c.Emit(TokenKind.Plain, start);

                    if (pendingBlockScalar.HasValue)
                    {
                        ScanBlockScalarBody(c, pendingBlockScalar.Value);
                        pendingBlockScalar = null;
                    }

                    lineStart = c.Position;
                    ownerIndent = 0;
                    if (flowDepth == 0)
                    {
                        nodeStart = true;
                    }

                    continue;
                }

                if (ch == ' ' || ch == '\t')
                {
                    c.SkipWhile(x => x == ' ' || x == '\t');
                    c.Emit(TokenKind.Plain, start);
                    if (start == lineStart)
                    {
                        ownerIndent = c.Position - lineStart;
                    }

                    continue;
                }

                // # 只有在行首或空白之后才是注释：a#b、url#anchor 都不是
                if (ch == '#' && (start == lineStart || char.IsWhiteSpace(source[start - 1])))
                {
                    c.SkipToLineEnd();
                    c.Emit(TokenKind.Comment, start);
                    continue;
                }

                if (start == lineStart)
                {
                    // 文档分隔符 --- 与文档结束符 ...
                    if ((c.Matches("---") || c.Matches("...")) && IsBlankOrEnd(c.Peek(3)))
                    {
                        c.Advance(3);
                        c.Emit(TokenKind.Keyword, start);
                        flowDepth = 0;
                        nodeStart = true;
                        ownerIndent = -1;
                        continue;
                    }

                    // %YAML 1.2、%TAG 指令
                    if (ch == '%')
                    {
                        c.SkipToLineEnd();
                        c.Emit(TokenKind.Annotation, start);
                        continue;
                    }
                }

                // 序列项 "- " 与复杂键 "? "。之后仍是节点开头：- name: x 里的 name 是键
                if ((ch == '-' || ch == '?') && nodeStart && IsBlankOrEnd(c.Peek()))
                {
                    c.Advance();
                    c.Emit(TokenKind.Punctuation, start);
                    ownerIndent = start - lineStart;
                    continue;
                }

                if (ch == ':' && IsValueIndicator(c.Peek(), flowDepth))
                {
                    c.Advance();
                    c.Emit(TokenKind.Punctuation, start);
                    nodeStart = false;
                    continue;
                }

                if (ch == '[' || ch == '{')
                {
                    c.Advance();
                    c.Emit(TokenKind.Punctuation, start);
                    flowDepth++;
                    nodeStart = true;
                    continue;
                }

                if (ch == ']' || ch == '}')
                {
                    c.Advance();
                    c.Emit(TokenKind.Punctuation, start);
                    flowDepth = Math.Max(0, flowDepth - 1);
                    nodeStart = false;
                    continue;
                }

                if (ch == ',' && flowDepth > 0)
                {
                    c.Advance();
                    c.Emit(TokenKind.Punctuation, start);
                    nodeStart = true;
                    continue;
                }

                // &anchor 定义锚点，*alias 引用锚点。锚点后面还可能跟键，所以只有别名会结束节点开头
                if ((ch == '&' || ch == '*') && !IsBlankOrEnd(c.Peek()))
                {
                    c.Advance();
                    c.SkipWhile(x => !char.IsWhiteSpace(x) && FlowIndicators.IndexOf(x) < 0);
                    c.Emit(TokenKind.Variable, start);
                    nodeStart &= ch == '&';
                    continue;
                }

                // 标签：!custom、!!str、!<tag:yaml.org,2002:str>
                if (ch == '!')
                {
                    c.Advance();
                    if (c.Current == '<')
                    {
                        c.SkipWhile(x => x != '>' && x != '\n' && x != '\r');
                        if (c.Current == '>')
                        {
                            c.Advance();
                        }
                    }
                    else
                    {
                        c.SkipWhile(x => !char.IsWhiteSpace(x) && FlowIndicators.IndexOf(x) < 0);
                    }

                    c.Emit(TokenKind.Type, start);
                    continue;
                }

                // 块标量头：|、>，可带保留/截断指示符和缩进数字，如 |-、>+、|2
                if ((ch == '|' || ch == '>') && flowDepth == 0)
                {
                    c.Advance();
                    c.SkipWhile(x => x == '-' || x == '+' || char.IsDigit(x));
                    c.Emit(TokenKind.Operator, start);
                    pendingBlockScalar = ownerIndent;
                    nodeStart = false;
                    continue;
                }

                if (ch == '"' || ch == '\'')
                {
                    ScanQuoted(c, ch);
                    var quotedColon = FindKeyColon(source, c.Position, flowDepth, true);
                    if (quotedColon >= 0)
                    {
                        EmitKey(c, start, quotedColon);
                        ownerIndent = start - lineStart;
                    }
                    else
                    {
                        c.Emit(TokenKind.String, start);
                    }

                    nodeStart = false;
                    continue;
                }

                // 剩下的都是 plain scalar：不带引号的键或值
                var end = FindPlainScalarEnd(source, start, flowDepth);
                c.Advance(end - start);
                var colon = FindKeyColon(source, end, flowDepth, false);
                if (colon >= 0)
                {
                    EmitKey(c, start, colon);
                    ownerIndent = start - lineStart;
                }
                else
                {
                    c.Emit(ClassifyValue(source.Substring(start, end - start)), start);
                }

                nodeStart = false;
            }

            return c.Finish();
        }

        /// <summary>把 [start, 当前位置) 标成键，再吃掉后面的冒号。冒号在 flow 里可以紧跟值（"a":1），必须在这里一并处理。</summary>
        private static void EmitKey(LexerCursor c, int start, int colon)
        {
            c.Emit(TokenKind.Attribute, start);
            c.Advance(colon - c.Position);
            c.Advance();
            c.Emit(TokenKind.Punctuation, colon);
        }

        /// <summary>
        /// plain scalar 的终点（不含末尾空白）：到行尾、键分隔符「: 」、行尾注释「 #」为止，
        /// flow 集合里还会被 , [ ] { } 打断。
        /// </summary>
        private static int FindPlainScalarEnd(string source, int position, int flowDepth)
        {
            var end = position + 1;
            for (var i = position; i < source.Length; i++)
            {
                var ch = source[i];
                if (ch == '\n' || ch == '\r')
                {
                    break;
                }

                if (i > position)
                {
                    if (ch == ':' && IsValueIndicator(i + 1 < source.Length ? source[i + 1] : '\0', flowDepth))
                    {
                        break;
                    }

                    if (ch == '#' && (source[i - 1] == ' ' || source[i - 1] == '\t'))
                    {
                        break;
                    }

                    if (flowDepth > 0 && FlowIndicators.IndexOf(ch) >= 0)
                    {
                        break;
                    }
                }

                if (ch != ' ' && ch != '\t')
                {
                    end = i + 1;
                }
            }

            return end;
        }

        /// <summary>
        /// 从 position 跳过同一行的空白，看是不是键分隔符，是就返回冒号的位置，否则 -1。
        /// 带引号的键在 flow 里可以和值挨着（{"a":1}），plain 键不行（http://x 里的冒号不是分隔符）。
        /// </summary>
        private static int FindKeyColon(string source, int position, int flowDepth, bool quoted)
        {
            var i = position;
            while (i < source.Length && (source[i] == ' ' || source[i] == '\t'))
            {
                i++;
            }

            if (i >= source.Length || source[i] != ':')
            {
                return -1;
            }

            var next = i + 1 < source.Length ? source[i + 1] : '\0';
            return IsValueIndicator(next, flowDepth) || (quoted && flowDepth > 0) ? i : -1;
        }

        /// <summary>冒号后面跟的是这个字符时，冒号才是键值分隔符。</summary>
        private static bool IsValueIndicator(char next, int flowDepth)
        {
            return IsBlankOrEnd(next) || (flowDepth > 0 && FlowIndicators.IndexOf(next) >= 0);
        }

        private static bool IsBlankOrEnd(char ch)
        {
            return ch == '\0' || ch == ' ' || ch == '\t' || ch == '\n' || ch == '\r';
        }

        /// <summary>单行引号字符串。双引号里 \ 是转义符，单引号里 '' 表示一个引号。遇到换行就停。</summary>
        private static void ScanQuoted(LexerCursor c, char quote)
        {
            c.Advance();
            while (!c.AtEnd && c.Current != '\n' && c.Current != '\r')
            {
                if (quote == '"' && c.Current == '\\')
                {
                    // 行尾的 \ 不能把换行也吃进来，否则下一行的行首状态就丢了
                    c.Advance(c.Peek() == '\n' || c.Peek() == '\r' ? 1 : 2);
                    continue;
                }

                if (c.Current == quote)
                {
                    if (quote == '\'' && c.Peek() == '\'')
                    {
                        c.Advance(2);
                        continue;
                    }

                    c.Advance();
                    return;
                }

                c.Advance();
            }
        }

        /// <summary>
        /// 块标量正文（调用时位于块标量头的下一行行首）：空行照吃，缩进比 ownerIndent 深的行照吃，
        /// 碰到第一条缩进不够的非空行、或第 0 列的文档分隔符就停。末尾的空行和最后一行的换行留给主循环。
        /// </summary>
        private static void ScanBlockScalarBody(LexerCursor c, int ownerIndent)
        {
            var source = c.Source;
            var start = c.Position;
            var end = start;
            var i = start;

            while (i < source.Length)
            {
                var textStart = i;
                while (textStart < source.Length && (source[textStart] == ' ' || source[textStart] == '\t'))
                {
                    textStart++;
                }

                var lineEnd = textStart;
                while (lineEnd < source.Length && source[lineEnd] != '\n' && source[lineEnd] != '\r')
                {
                    lineEnd++;
                }

                if (textStart < lineEnd)
                {
                    if (textStart - i <= ownerIndent || (textStart == i && IsDocumentMarker(source, i)))
                    {
                        break;
                    }

                    end = lineEnd;
                }

                if (lineEnd >= source.Length)
                {
                    break;
                }

                i = lineEnd + (source[lineEnd] == '\r' && lineEnd + 1 < source.Length && source[lineEnd + 1] == '\n' ? 2 : 1);
            }

            c.Advance(end - start);
            c.Emit(TokenKind.String, start);
        }

        private static bool IsDocumentMarker(string source, int position)
        {
            if (position + 3 > source.Length)
            {
                return false;
            }

            var marker = source.Substring(position, 3);
            return (marker == "---" || marker == "...")
                   && (position + 3 == source.Length || char.IsWhiteSpace(source[position + 3]));
        }

        private static TokenKind ClassifyValue(string text)
        {
            if (Literals.Contains(text))
            {
                return TokenKind.Literal;
            }

            return NumberPattern.IsMatch(text) ? TokenKind.Number : TokenKind.String;
        }

        public int ScoreLikelihood(string source)
        {
            // 以 { [ 开头的归 JSON，以 < 开头的归 XML/HTML
            if (LikelihoodPatterns.IsMatch(source, @"\A\s*[\[{<]"))
            {
                return 0;
            }

            var score = 0;

            // 文档分隔符
            score += 3 * Count(source, @"^---[^\S\r\n]*\r?$");

            // - name: x，序列里的映射，YAML 独有的写法
            score += 3 * Count(source, @"^[^\S\r\n]*-[^\S\r\n]+[A-Za-z_""'][\w.""'-]*:(\s|$)");

            // spec: 这种独占一行的块开头。排除 Python 的 else:/try:、C++ 的 public:、switch 的 default:
            score += 2 * Count(source,
                @"^[^\S\r\n]*(?!(else|try|finally|except|default|public|private|protected)\b)[A-Za-z_][\w.-]*:[^\S\r\n]*\r?$");

            // key: value。结尾是 ; 的 CSS 声明、结尾是 , { [ ( 的 JS 对象字面量不算
            score += Count(source, @"^[^\S\r\n]*[A-Za-z_][\w.-]*:[^\S\r\n]+[^\r\n]*[^;,{(\[\s][^\S\r\n]*\r?$");

            // 块标量头 key: |、key: >-
            score += 3 * Count(source, @":[^\S\r\n]+[|>][-+]?[^\S\r\n]*\r?$");

            // 锚点、别名、merge key
            score += 2 * Count(source, @"<<:[^\S\r\n]+\*|:[^\S\r\n]+[&*][\w-]+");
            return score;
        }

        private static int Count(string source, string pattern)
        {
            return LikelihoodPatterns.Count(source, pattern, RegexOptions.Multiline);
        }
    }
}
