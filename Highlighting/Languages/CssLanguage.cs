using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OneNoteCodeHelper.Highlighting.Languages
{
    /// <summary>
    /// CSS 词法着色。难点是同一个标识符在选择器里和声明里含义完全不同（div 是元素，color 是属性），
    /// 这里在每条语句开头向后看：先碰到 { 就是选择器（或 at-rule 前导），先碰到 ; 或 } 就是声明。
    /// 这样不用维护嵌套栈，@media 块和 CSS 嵌套写法都能正确处理。
    /// </summary>
    internal sealed class CssLanguage : ILanguage
    {
        private enum Mode
        {
            /// <summary>选择器：div.cls#id:hover</summary>
            Selector,

            /// <summary>@media screen and (max-width: 600px) 这类 at-rule 前导。</summary>
            AtRule,

            /// <summary>声明的属性名，冒号之前。</summary>
            Property,

            /// <summary>声明的值，冒号之后到分号。</summary>
            Value
        }

        private static readonly HashSet<string> AtRuleKeywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "and", "or", "not", "only"
        };

        public string Id => "css";

        public string DisplayName => "CSS";

        public IEnumerable<Token> Tokenize(string source)
        {
            var c = new LexerCursor(source);
            var depth = 0;
            var parenDepth = 0;
            var atStatementStart = true;
            var mode = Mode.Selector;

            while (!c.AtEnd)
            {
                var start = c.Position;
                var ch = c.Current;

                if (char.IsWhiteSpace(ch))
                {
                    c.SkipWhile(char.IsWhiteSpace);
                    c.Emit(TokenKind.Plain, start);
                    continue;
                }

                if (ch == '/' && c.Peek() == '*')
                {
                    c.Advance(2);
                    while (!c.AtEnd && !(c.Current == '*' && c.Peek() == '/'))
                    {
                        c.Advance();
                    }

                    c.Advance(2);
                    c.Emit(TokenKind.Comment, start);
                    continue;
                }

                if (atStatementStart)
                {
                    atStatementStart = false;
                    parenDepth = 0;
                    mode = ch == '@' ? Mode.AtRule
                        : LooksLikePrelude(source, c.Position, depth) ? Mode.Selector
                        : Mode.Property;
                }

                if (ch == '{')
                {
                    c.Advance();
                    c.Emit(TokenKind.Punctuation, start);
                    depth++;
                    atStatementStart = true;
                    continue;
                }

                if (ch == '}')
                {
                    c.Advance();
                    c.Emit(TokenKind.Punctuation, start);
                    depth = Math.Max(0, depth - 1);
                    atStatementStart = true;
                    continue;
                }

                if (ch == ';')
                {
                    c.Advance();
                    c.Emit(TokenKind.Punctuation, start);
                    atStatementStart = true;
                    continue;
                }

                if (ch == '"' || ch == '\'')
                {
                    ScanQuoted(c, ch);
                    c.Emit(TokenKind.String, start);
                    continue;
                }

                if (ch == '@' && IsIdentStart(c, 1))
                {
                    c.Advance();
                    SkipIdent(c);
                    c.Emit(TokenKind.Keyword, start);
                    continue;
                }

                if (ch == '!' && IsIdentStart(c, 1))
                {
                    c.Advance();
                    SkipIdent(c);
                    c.Emit(TokenKind.Keyword, start);
                    continue;
                }

                if (ch == '(' || ch == ')')
                {
                    parenDepth = Math.Max(0, parenDepth + (ch == '(' ? 1 : -1));
                    c.Advance();
                    c.Emit(TokenKind.Punctuation, start);
                    continue;
                }

                switch (mode)
                {
                    case Mode.Selector:
                        ScanSelectorPart(c, parenDepth);
                        break;
                    case Mode.AtRule:
                        ScanAtRulePart(c);
                        break;
                    case Mode.Property:
                        if (ch == ':')
                        {
                            c.Advance();
                            c.Emit(TokenKind.Punctuation, start);
                            mode = Mode.Value;
                        }
                        else if (IsIdentStart(c, 0))
                        {
                            SkipIdent(c);
                            c.Emit(IsCustomProperty(source, start) ? TokenKind.Variable : TokenKind.Attribute, start);
                        }
                        else
                        {
                            c.Advance();
                            c.Emit(TokenKind.Plain, start);
                        }

                        break;
                    default:
                        ScanValuePart(c);
                        break;
                }
            }

            return c.Finish();
        }

        /// <summary>选择器里的一个片段：元素名、.class、#id、:pseudo、[attr]、组合符。</summary>
        private static void ScanSelectorPart(LexerCursor c, int parenDepth)
        {
            var start = c.Position;
            var ch = c.Current;

            if (ch == '.' && IsIdentStart(c, 1))
            {
                c.Advance();
                SkipIdent(c);
                c.Emit(TokenKind.Type, start);
                return;
            }

            if (ch == '#' && IsIdentPart(c.Peek()))
            {
                c.Advance();
                SkipIdent(c);
                c.Emit(TokenKind.Constant, start);
                return;
            }

            if (ch == ':')
            {
                c.Advance(c.Peek() == ':' ? 2 : 1);
                SkipIdent(c);
                c.Emit(TokenKind.Annotation, start);
                return;
            }

            if (ch == '[')
            {
                c.Advance();
                c.Emit(TokenKind.Punctuation, start);
                ScanAttributeSelector(c);
                return;
            }

            if (char.IsDigit(ch))
            {
                ScanNumber(c);
                c.Emit(TokenKind.Number, start);
                return;
            }

            if (IsIdentStart(c, 0))
            {
                SkipIdent(c);

                // 伪类参数里的 odd、2n+1 这类不是元素名
                c.Emit(parenDepth > 0 ? TokenKind.Plain : TokenKind.Tag, start);
                return;
            }

            c.Advance();
            c.Emit(ch == ',' ? TokenKind.Punctuation : TokenKind.Operator, start);
        }

        /// <summary>[attr]、[attr="v"]、[attr~=v i]，吃到 ] 为止。</summary>
        private static void ScanAttributeSelector(LexerCursor c)
        {
            var afterOperator = false;
            while (!c.AtEnd && c.Current != '{' && c.Current != '\n')
            {
                var start = c.Position;
                var ch = c.Current;

                if (ch == ']')
                {
                    c.Advance();
                    c.Emit(TokenKind.Punctuation, start);
                    return;
                }

                if (ch == '"' || ch == '\'')
                {
                    ScanQuoted(c, ch);
                    c.Emit(TokenKind.String, start);
                    continue;
                }

                if (IsIdentStart(c, 0))
                {
                    SkipIdent(c);
                    c.Emit(afterOperator ? TokenKind.String : TokenKind.Attribute, start);
                    continue;
                }

                if (ch == '=')
                {
                    afterOperator = true;
                }

                c.Advance();
                c.Emit(char.IsWhiteSpace(ch) ? TokenKind.Plain : TokenKind.Operator, start);
            }
        }

        /// <summary>@media 之后、{ 之前的部分：screen and (max-width: 600px)。</summary>
        private static void ScanAtRulePart(LexerCursor c)
        {
            var start = c.Position;
            var ch = c.Current;

            if (char.IsDigit(ch) || (ch == '.' && char.IsDigit(c.Peek())))
            {
                ScanNumber(c);
                c.Emit(TokenKind.Number, start);
                return;
            }

            if (IsIdentStart(c, 0))
            {
                SkipIdent(c);
                var word = c.Source.Substring(start, c.Position - start);

                if (c.Current == '(')
                {
                    ScanFunctionName(c, start, word);
                    return;
                }

                c.Emit(AtRuleKeywords.Contains(word) ? TokenKind.Keyword
                    : c.NextNonWhitespaceIs(':') ? TokenKind.Attribute
                    : TokenKind.Plain, start);
                return;
            }

            c.Advance();
            c.Emit(ch == ',' ? TokenKind.Punctuation : TokenKind.Operator, start);
        }

        /// <summary>声明值里的一个片段：数字与单位、#颜色、函数、var(--x)。</summary>
        private static void ScanValuePart(LexerCursor c)
        {
            var start = c.Position;
            var ch = c.Current;

            if (char.IsDigit(ch)
                || (ch == '.' && char.IsDigit(c.Peek()))
                || ((ch == '-' || ch == '+') && (char.IsDigit(c.Peek()) || (c.Peek() == '.' && char.IsDigit(c.Peek(2))))))
            {
                if (ch == '-' || ch == '+')
                {
                    c.Advance();
                }

                ScanNumber(c);
                c.Emit(TokenKind.Number, start);
                return;
            }

            if (ch == '#' && LexerCursor.IsHexDigit(c.Peek()))
            {
                c.Advance();
                c.SkipWhile(IsIdentPart);
                c.Emit(TokenKind.Number, start);
                return;
            }

            if (IsIdentStart(c, 0))
            {
                SkipIdent(c);
                var word = c.Source.Substring(start, c.Position - start);

                if (c.Current == '(')
                {
                    ScanFunctionName(c, start, word);
                    return;
                }

                c.Emit(IsCustomProperty(c.Source, start) ? TokenKind.Variable : TokenKind.Plain, start);
                return;
            }

            c.Advance();
            c.Emit(ch == ',' || ch == ':' ? TokenKind.Punctuation : TokenKind.Operator, start);
        }

        /// <summary>
        /// 标识符后紧跟 ( 的情况。url( 的参数可以不加引号、里面还会有 // 和 :，
        /// 必须整段当字符串吃掉，否则会被误切成注释或声明。
        /// </summary>
        private static void ScanFunctionName(LexerCursor c, int start, string word)
        {
            c.Emit(TokenKind.Function, start);

            if (!string.Equals(word, "url", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            var parenStart = c.Position;
            c.Advance();
            c.Emit(TokenKind.Punctuation, parenStart);

            var argStart = c.Position;
            c.SkipWhile(char.IsWhiteSpace);
            if (c.Current == '"' || c.Current == '\'')
            {
                return;
            }

            c.SkipWhile(x => x != ')' && x != '\n' && x != '\r');
            c.Emit(TokenKind.String, argStart);
        }

        /// <summary>
        /// 从语句开头往后看，先碰到 { 就是选择器/前导，先碰到 ; 或 } 就是声明。
        /// 到末尾都没碰到时（比如只贴了一行），顶层按选择器、块内按声明。
        /// </summary>
        private static bool LooksLikePrelude(string source, int position, int depth)
        {
            for (var i = position; i < source.Length; i++)
            {
                var ch = source[i];
                switch (ch)
                {
                    case '{':
                        return true;
                    case ';':
                    case '}':
                        return false;
                    case '"':
                    case '\'':
                        i = source.IndexOf(ch, i + 1);
                        if (i < 0)
                        {
                            return depth == 0;
                        }

                        break;
                    case '/':
                        if (i + 1 < source.Length && source[i + 1] == '*')
                        {
                            i = source.IndexOf("*/", i + 2, StringComparison.Ordinal);
                            if (i < 0)
                            {
                                return depth == 0;
                            }

                            i++;
                        }

                        break;
                }
            }

            return depth == 0;
        }

        private static bool IsCustomProperty(string source, int start)
        {
            return start + 1 < source.Length && source[start] == '-' && source[start + 1] == '-';
        }

        /// <summary>CSS 标识符可以以 - 或 -- 开头（-webkit-xxx、--custom），但 -1 这类是数字。</summary>
        private static bool IsIdentStart(LexerCursor c, int offset)
        {
            var ch = c.Peek(offset);
            if (ch == '-')
            {
                var next = c.Peek(offset + 1);
                return char.IsLetter(next) || next == '_' || next == '-';
            }

            return char.IsLetter(ch) || ch == '_' || ch > 127;
        }

        private static bool IsIdentPart(char c) => char.IsLetterOrDigit(c) || c == '-' || c == '_' || c > 127;

        private static void SkipIdent(LexerCursor c)
        {
            c.SkipWhile(IsIdentPart);
        }

        private static void ScanQuoted(LexerCursor c, char quote)
        {
            c.Advance();
            while (!c.AtEnd && c.Current != '\n' && c.Current != '\r')
            {
                if (c.Current == '\\')
                {
                    c.Advance(2);
                    continue;
                }

                if (c.Current == quote)
                {
                    c.Advance();
                    return;
                }

                c.Advance();
            }
        }

        /// <summary>数字连同单位一起吃：12px、1.5em、50%、1e3。</summary>
        private static void ScanNumber(LexerCursor c)
        {
            c.SkipWhile(char.IsDigit);

            if (c.Current == '.' && char.IsDigit(c.Peek()))
            {
                c.Advance();
                c.SkipWhile(char.IsDigit);
            }

            if ((c.Current == 'e' || c.Current == 'E')
                && (char.IsDigit(c.Peek()) || ((c.Peek() == '+' || c.Peek() == '-') && char.IsDigit(c.Peek(2)))))
            {
                c.Advance(2);
                c.SkipWhile(char.IsDigit);
            }

            if (c.Current == '%')
            {
                c.Advance();
                return;
            }

            c.SkipWhile(char.IsLetter);
        }

        public int ScoreLikelihood(string source)
        {
            // 带标签的是 HTML/XML，里面的 <style> 不能把整段拉成 CSS
            if (LikelihoodPatterns.IsMatch(source, @"<[A-Za-z!/]"))
            {
                return 0;
            }

            var score = 0;
            score += 3 * Count(source,
                @"(^|[{;\s])(color|background(-\w+)?|margin(-\w+)?|padding(-\w+)?|font(-\w+)?|display|width|height" +
                @"|min-width|max-width|min-height|max-height|border(-\w+)?|position|top|left|right|bottom|flex(-\w+)?" +
                @"|text-\w+|z-index|opacity|overflow(-[xy])?|cursor|align-\w+|justify-\w+|gap|grid(-\w+)?|transition" +
                @"|transform|animation(-\w+)?|box-\w+|line-height|content|visibility|outline|white-space)\s*:[^:]");
            score += 3 * Count(source, @"(^|[{;\s])--[\w-]+\s*:");
            score += 4 * Count(source, @"@(media|import|keyframes|font-face|supports|charset|layer|container)\b");
            // 选择器那一行后面紧跟 {，或者 { 另起一行（Allman 风格）。[^;{}\r\n] 不能跨行：
            // 不排除换行的话，#define 这种没有 ;{} 的行会一路扫到文件末尾，几千行就要几十秒。
            score += 3 * Count(source, @"^[^\S\r\n]*[.#][\w-]+[^;{}\r\n]*\s*\{");
            score += Count(source, @"\b\d+(\.\d+)?(px|em|rem|vh|vw|pt|ms|deg)\b");
            score += 2 * Count(source, @"!important\b");
            return score;
        }

        private static int Count(string source, string pattern)
        {
            return LikelihoodPatterns.Count(source, pattern, RegexOptions.Multiline);
        }
    }
}
