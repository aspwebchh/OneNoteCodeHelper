using System;
using System.Collections.Generic;

namespace OneNoteCodeHelper.Highlighting.Languages
{
    /// <summary>
    /// XML 与 HTML 共用的标记语言词法。两者的差别只在 HTML 模式下的几处宽松规则：
    /// &lt;script&gt;/&lt;style&gt; 的内容是原始文本而不是标记，标签名不区分大小写。
    /// </summary>
    internal static class MarkupLexer
    {
        /// <summary>HTML 里 &lt;style&gt; 的内容交给 CSS 词法着色。</summary>
        private static readonly ILanguage Css = new CssLanguage();

        internal static IEnumerable<Token> Tokenize(string source, bool html)
        {
            var c = new LexerCursor(source);

            while (!c.AtEnd)
            {
                var start = c.Position;
                var ch = c.Current;

                if (ch == '<')
                {
                    if (c.Matches("<!--"))
                    {
                        c.Advance(4);
                        SkipPast(c, "-->");
                        c.Emit(TokenKind.Comment, start);
                        continue;
                    }

                    if (c.Matches("<![CDATA["))
                    {
                        SkipPast(c, "]]>");
                        c.Emit(TokenKind.String, start);
                        continue;
                    }

                    // <?xml ... ?> 声明、处理指令
                    if (c.Peek() == '?')
                    {
                        SkipPast(c, "?>");
                        c.Emit(TokenKind.Annotation, start);
                        continue;
                    }

                    // <!DOCTYPE ...> 等声明
                    if (c.Peek() == '!')
                    {
                        SkipPast(c, ">");
                        c.Emit(TokenKind.Annotation, start);
                        continue;
                    }

                    var closing = c.Peek() == '/';
                    if (IsNameStart(c.Peek(closing ? 2 : 1)))
                    {
                        var name = ScanTag(c, out var selfClosed);
                        if (html && !closing && !selfClosed)
                        {
                            SkipRawText(c, name);
                        }

                        continue;
                    }

                    // 孤立的 <，比如文本里的 a < b，当普通文字
                    c.Advance();
                    continue;
                }

                if (ch == '&' && TryScanEntity(c))
                {
                    c.Emit(TokenKind.Constant, start);
                    continue;
                }

                c.Advance();
                c.SkipWhile(x => x != '<' && x != '&');
                c.Emit(TokenKind.Plain, start);
            }

            return c.Finish();
        }

        /// <summary>吃掉一个开/闭标签（从 &lt; 到 &gt;），返回标签名。</summary>
        private static string ScanTag(LexerCursor c, out bool selfClosed)
        {
            selfClosed = false;

            var start = c.Position;
            c.Advance(c.Peek() == '/' ? 2 : 1);
            c.Emit(TokenKind.Punctuation, start);

            var nameStart = c.Position;
            c.SkipWhile(IsNameChar);
            var name = c.Source.Substring(nameStart, c.Position - nameStart);
            c.Emit(TokenKind.Tag, nameStart);

            var afterEquals = false;
            while (!c.AtEnd)
            {
                start = c.Position;
                var ch = c.Current;

                if (char.IsWhiteSpace(ch))
                {
                    c.SkipWhile(char.IsWhiteSpace);
                    continue;
                }

                if (ch == '>')
                {
                    c.Advance();
                    c.Emit(TokenKind.Punctuation, start);
                    return name;
                }

                if (ch == '/' && c.Peek() == '>')
                {
                    c.Advance(2);
                    c.Emit(TokenKind.Punctuation, start);
                    selfClosed = true;
                    return name;
                }

                // 标签没闭合就碰到下一个 <，交还给外层处理，免得把后面的内容都当成属性
                if (ch == '<')
                {
                    return name;
                }

                if (ch == '=')
                {
                    c.Advance();
                    c.Emit(TokenKind.Punctuation, start);
                    afterEquals = true;
                    continue;
                }

                if (ch == '"' || ch == '\'')
                {
                    c.Advance();
                    c.SkipWhile(x => x != ch);
                    c.Advance();
                    c.Emit(TokenKind.String, start);
                    afterEquals = false;
                    continue;
                }

                // 属性名，或者 HTML 里不加引号的属性值
                while (!c.AtEnd
                       && !char.IsWhiteSpace(c.Current)
                       && c.Current != '=' && c.Current != '>' && c.Current != '<'
                       && c.Current != '"' && c.Current != '\''
                       && !(c.Current == '/' && c.Peek() == '>'))
                {
                    c.Advance();
                }

                if (c.Position == start)
                {
                    c.Advance();
                }

                c.Emit(afterEquals ? TokenKind.String : TokenKind.Attribute, start);
                afterEquals = false;
            }

            return name;
        }

        /// <summary>
        /// HTML 的 &lt;script&gt; 与 &lt;style&gt; 里是原始文本，里面的 &lt; 不是标签。
        /// style 的内容交给 CSS 着色；script 的内容保持普通文字。
        /// </summary>
        private static void SkipRawText(LexerCursor c, string tagName)
        {
            var isStyle = string.Equals(tagName, "style", StringComparison.OrdinalIgnoreCase);
            if (!isStyle && !string.Equals(tagName, "script", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            var end = c.Source.IndexOf("</" + tagName, c.Position, StringComparison.OrdinalIgnoreCase);
            if (end < 0)
            {
                end = c.Source.Length;
            }

            if (isStyle)
            {
                c.EmitEmbedded(Css, end);
                return;
            }

            var start = c.Position;
            c.Advance(end - start);
            c.Emit(TokenKind.Plain, start);
        }

        /// <summary>&amp;name; / &amp;#123; / &amp;#x1F; 这类实体。不以分号结尾的不算，原地不动返回 false。</summary>
        private static bool TryScanEntity(LexerCursor c)
        {
            var length = 1;
            if (c.Peek(length) == '#')
            {
                length++;
                if (c.Peek(length) == 'x' || c.Peek(length) == 'X')
                {
                    length++;
                }

                while (LexerCursor.IsHexDigit(c.Peek(length)))
                {
                    length++;
                }
            }
            else
            {
                while (char.IsLetterOrDigit(c.Peek(length)))
                {
                    length++;
                }
            }

            if (length < 2 || c.Peek(length) != ';')
            {
                return false;
            }

            c.Advance(length + 1);
            return true;
        }

        /// <summary>吃到 terminator 之后；找不到就吃到末尾。</summary>
        private static void SkipPast(LexerCursor c, string terminator)
        {
            while (!c.AtEnd && !c.Matches(terminator))
            {
                c.Advance();
            }

            c.Advance(terminator.Length);
        }

        private static bool IsNameStart(char c) => char.IsLetter(c) || c == '_' || c == ':';

        private static bool IsNameChar(char c) =>
            char.IsLetterOrDigit(c) || c == '_' || c == ':' || c == '-' || c == '.';
    }
}
