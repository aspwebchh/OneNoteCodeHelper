namespace OneNoteCodeHelper.Highlighting.Languages
{
    /// <summary>
    /// C 系语言（C#、C/C++、JavaScript）和 Python 共用的扫描片段：块注释、引号字符串、数字、插值字符串。
    /// 各语言规则一致的部分放这里，差异通过参数表达。
    /// </summary>
    internal static class CommonScanners
    {
        /// <summary>吃掉 /* ... */；未闭合时吃到末尾。调用时当前位置在 / 上。</summary>
        internal static void ScanBlockComment(LexerCursor c)
        {
            c.Advance(2);
            while (!c.AtEnd && !(c.Current == '*' && c.Peek() == '/'))
            {
                c.Advance();
            }

            c.Advance(2);
        }

        /// <summary>/** ... */ 是文档注释，但 /**/ 只是个空的块注释。</summary>
        internal static bool IsDocBlockComment(LexerCursor c)
        {
            return c.Peek(2) == '*' && c.Peek(3) != '/';
        }

        /// <summary>吃掉一段单行引号字面量，\ 是转义符。遇到换行就停，避免未闭合的引号吞掉整个文件。</summary>
        internal static void ScanQuoted(LexerCursor c, char quote)
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

        /// <summary>
        /// C 风格数字：0x / 0b / 0o 前缀、小数、指数，以及紧跟的类型后缀（u、L、f、m、n、j……）。
        /// separator 是数字分隔符：多数语言是 _，C++14 是单引号。
        /// </summary>
        internal static void ScanNumber(LexerCursor c, char separator)
        {
            bool IsDigitOrSeparator(char ch, bool hex) =>
                char.IsDigit(ch) || (hex && LexerCursor.IsHexDigit(ch))
                || (ch == separator && (hex ? LexerCursor.IsHexDigit(c.Peek()) : char.IsDigit(c.Peek())));

            if (c.Current == '0' && (c.Peek() == 'x' || c.Peek() == 'X'))
            {
                c.Advance(2);
                while (!c.AtEnd && IsDigitOrSeparator(c.Current, true))
                {
                    c.Advance();
                }
            }
            else if (c.Current == '0' && (c.Peek() == 'b' || c.Peek() == 'B' || c.Peek() == 'o' || c.Peek() == 'O'))
            {
                c.Advance(2);
                while (!c.AtEnd && IsDigitOrSeparator(c.Current, false))
                {
                    c.Advance();
                }
            }
            else
            {
                while (!c.AtEnd && IsDigitOrSeparator(c.Current, false))
                {
                    c.Advance();
                }

                if (c.Current == '.' && char.IsDigit(c.Peek()))
                {
                    c.Advance();
                    while (!c.AtEnd && IsDigitOrSeparator(c.Current, false))
                    {
                        c.Advance();
                    }
                }

                if ((c.Current == 'e' || c.Current == 'E')
                    && (char.IsDigit(c.Peek()) || ((c.Peek() == '+' || c.Peek() == '-') && char.IsDigit(c.Peek(2)))))
                {
                    c.Advance(2);
                    c.SkipWhile(char.IsDigit);
                }
            }

            c.SkipWhile(char.IsLetter);
        }

        /// <summary>
        /// 插值字符串的主体（开头的前缀和引号已经吃掉，start 是整个字符串的起点）：
        /// 字面部分标 String，插值洞里的表达式交给 self 按本语言重新着色，洞两边的括号标 Punctuation。
        /// </summary>
        /// <param name="terminator">结束符，如 "、"""、`。</param>
        /// <param name="escape">转义符；逐字字符串没有转义符，传 '\0'。</param>
        /// <param name="holeOpen">插值洞的开头：C#/Python 是 {，JavaScript 是 ${。</param>
        /// <param name="multiLine">能否跨行。单行字符串遇到换行就停。</param>
        /// <param name="doubledTerminator">C# 逐字字符串里 "" 表示一个引号。</param>
        internal static void ScanInterpolatedBody(
            LexerCursor c, ILanguage self, int start, string terminator, char escape, string holeOpen,
            bool multiLine, bool doubledTerminator = false)
        {
            var segment = start;
            var braceEscapes = holeOpen == "{";

            while (!c.AtEnd)
            {
                var ch = c.Current;

                if (!multiLine && (ch == '\n' || ch == '\r'))
                {
                    break;
                }

                if (escape != '\0' && ch == escape)
                {
                    c.Advance(2);
                    continue;
                }

                if (doubledTerminator && c.Matches(terminator + terminator))
                {
                    c.Advance(terminator.Length * 2);
                    continue;
                }

                if (c.Matches(terminator))
                {
                    c.Advance(terminator.Length);
                    break;
                }

                // {{ 和 }} 是转义的花括号
                if (braceEscapes && (ch == '{' || ch == '}') && c.Peek() == ch)
                {
                    c.Advance(2);
                    continue;
                }

                if (c.Matches(holeOpen))
                {
                    var holeEnd = FindHoleEnd(c.Source, c.Position + holeOpen.Length, multiLine);
                    if (holeEnd < 0)
                    {
                        c.Advance(holeOpen.Length);
                        continue;
                    }

                    c.Emit(TokenKind.String, segment);
                    var braceStart = c.Position;
                    c.Advance(holeOpen.Length);
                    c.Emit(TokenKind.Punctuation, braceStart);
                    c.EmitEmbedded(self, holeEnd);
                    braceStart = c.Position;
                    c.Advance();
                    c.Emit(TokenKind.Punctuation, braceStart);
                    segment = c.Position;
                    continue;
                }

                c.Advance();
            }

            c.Emit(TokenKind.String, segment);
        }

        /// <summary>
        /// 从插值洞的内容起点找配对的 }。跳过洞里嵌套的花括号和字符串；找不到返回 -1。
        /// </summary>
        private static int FindHoleEnd(string source, int position, bool multiLine)
        {
            var depth = 0;
            for (var i = position; i < source.Length; i++)
            {
                var ch = source[i];

                if (!multiLine && (ch == '\n' || ch == '\r'))
                {
                    return -1;
                }

                if (ch == '"' || ch == '\'' || ch == '`')
                {
                    for (i++; i < source.Length && source[i] != ch && source[i] != '\n'; i++)
                    {
                        if (source[i] == '\\')
                        {
                            i++;
                        }
                    }

                    continue;
                }

                if (ch == '{')
                {
                    depth++;
                }
                else if (ch == '}')
                {
                    if (depth == 0)
                    {
                        return i;
                    }

                    depth--;
                }
            }

            return -1;
        }

        /// <summary>全大写（允许下划线和数字）按常量看，如 MAX_SIZE。至少两个字符，免得单字母类型参数 T 也算。</summary>
        internal static bool IsAllCaps(string word)
        {
            if (word.Length < 2)
            {
                return false;
            }

            var hasLetter = false;
            foreach (var ch in word)
            {
                if (char.IsLower(ch))
                {
                    return false;
                }

                hasLetter |= char.IsLetter(ch);
            }

            return hasLetter;
        }

        /// <summary>
        /// position 之前、中间只隔着空白的那个单词；前面紧挨的不是单词就返回空串。
        /// 用来看「new Foo」「class Foo」「def foo」里的 Foo 前面是什么。
        /// </summary>
        internal static string PreviousWord(LexerCursor c, int position)
        {
            var end = position;
            while (end > 0 && char.IsWhiteSpace(c.Source[end - 1]))
            {
                end--;
            }

            var begin = end;
            while (begin > 0 && (char.IsLetterOrDigit(c.Source[begin - 1]) || c.Source[begin - 1] == '_'))
            {
                begin--;
            }

            return c.Source.Substring(begin, end - begin);
        }

        /// <summary>当前位置往回跳过空白后的那个字符；到开头返回 '\0'。</summary>
        internal static char PreviousNonWhitespace(LexerCursor c, int position)
        {
            for (var i = position - 1; i >= 0; i--)
            {
                if (!char.IsWhiteSpace(c.Source[i]))
                {
                    return c.Source[i];
                }
            }

            return '\0';
        }
    }
}
