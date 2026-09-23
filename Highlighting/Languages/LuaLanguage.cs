using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OneNoteCodeHelper.Highlighting.Languages
{
    /// <summary>Lua 词法着色。重点是长括号（--[==[ ]==] 与 [[ ]]）的层级匹配。</summary>
    internal sealed class LuaLanguage : ILanguage
    {
        private static readonly HashSet<string> Keywords = new HashSet<string>(StringComparer.Ordinal)
        {
            "and", "break", "do", "else", "elseif", "end", "for", "function", "goto", "if", "in",
            "local", "not", "or", "repeat", "return", "then", "until", "while"
        };

        private static readonly HashSet<string> Literals = new HashSet<string>(StringComparer.Ordinal)
        {
            "true", "false", "nil"
        };

        private static readonly HashSet<string> Builtins = new HashSet<string>(StringComparer.Ordinal)
        {
            // 基础库
            "assert", "collectgarbage", "dofile", "error", "getmetatable", "ipairs", "load", "loadfile",
            "next", "pairs", "pcall", "print", "rawequal", "rawget", "rawlen", "rawset", "require",
            "select", "setmetatable", "tonumber", "tostring", "type", "unpack", "xpcall",
            // 标准库表
            "coroutine", "debug", "io", "math", "os", "package", "string", "table", "utf8",
            // 常见全局
            "_G", "_VERSION", "self"
        };

        // 多字符运算符，按长度降序匹配，避免 "..." 被拆成 ".." + "." 或三个 "."
        private static readonly string[] MultiCharOperators =
        {
            "...", "..", "==", "~=", "<=", ">=", "::", "//", "<<", ">>"
        };

        private const string PunctuationChars = "(){}[];,.";

        public string Id => "lua";

        public string DisplayName => "Lua";

        public IEnumerable<Token> Tokenize(string source)
        {
            var c = new LexerCursor(source);

            while (!c.AtEnd)
            {
                var start = c.Position;
                var ch = c.Current;

                if (ch == '-' && c.Peek() == '-')
                {
                    c.Advance(2);

                    // --[[ ... ]] / --[==[ ... ]==] 长注释
                    var level = TryReadLongBracketOpen(c);
                    if (level >= 0)
                    {
                        SkipToLongBracketClose(c, level);
                    }
                    else
                    {
                        c.SkipToLineEnd();
                    }

                    c.Emit(TokenKind.Comment, start);
                    continue;
                }

                // [[ ... ]] / [==[ ... ]==] 长字符串
                if (ch == '[')
                {
                    var level = TryReadLongBracketOpen(c);
                    if (level >= 0)
                    {
                        SkipToLongBracketClose(c, level);
                        c.Emit(TokenKind.String, start);
                        continue;
                    }
                }

                if (ch == '"' || ch == '\'')
                {
                    ScanQuoted(c, ch);
                    c.Emit(TokenKind.String, start);
                    continue;
                }

                if (char.IsDigit(ch) || (ch == '.' && char.IsDigit(c.Peek())))
                {
                    ScanNumber(c);
                    c.Emit(TokenKind.Number, start);
                    continue;
                }

                if (LexerCursor.IsIdentifierStart(ch))
                {
                    c.SkipWhile(LexerCursor.IsIdentifierPart);
                    var word = source.Substring(start, c.Position - start);
                    c.Emit(ClassifyIdentifier(c, word), start);
                    continue;
                }

                if (char.IsWhiteSpace(ch))
                {
                    c.SkipWhile(char.IsWhiteSpace);
                    c.Emit(TokenKind.Plain, start);
                    continue;
                }

                var matched = false;
                foreach (var op in MultiCharOperators)
                {
                    if (c.Matches(op))
                    {
                        c.Advance(op.Length);
                        c.Emit(TokenKind.Operator, start);
                        matched = true;
                        break;
                    }
                }

                if (matched)
                {
                    continue;
                }

                c.Advance();
                c.Emit(PunctuationChars.IndexOf(ch) >= 0 ? TokenKind.Punctuation : TokenKind.Operator, start);
            }

            return c.Finish();
        }

        private static TokenKind ClassifyIdentifier(LexerCursor c, string word)
        {
            if (Keywords.Contains(word))
            {
                return TokenKind.Keyword;
            }

            if (Literals.Contains(word))
            {
                return TokenKind.Literal;
            }

            if (Builtins.Contains(word))
            {
                return TokenKind.Builtin;
            }

            return c.NextNonWhitespaceIs('(') ? TokenKind.Function : TokenKind.Plain;
        }

        /// <summary>
        /// 当前位置如果是长括号开头 [=*[ 就吃掉它并返回等号个数（层级），否则原地不动返回 -1。
        /// </summary>
        private static int TryReadLongBracketOpen(LexerCursor c)
        {
            if (c.Current != '[')
            {
                return -1;
            }

            var level = 0;
            while (c.Peek(1 + level) == '=')
            {
                level++;
            }

            if (c.Peek(1 + level) != '[')
            {
                return -1;
            }

            c.Advance(level + 2);
            return level;
        }

        /// <summary>吃到与指定层级匹配的 ]=*] 为止；找不到就吃到文件末尾。</summary>
        private static void SkipToLongBracketClose(LexerCursor c, int level)
        {
            while (!c.AtEnd)
            {
                if (c.Current == ']')
                {
                    var equals = 0;
                    while (c.Peek(1 + equals) == '=')
                    {
                        equals++;
                    }

                    if (equals == level && c.Peek(1 + equals) == ']')
                    {
                        c.Advance(level + 2);
                        return;
                    }
                }

                c.Advance();
            }
        }

        /// <summary>吃掉一段引号字面量。遇到换行就停，避免未闭合的引号吞掉整个文件。</summary>
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

        private static void ScanNumber(LexerCursor c)
        {
            if (c.Current == '0' && (c.Peek() == 'x' || c.Peek() == 'X'))
            {
                c.Advance(2);
                c.SkipWhile(LexerCursor.IsHexDigit);

                if (c.Current == '.')
                {
                    c.Advance();
                    c.SkipWhile(LexerCursor.IsHexDigit);
                }

                // 十六进制浮点的二进制指数，如 0x1p4
                if (c.Current == 'p' || c.Current == 'P')
                {
                    c.Advance();
                    if (c.Current == '+' || c.Current == '-')
                    {
                        c.Advance();
                    }

                    c.SkipWhile(char.IsDigit);
                }

                return;
            }

            c.SkipWhile(char.IsDigit);

            if (c.Current == '.')
            {
                c.Advance();
                c.SkipWhile(char.IsDigit);
            }

            if (c.Current == 'e' || c.Current == 'E')
            {
                c.Advance();
                if (c.Current == '+' || c.Current == '-')
                {
                    c.Advance();
                }

                c.SkipWhile(char.IsDigit);
            }
        }

        public int ScoreLikelihood(string source)
        {
            var score = 0;
            score += 4 * Count(source, @"\blocal\s+\w+");
            score += 4 * Count(source, @"\bfunction\b[^\n]*\)\s*$");
            score += 3 * Count(source, @"^\s*end\s*$");
            // 排除 CSS 自定义属性 --main-color: #fff; 这种形状的行
            score += 3 * Count(source, @"^\s*--(?![\w-]+\s*:.*;\s*$)");
            score += 2 * Count(source, @"\b(then|elseif|repeat|until)\b");
            score += 2 * Count(source, @"\bnil\b");
            score += 2 * Count(source, @"~=");
            score += 2 * Count(source, @"\b(ipairs|pairs|setmetatable)\s*\(");
            return score;
        }

        private static int Count(string source, string pattern)
        {
            return Regex.Matches(source, pattern, RegexOptions.Multiline).Count;
        }
    }
}
