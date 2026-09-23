using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OneNoteCodeHelper.Highlighting.Languages
{
    /// <summary>
    /// Windows 批处理（.bat / .cmd）词法着色。整个语言不区分大小写。
    /// 要特别处理的是 echo 与 rem：它们后面到行尾都是普通文字，不能按关键字着色，
    /// 否则 echo if you do 里的 if、do 都会被染色。
    /// </summary>
    internal sealed class BatLanguage : ILanguage
    {
        private static readonly HashSet<string> Keywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "if", "else", "for", "in", "do", "goto", "call", "exit", "not", "exist", "defined", "errorlevel",
            "cmdextversion", "equ", "neq", "lss", "leq", "gtr", "geq", "setlocal", "endlocal",
            "enabledelayedexpansion", "disabledelayedexpansion", "enableextensions", "disableextensions",
            "nul", "con", "prn", "aux"
        };

        private static readonly HashSet<string> Builtins = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "assoc", "attrib", "break", "cd", "chdir", "chcp", "cls", "color", "copy", "date", "del", "dir",
            "echo", "erase", "find", "findstr", "ftype", "md", "mkdir", "mklink", "move", "path", "pause",
            "popd", "prompt", "pushd", "rd", "ren", "rename", "rmdir", "robocopy", "set", "shift", "start",
            "time", "timeout", "title", "type", "ver", "verify", "vol", "where", "xcopy", "choice", "taskkill",
            "tasklist", "reg", "sc", "net"
        };

        /// <summary>单词的终止符：空白、重定向与连接符、括号、引号、变量起始符、赋值号。</summary>
        private const string WordBreakChars = "&|<>()\"%!^=,;@";

        private const string SetArithmeticChars = "+-*/";

        public string Id => "bat";

        public string DisplayName => "Bat";

        public IEnumerable<Token> Tokenize(string source)
        {
            var c = new LexerCursor(source);
            var commandPosition = true;
            var parenDepth = 0;
            var expectLabel = false;
            var expectSetName = false;

            while (!c.AtEnd)
            {
                var start = c.Position;
                var ch = c.Current;

                if (ch == '\n' || ch == '\r')
                {
                    c.Advance();
                    commandPosition = true;
                    expectLabel = false;
                    expectSetName = false;
                    continue;
                }

                if (char.IsWhiteSpace(ch))
                {
                    c.SkipWhile(x => char.IsWhiteSpace(x) && x != '\n' && x != '\r');
                    continue;
                }

                // :: 注释与 :label 标签都只在行首成立
                if (ch == ':' && c.IsAtLineStart())
                {
                    var isComment = c.Peek() == ':';
                    c.SkipToLineEnd();
                    c.Emit(isComment ? TokenKind.Comment : TokenKind.Annotation, start);
                    continue;
                }

                // rem 注释：必须在命令位置，且后面是空白、行尾或 . 这类分隔符
                if (commandPosition && c.MatchesIgnoreCase("rem") && !IsWordChar(c.Peek(3)))
                {
                    c.SkipToLineEnd();
                    c.Emit(TokenKind.Comment, start);
                    continue;
                }

                if (ch == '%' || ch == '!')
                {
                    if (TryScanVariable(c))
                    {
                        c.Emit(TokenKind.Variable, start);
                        commandPosition = false;
                        continue;
                    }

                    c.Advance();
                    c.Emit(TokenKind.Operator, start);
                    continue;
                }

                if (ch == '"')
                {
                    ScanQuoted(c);
                    commandPosition = false;
                    continue;
                }

                // ^ 是转义符：^& ^| ^> 都当普通字符；行尾的 ^ 是续行
                if (ch == '^')
                {
                    c.Advance(c.Peek() == '\n' || c.Peek() == '\r' || c.Peek() == '\0' ? 1 : 2);
                    c.Emit(TokenKind.Operator, start);
                    continue;
                }

                if (ch == '(' || ch == ')')
                {
                    c.Advance();
                    c.Emit(TokenKind.Punctuation, start);
                    parenDepth = Math.Max(0, parenDepth + (ch == '(' ? 1 : -1));
                    commandPosition = ch == '(';
                    continue;
                }

                if (ch == '&' || ch == '|' || ch == '<' || ch == '>' || ch == '@' || ch == '=')
                {
                    c.Advance();
                    c.SkipWhile(x => x == ch || (ch == '>' && x == '&'));
                    c.Emit(TokenKind.Operator, start);
                    commandPosition = ch == '&' || ch == '|' || ch == '@';
                    continue;
                }

                if (ch == ',' || ch == ';')
                {
                    c.Advance();
                    c.Emit(TokenKind.Punctuation, start);
                    continue;
                }

                // /Q、/s、/A:H 这类开关，前面必须是空白
                if (ch == '/' && (char.IsLetterOrDigit(c.Peek()) || c.Peek() == '?') && char.IsWhiteSpace(c.Peek(-1)))
                {
                    c.Advance();
                    c.SkipWhile(x => !char.IsWhiteSpace(x) && WordBreakChars.IndexOf(x) < 0);
                    c.Emit(TokenKind.Attribute, start);
                    continue;
                }

                // echo、echo.、echo: 之后都是原样输出的文字
                if (c.MatchesIgnoreCase("echo") && !char.IsLetterOrDigit(c.Peek(4)) && c.Peek(4) != '_')
                {
                    c.Advance(4);
                    c.Emit(TokenKind.Builtin, start);
                    ScanEchoText(c, parenDepth);
                    commandPosition = false;
                    continue;
                }

                ScanWord(c, start, expectSetName);
                var word = source.Substring(start, c.Position - start);
                var kind = ClassifyWord(c, word, expectLabel, expectSetName);
                c.Emit(kind, start);

                expectLabel = kind == TokenKind.Keyword && string.Equals(word, "goto", StringComparison.OrdinalIgnoreCase);
                expectSetName = string.Equals(word, "set", StringComparison.OrdinalIgnoreCase)
                                || (expectSetName && kind == TokenKind.Attribute);

                // do ( ...、else ...、if 条件之后都可能接着命令，但只认 do/else 这两个明确的位置
                commandPosition = kind == TokenKind.Keyword
                                  && (string.Equals(word, "do", StringComparison.OrdinalIgnoreCase)
                                      || string.Equals(word, "else", StringComparison.OrdinalIgnoreCase));
            }

            return c.Finish();
        }

        private static TokenKind ClassifyWord(LexerCursor c, string word, bool expectLabel, bool expectSetName)
        {
            // call :sub、goto :eof、goto end
            if (word[0] == ':' || (expectLabel && !Keywords.Contains(word)))
            {
                return TokenKind.Annotation;
            }

            // set name=value、set /a name+=1 里的 name
            if (expectSetName && (c.Current == '=' || (SetArithmeticChars.IndexOf(c.Current) >= 0 && c.Peek() == '=')))
            {
                return TokenKind.Variable;
            }

            if (Keywords.Contains(word))
            {
                return TokenKind.Keyword;
            }

            if (Builtins.Contains(word))
            {
                return TokenKind.Builtin;
            }

            foreach (var ch in word)
            {
                if (!char.IsDigit(ch))
                {
                    return TokenKind.Plain;
                }
            }

            return TokenKind.Number;
        }

        /// <summary>set 之后的变量名遇到 set /a 的复合赋值运算符（+= -= *= /=）也要断开。</summary>
        private static void ScanWord(LexerCursor c, int start, bool setName)
        {
            c.SkipWhile(x => !char.IsWhiteSpace(x) && WordBreakChars.IndexOf(x) < 0
                             && !(setName && SetArithmeticChars.IndexOf(x) >= 0));

            if (c.Position == start)
            {
                c.Advance();
            }
        }

        /// <summary>
        /// echo 之后到行尾都是原样输出的文字，只有变量会被展开。
        /// 遇到未转义的 &amp; | &lt; &gt;（以及括号块里的 )）才结束。echo off / echo on 按关键字着色。
        /// </summary>
        private static void ScanEchoText(LexerCursor c, int parenDepth)
        {
            // echo. echo: echo( 这类写法紧跟的那个字符是分隔符
            var segment = c.Position;
            if (c.Current == '.' || c.Current == ':' || c.Current == '(' || c.Current == '/')
            {
                c.Advance();
            }

            c.SkipWhile(x => x == ' ' || x == '\t');
            var length = 0;
            while (char.IsLetter(c.Peek(length)))
            {
                length++;
            }

            var first = c.Source.Substring(c.Position, length);
            if ((string.Equals(first, "off", StringComparison.OrdinalIgnoreCase)
                 || string.Equals(first, "on", StringComparison.OrdinalIgnoreCase))
                && RestOfCommandIsBlank(c, length))
            {
                c.Emit(TokenKind.Plain, segment);
                var wordStart = c.Position;
                c.Advance(length);
                c.Emit(TokenKind.Keyword, wordStart);
                return;
            }

            while (!c.AtEnd && c.Current != '\n' && c.Current != '\r')
            {
                var ch = c.Current;

                if (ch == '^')
                {
                    c.Advance(2);
                    continue;
                }

                if (ch == '&' || ch == '|' || ch == '<' || ch == '>' || (ch == ')' && parenDepth > 0))
                {
                    break;
                }

                if (ch == '%' || ch == '!')
                {
                    var variableStart = c.Position;
                    c.Emit(TokenKind.Plain, segment);
                    segment = variableStart;
                    if (TryScanVariable(c))
                    {
                        c.Emit(TokenKind.Variable, variableStart);
                        segment = c.Position;
                        continue;
                    }
                }

                c.Advance();
            }

            c.Emit(TokenKind.Plain, segment);
        }

        /// <summary>echo off 之后只能是空白、行尾或命令连接符。</summary>
        private static bool RestOfCommandIsBlank(LexerCursor c, int offset)
        {
            for (var i = offset; ; i++)
            {
                var ch = c.Peek(i);
                if (ch == '\0' || ch == '\n' || ch == '\r' || ch == '&' || ch == '|' || ch == ')')
                {
                    return true;
                }

                if (ch != ' ' && ch != '\t')
                {
                    return false;
                }
            }
        }

        /// <summary>
        /// %VAR%、%VAR:~0,5%、%~dp0、%1、%*、%%i、%%~nxi、!VAR!。认不出来就原地不动返回 false。
        /// </summary>
        private static bool TryScanVariable(LexerCursor c)
        {
            var ch = c.Current;
            var next = c.Peek();

            if (ch == '%')
            {
                // for 循环变量 %%i / %%~nxi
                if (next == '%')
                {
                    var length = 2;
                    if (c.Peek(length) == '~')
                    {
                        length++;
                        while (char.IsLetter(c.Peek(length)) && char.IsLetter(c.Peek(length + 1)))
                        {
                            length++;
                        }
                    }

                    if (!char.IsLetter(c.Peek(length)))
                    {
                        return false;
                    }

                    c.Advance(length + 1);
                    return true;
                }

                // 参数 %1、%*、%~dp0、%~f1
                if (char.IsDigit(next) || next == '*')
                {
                    c.Advance(2);
                    return true;
                }

                if (next == '~')
                {
                    var length = 2;
                    while (char.IsLetter(c.Peek(length)))
                    {
                        length++;
                    }

                    if (!char.IsDigit(c.Peek(length)))
                    {
                        return false;
                    }

                    c.Advance(length + 1);
                    return true;
                }
            }

            // %NAME% / !NAME!：同一行里能找到配对的结束符才算
            if (!(char.IsLetter(next) || next == '_'))
            {
                return false;
            }

            for (var i = 2; ; i++)
            {
                var x = c.Peek(i);
                if (x == ch)
                {
                    c.Advance(i + 1);
                    return true;
                }

                if (x == '\0' || x == '\n' || x == '\r' || (ch == '!' && char.IsWhiteSpace(x)))
                {
                    return false;
                }
            }
        }

        /// <summary>双引号字符串，不能跨行；里面的变量单独着色。</summary>
        private static void ScanQuoted(LexerCursor c)
        {
            var segment = c.Position;
            c.Advance();

            while (!c.AtEnd && c.Current != '\n' && c.Current != '\r')
            {
                var ch = c.Current;
                if (ch == '"')
                {
                    c.Advance();
                    break;
                }

                if (ch == '%' || ch == '!')
                {
                    var variableStart = c.Position;
                    c.Emit(TokenKind.String, segment);
                    if (TryScanVariable(c))
                    {
                        c.Emit(TokenKind.Variable, variableStart);
                        segment = c.Position;
                        continue;
                    }

                    segment = variableStart;
                }

                c.Advance();
            }

            c.Emit(TokenKind.String, segment);
        }

        private static bool IsWordChar(char c) => c != '\0' && !char.IsWhiteSpace(c) && WordBreakChars.IndexOf(c) < 0 && c != '.';

        public int ScoreLikelihood(string source)
        {
            var score = 0;
            score += 10 * Count(source, @"^[^\S\r\n]*@echo\s+off\b");
            score += 3 * Count(source, @"^[^\S\r\n]*@?rem(\s|$)");
            score += 3 * Count(source, @"^[^\S\r\n]*::(?!\w+::)");
            score += 3 * Count(source, @"%\w+%|%~\w*\d|%%~?\w");
            score += 4 * Count(source, @"\bsetlocal\b");
            score += 4 * Count(source, @"\bif\s+(not\s+)?(exist|defined|errorlevel)\b");
            score += 2 * Count(source, @"\s(equ|neq|lss|leq|gtr|geq)\s");
            score += 2 * Count(source, @"\bgoto\s+:?\w+|\bcall\s+:\w+");
            score += 2 * Count(source, @"^[^\S\r\n]*set\s+(/[ap]\s+)?""?\w+=");
            score += 2 * Count(source, @"!\w+!");
            return score;
        }

        private static int Count(string source, string pattern)
        {
            return LikelihoodPatterns.Count(source, pattern, RegexOptions.Multiline | RegexOptions.IgnoreCase);
        }
    }
}
