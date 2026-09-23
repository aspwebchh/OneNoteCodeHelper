using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OneNoteCodeHelper.Highlighting.Languages
{
    /// <summary>
    /// PowerShell 词法着色。整个语言不区分大小写；双引号字符串与 here-string 里的 $变量 单独着色，
    /// 形如 Verb-Noun 的命令按函数着色，-Name 按参数着色，-eq/-match 这类按运算符着色。
    /// </summary>
    internal sealed class PowerShellLanguage : ILanguage
    {
        private static readonly HashSet<string> Keywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "begin", "break", "catch", "class", "clean", "continue", "data", "define", "do", "dynamicparam",
            "else", "elseif", "end", "enum", "exit", "filter", "finally", "for", "foreach", "from", "function",
            "hidden", "if", "in", "param", "process", "return", "static", "switch", "throw", "trap", "try",
            "until", "using", "var", "while", "workflow", "parallel", "sequence", "inlinescript", "configuration"
        };

        private static readonly HashSet<string> Literals = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "true", "false", "null"
        };

        /// <summary>-eq、-match 这类「减号 + 单词」的运算符。其余 -Name 都是命令参数。</summary>
        private static readonly HashSet<string> WordOperators = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "eq", "ne", "gt", "ge", "lt", "le", "like", "notlike", "match", "notmatch", "contains", "notcontains",
            "in", "notin", "replace", "split", "join", "is", "isnot", "as", "and", "or", "not", "xor", "band",
            "bor", "bxor", "bnot", "shl", "shr", "f",
            "ieq", "ine", "igt", "ige", "ilt", "ile", "ilike", "inotlike", "imatch", "inotmatch", "icontains",
            "inotcontains", "iin", "inotin", "ireplace", "isplit",
            "ceq", "cne", "cgt", "cge", "clt", "cle", "clike", "cnotlike", "cmatch", "cnotmatch", "ccontains",
            "cnotcontains", "cin", "cnotin", "creplace", "csplit"
        };

        private const string PunctuationChars = "(){}[];,.";

        public string Id => "powershell";

        public string DisplayName => "PowerShell";

        public IEnumerable<Token> Tokenize(string source)
        {
            var c = new LexerCursor(source);
            var expectFunctionName = false;

            while (!c.AtEnd)
            {
                var start = c.Position;
                var ch = c.Current;

                if (ch == '<' && c.Peek() == '#')
                {
                    c.Advance(2);
                    while (!c.AtEnd && !(c.Current == '#' && c.Peek() == '>'))
                    {
                        c.Advance();
                    }

                    c.Advance(2);
                    c.Emit(TokenKind.Comment, start);
                    continue;
                }

                if (ch == '#')
                {
                    c.SkipToLineEnd();
                    c.Emit(TokenKind.Comment, start);
                    continue;
                }

                // here-string：@" 或 @' 之后必须紧跟换行，结束符 "@ / '@ 必须在行首
                if (ch == '@' && (c.Peek() == '"' || c.Peek() == '\'') && RestOfLineIsBlank(c, 2))
                {
                    var quote = c.Peek();
                    c.Advance(2);
                    if (quote == '"')
                    {
                        ScanDoubleQuotedBody(c, start, hereString: true);
                    }
                    else
                    {
                        while (!c.AtEnd && !(c.Current == '\'' && c.Peek() == '@' && IsAtColumnZero(c)))
                        {
                            c.Advance();
                        }

                        c.Advance(2);
                        c.Emit(TokenKind.String, start);
                    }

                    continue;
                }

                if (ch == '"')
                {
                    c.Advance();
                    ScanDoubleQuotedBody(c, start, hereString: false);
                    continue;
                }

                if (ch == '\'')
                {
                    ScanSingleQuoted(c);
                    c.Emit(TokenKind.String, start);
                    continue;
                }

                if (ch == '$' && IsVariableStart(c.Peek()))
                {
                    var literal = ScanVariable(c);
                    c.Emit(literal ? TokenKind.Literal : TokenKind.Variable, start);
                    continue;
                }

                // @splat 展开
                if (ch == '@' && (char.IsLetter(c.Peek()) || c.Peek() == '_'))
                {
                    c.Advance();
                    c.SkipWhile(IsWordChar);
                    c.Emit(TokenKind.Variable, start);
                    continue;
                }

                // [string]、[System.IO.Path]、[CmdletBinding()]。前面紧挨着标识符或 ] ) 的是下标，不是类型。
                if (ch == '[' && (char.IsLetter(c.Peek()) || c.Peek() == '_') && !IsIndexer(c))
                {
                    c.Advance();
                    c.Emit(TokenKind.Punctuation, start);
                    var nameStart = c.Position;
                    c.SkipWhile(x => IsWordChar(x) || x == '.' || x == '`');
                    c.Emit(c.Current == '(' ? TokenKind.Annotation : TokenKind.Type, nameStart);
                    continue;
                }

                // -Name 参数与 -eq 运算符。前面紧挨着单词字符时是减号（如 $a-1），不是参数。
                if ((ch == '-' || ch == '–') && char.IsLetter(c.Peek()) && !IsWordChar(c.Peek(-1)))
                {
                    c.Advance();
                    c.SkipWhile(IsWordChar);
                    var name = source.Substring(start + 1, c.Position - start - 1);
                    c.Emit(WordOperators.Contains(name) ? TokenKind.Operator : TokenKind.Attribute, start);
                    continue;
                }

                if (char.IsDigit(ch) || (ch == '.' && char.IsDigit(c.Peek()) && !IsWordChar(c.Peek(-1))))
                {
                    ScanNumber(c);
                    c.Emit(TokenKind.Number, start);
                    continue;
                }

                if (char.IsLetter(ch) || ch == '_')
                {
                    ScanWord(c);
                    var word = source.Substring(start, c.Position - start);
                    c.Emit(ClassifyWord(c, word, start, expectFunctionName), start);
                    expectFunctionName = string.Equals(word, "function", StringComparison.OrdinalIgnoreCase)
                                         || string.Equals(word, "filter", StringComparison.OrdinalIgnoreCase);
                    continue;
                }

                if (char.IsWhiteSpace(ch))
                {
                    c.SkipWhile(char.IsWhiteSpace);
                    c.Emit(TokenKind.Plain, start);
                    continue;
                }

                c.Advance();
                c.Emit(PunctuationChars.IndexOf(ch) >= 0 ? TokenKind.Punctuation : TokenKind.Operator, start);
            }

            return c.Finish();
        }

        private static TokenKind ClassifyWord(LexerCursor c, string word, int start, bool expectFunctionName)
        {
            // 成员访问 $x.Keys、[Math]::Round 后面的名字不是关键字
            var afterMember = start > 0 && (c.Source[start - 1] == '.' || c.Source[start - 1] == ':');

            if (!afterMember && Keywords.Contains(word))
            {
                return TokenKind.Keyword;
            }

            if (expectFunctionName || c.NextNonWhitespaceIs('('))
            {
                return TokenKind.Function;
            }

            // Get-ChildItem 这类 Verb-Noun 命令
            return !afterMember && word.IndexOf('-') > 0 ? TokenKind.Function : TokenKind.Plain;
        }

        /// <summary>单词里允许夹着连字符（Get-ChildItem），但连字符后面必须还是字母或数字。</summary>
        private static void ScanWord(LexerCursor c)
        {
            while (!c.AtEnd)
            {
                if (IsWordChar(c.Current))
                {
                    c.Advance();
                }
                else if (c.Current == '-' && char.IsLetterOrDigit(c.Peek()))
                {
                    c.Advance();
                }
                else
                {
                    break;
                }
            }
        }

        /// <summary>吃掉 $name / ${any thing} / $env:PATH / $_ / $? 等，返回是不是 $true/$false/$null。</summary>
        private static bool ScanVariable(LexerCursor c)
        {
            c.Advance(); // $

            if (c.Current == '{')
            {
                c.SkipWhile(x => x != '}' && x != '\n');
                c.Advance();
                return false;
            }

            if (!IsWordChar(c.Current))
            {
                c.Advance(); // $? $$ $^
                return false;
            }

            var nameStart = c.Position;
            c.SkipWhile(IsWordChar);

            // $env:PATH、$script:count 这类带作用域的变量；:: 是静态成员访问，不算
            if (c.Current == ':' && (char.IsLetter(c.Peek()) || c.Peek() == '_'))
            {
                c.Advance();
                c.SkipWhile(IsWordChar);
                return false;
            }

            return Literals.Contains(c.Source.Substring(nameStart, c.Position - nameStart));
        }

        /// <summary>
        /// 双引号字符串（或 here-string）的主体，开头的引号已经吃掉，start 是整个字符串的起点。
        /// 里面的 $变量 单独标 Variable，其余部分标 String。`（反引号）是转义符。
        /// </summary>
        private static void ScanDoubleQuotedBody(LexerCursor c, int start, bool hereString)
        {
            var segment = start;
            while (!c.AtEnd)
            {
                var ch = c.Current;

                if (ch == '`')
                {
                    c.Advance(2);
                    continue;
                }

                if (hereString)
                {
                    if (ch == '"' && c.Peek() == '@' && IsAtColumnZero(c))
                    {
                        c.Advance(2);
                        break;
                    }
                }
                else if (ch == '"')
                {
                    // "" 是转义的双引号
                    if (c.Peek() == '"')
                    {
                        c.Advance(2);
                        continue;
                    }

                    c.Advance();
                    break;
                }

                if (ch == '$' && IsVariableStart(c.Peek()) && c.Peek() != '(')
                {
                    c.Emit(TokenKind.String, segment);
                    var variableStart = c.Position;
                    ScanVariable(c);
                    c.Emit(TokenKind.Variable, variableStart);
                    segment = c.Position;
                    continue;
                }

                c.Advance();
            }

            c.Emit(TokenKind.String, segment);
        }

        /// <summary>单引号字符串，没有转义，'' 表示一个单引号。可以跨行。</summary>
        private static void ScanSingleQuoted(LexerCursor c)
        {
            c.Advance();
            while (!c.AtEnd)
            {
                if (c.Current == '\'')
                {
                    if (c.Peek() == '\'')
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

        private static void ScanNumber(LexerCursor c)
        {
            if (c.Current == '0' && (c.Peek() == 'x' || c.Peek() == 'X'))
            {
                c.Advance(2);
                c.SkipWhile(LexerCursor.IsHexDigit);
            }
            else
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
            }

            // 类型后缀 l/d/u 与数量级后缀 kb/mb/gb/tb/pb
            if (c.MatchesIgnoreCase("kb") || c.MatchesIgnoreCase("mb") || c.MatchesIgnoreCase("gb")
                || c.MatchesIgnoreCase("tb") || c.MatchesIgnoreCase("pb"))
            {
                c.Advance(2);
            }
            else if ("lLdDuU".IndexOf(c.Current) >= 0 && !IsWordChar(c.Peek()))
            {
                c.Advance();
            }
        }

        /// <summary>从当前位置 + offset 到行尾是不是只有空白。</summary>
        private static bool RestOfLineIsBlank(LexerCursor c, int offset)
        {
            for (var i = offset; ; i++)
            {
                var ch = c.Peek(i);
                if (ch == '\n' || ch == '\r' || ch == '\0')
                {
                    return true;
                }

                if (ch != ' ' && ch != '\t')
                {
                    return false;
                }
            }
        }

        /// <summary>here-string 的结束符 "@ / '@ 必须顶格，前面有空白都不算（PowerShell 本身也会报错）。</summary>
        private static bool IsAtColumnZero(LexerCursor c)
        {
            var previous = c.Peek(-1);
            return previous == '\n' || previous == '\r';
        }

        /// <summary>[ 前面紧挨着标识符、] 或 ) 时是下标访问：$list[0]、$map[$key]。</summary>
        private static bool IsIndexer(LexerCursor c)
        {
            var previous = c.Peek(-1);
            return IsWordChar(previous) || previous == ']' || previous == ')';
        }

        private static bool IsVariableStart(char c) =>
            IsWordChar(c) || c == '{' || c == '?' || c == '$' || c == '^';

        private static bool IsWordChar(char c) => char.IsLetterOrDigit(c) || c == '_';

        public int ScoreLikelihood(DetectionSample sample)
        {
            var code = sample.Code;
            var score = 0;
            score += 4 * Count(code,
                @"\b(Get|Set|New|Remove|Add|Write|Read|Start|Stop|Test|Invoke|Import|Export|Out|Select|Where|ForEach" +
                @"|Sort|Group|Format|ConvertTo|ConvertFrom|Join|Split|Copy|Move|Clear|Enable|Disable|Register" +
                @"|Unregister|Resolve|Wait|Measure|Update|Install|Uninstall|Push|Pop|Restart|Show|Use)-[A-Z][A-Za-z]+");
            score += 4 * Count(code, @"^[^\S\r\n]*param\s*\(", RegexOptions.IgnoreCase);

            // 变量常写在双引号字符串里，要看原文
            score += 3 * Count(sample.Raw, @"\$(true|false|null|_|PSScriptRoot|PSCmdlet|args|env:\w+)\b", RegexOptions.IgnoreCase);
            score += 3 * Count(code,
                @"\[(string|int|bool|switch|object|hashtable|array|datetime|Parameter|CmdletBinding|System\.[\w.]+)(\[\])?[\]\(]",
                RegexOptions.IgnoreCase);
            score += 2 * Count(code, @"\s-(eq|ne|gt|ge|lt|le|like|notlike|match|notmatch|contains|notcontains|and|or|not)\s",
                RegexOptions.IgnoreCase);
            score += 3 * Count(sample.Raw, @"<#|#>");
            score += 2 * Count(code, @"\$\w+\s*=[^=~]");
            score += Count(code, @"\s-[A-Z][a-z]+[A-Z]?\w*\b");
            return score;
        }

        private static int Count(string source, string pattern, RegexOptions options = RegexOptions.None)
        {
            return LikelihoodPatterns.Count(source, pattern, RegexOptions.Multiline | options);
        }
    }
}
