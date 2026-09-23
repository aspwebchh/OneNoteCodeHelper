using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OneNoteCodeHelper.Highlighting.Languages
{
    /// <summary>
    /// Bash / sh 词法着色。Shell 的单词边界很宽（apt-get、./run.sh、--flag=x 都是一个词），
    /// 所以按「词」而不是按标识符切分；关键字和内置命令只在命令位置才认，
    /// 否则 echo log in now 里的 in 也会被当成关键字。heredoc 的正文要等到当前行结束后才开始。
    /// </summary>
    internal sealed class BashLanguage : ILanguage
    {
        private static readonly HashSet<string> Keywords = new HashSet<string>(StringComparer.Ordinal)
        {
            "if", "then", "else", "elif", "fi", "case", "esac", "for", "while", "until", "do", "done",
            "function", "select", "time", "coproc", "[[", "]]", "!"
        };

        private static readonly HashSet<string> Builtins = new HashSet<string>(StringComparer.Ordinal)
        {
            "alias", "bg", "bind", "break", "builtin", "caller", "cd", "command", "compgen", "complete",
            "continue", "declare", "dirs", "disown", "echo", "enable", "eval", "exec", "exit", "export", "false",
            "fc", "fg", "getopts", "hash", "help", "history", "jobs", "kill", "let", "local", "logout", "mapfile",
            "popd", "printf", "pushd", "pwd", "read", "readarray", "readonly", "return", "set", "shift", "shopt",
            "source", "suspend", "test", "times", "trap", "true", "type", "typeset", "ulimit", "umask", "unalias",
            "unset", "wait", "sudo"
        };

        /// <summary>这些关键字之后仍然是命令位置：if cmd、then cmd、do cmd……</summary>
        private static readonly HashSet<string> CommandPrefixKeywords = new HashSet<string>(StringComparer.Ordinal)
        {
            "if", "then", "else", "elif", "while", "until", "do", "time", "!", "sudo", "exec", "command", "builtin"
        };

        /// <summary>单词的终止符：空白与 shell 元字符、引号、$、反引号。</summary>
        private const string WordBreakChars = "|&;()<>'\"`${}";

        private sealed class PendingHeredoc
        {
            internal PendingHeredoc(string delimiter, bool stripTabs)
            {
                Delimiter = delimiter;
                StripTabs = stripTabs;
            }

            internal string Delimiter { get; }

            internal bool StripTabs { get; }
        }

        public string Id => "bash";

        public string DisplayName => "Bash";

        public IEnumerable<Token> Tokenize(string source)
        {
            var c = new LexerCursor(source);
            var pending = new List<PendingHeredoc>();
            var commandPosition = true;
            var expectFunctionName = false;
            var expectIn = false;

            // 我们自己开的括号层数。碰到没配对的 ) 说明是 case 分支的 pattern)，后面是命令位置。
            var parenDepth = 0;

            while (!c.AtEnd)
            {
                var start = c.Position;
                var ch = c.Current;

                if (ch == '\n')
                {
                    c.Advance();
                    c.Emit(TokenKind.Plain, start);
                    commandPosition = true;
                    expectIn = false;

                    foreach (var heredoc in pending)
                    {
                        ScanHeredocBody(c, heredoc);
                    }

                    pending.Clear();
                    continue;
                }

                if (char.IsWhiteSpace(ch))
                {
                    c.SkipWhile(x => char.IsWhiteSpace(x) && x != '\n');
                    c.Emit(TokenKind.Plain, start);
                    continue;
                }

                // 行尾的 \ 续行
                if (ch == '\\' && (c.Peek() == '\n' || c.Peek() == '\r'))
                {
                    c.Advance();
                    c.Emit(TokenKind.Operator, start);
                    c.Advance(c.Current == '\r' && c.Peek() == '\n' ? 2 : 1);
                    continue;
                }

                // # 只在词首才是注释：$#、a#b 都不是
                if (ch == '#')
                {
                    c.SkipToLineEnd();
                    c.Emit(TokenKind.Comment, start);
                    continue;
                }

                if (ch == '\'')
                {
                    c.Advance();
                    c.SkipWhile(x => x != '\'');
                    c.Advance();
                    c.Emit(TokenKind.String, start);
                    commandPosition = false;
                    continue;
                }

                if (ch == '$' && c.Peek() == '\'')
                {
                    c.Advance(2);
                    ScanEscapedUntil(c, '\'');
                    c.Emit(TokenKind.String, start);
                    commandPosition = false;
                    continue;
                }

                if (ch == '"')
                {
                    ScanDoubleQuoted(c);
                    commandPosition = false;
                    continue;
                }

                if (ch == '`')
                {
                    c.Advance();
                    ScanEscapedUntil(c, '`');
                    c.Emit(TokenKind.String, start);
                    commandPosition = false;
                    continue;
                }

                if (ch == '$')
                {
                    // $( 与 $(( 是命令替换 / 算术展开，括号里又回到命令位置
                    if (c.Peek() == '(')
                    {
                        var arithmetic = c.Peek(2) == '(';
                        c.Advance(arithmetic ? 3 : 2);
                        c.Emit(TokenKind.Operator, start);
                        parenDepth += arithmetic ? 2 : 1;
                        commandPosition = true;
                        continue;
                    }

                    if (TryScanVariable(c))
                    {
                        c.Emit(TokenKind.Variable, start);
                        commandPosition = false;
                        continue;
                    }
                }

                if (ch == '<' && c.Peek() == '<' && c.Peek(2) != '<')
                {
                    c.Advance(2);
                    var stripTabs = c.Current == '-';
                    if (stripTabs)
                    {
                        c.Advance();
                    }

                    c.Emit(TokenKind.Operator, start);

                    var heredoc = TryScanHeredocDelimiter(c, stripTabs);
                    if (heredoc != null)
                    {
                        pending.Add(heredoc);
                    }

                    continue;
                }

                if (ch == ';' || ch == '(' || ch == ')' || ch == '{' || ch == '}')
                {
                    c.Advance(ch == ';' && c.Peek() == ';' ? 2 : 1);
                    c.Emit(TokenKind.Punctuation, start);
                    expectIn = false;

                    if (ch == '(')
                    {
                        parenDepth++;
                        commandPosition = true;
                    }
                    else if (ch == ')')
                    {
                        commandPosition = parenDepth == 0;
                        parenDepth = Math.Max(0, parenDepth - 1);
                    }
                    else
                    {
                        commandPosition = true;
                    }

                    continue;
                }

                if (ch == '|' || ch == '&' || ch == '<' || ch == '>')
                {
                    c.Advance();
                    c.SkipWhile(x => x == '|' || x == '&' || x == '<' || x == '>');
                    c.Emit(TokenKind.Operator, start);
                    commandPosition = ch == '|' || ch == '&';
                    continue;
                }

                // 赋值 name=value / name+=value：左边标变量，= 之后是值
                if (char.IsLetter(ch) || ch == '_')
                {
                    c.SkipWhile(x => char.IsLetterOrDigit(x) || x == '_');
                    if (c.Current == '=' || (c.Current == '+' && c.Peek() == '='))
                    {
                        c.Emit(TokenKind.Variable, start);
                        var operatorStart = c.Position;
                        c.Advance(c.Current == '+' ? 2 : 1);
                        c.Emit(TokenKind.Operator, operatorStart);
                        continue;
                    }
                }

                ScanWordRest(c, start);
                var word = source.Substring(start, c.Position - start);
                var kind = ClassifyWord(c, word, start, commandPosition, expectFunctionName, expectIn);
                c.Emit(kind, start);

                expectFunctionName = kind == TokenKind.Keyword && word == "function";
                if (kind == TokenKind.Keyword && (word == "for" || word == "case" || word == "select"))
                {
                    expectIn = true;
                }
                else if (word == "in")
                {
                    expectIn = false;
                }

                commandPosition = (kind == TokenKind.Keyword || kind == TokenKind.Builtin)
                                  && CommandPrefixKeywords.Contains(word);
            }

            return c.Finish();
        }

        private static TokenKind ClassifyWord(
            LexerCursor c, string word, int start, bool commandPosition, bool expectFunctionName, bool expectIn)
        {
            if (expectFunctionName)
            {
                return TokenKind.Function;
            }

            if (expectIn && word == "in")
            {
                return TokenKind.Keyword;
            }

            if (commandPosition)
            {
                if (Keywords.Contains(word))
                {
                    return TokenKind.Keyword;
                }

                if (Builtins.Contains(word))
                {
                    return TokenKind.Builtin;
                }

                // name() { ... } 形式的函数定义
                if (c.Current == '(' && c.Peek() == ')')
                {
                    return TokenKind.Function;
                }
            }

            if (word == "]]")
            {
                return TokenKind.Keyword;
            }

            // -x、--flag、--out=file：前面是空白才算选项，a-b 里的减号不算
            if (word.Length > 1 && word[0] == '-' && (start == 0 || char.IsWhiteSpace(c.Source[start - 1])))
            {
                return IsAllDigits(word, 1) ? TokenKind.Number : TokenKind.Attribute;
            }

            return IsAllDigits(word, 0) ? TokenKind.Number : TokenKind.Plain;
        }

        /// <summary>吃到词尾。词里可以有 - . / = : 等，遇到空白或元字符才断；\x 是转义，不断词。</summary>
        private static void ScanWordRest(LexerCursor c, int start)
        {
            while (!c.AtEnd)
            {
                var ch = c.Current;
                if (ch == '\\' && c.Peek() != '\n' && c.Peek() != '\r' && c.Peek() != '\0')
                {
                    c.Advance(2);
                    continue;
                }

                if (char.IsWhiteSpace(ch) || WordBreakChars.IndexOf(ch) >= 0)
                {
                    break;
                }

                c.Advance();
            }

            // 落单的 $ 这类字符，自成一段，保证循环前进
            if (c.Position == start)
            {
                c.Advance();
            }
        }

        /// <summary>$name、${...}、$1、$@、$#、$?、$$、$!、$*、$-、$0。</summary>
        private static bool TryScanVariable(LexerCursor c)
        {
            var next = c.Peek();

            if (next == '{')
            {
                c.Advance(2);
                c.SkipWhile(x => x != '}' && x != '\n');
                c.Advance();
                return true;
            }

            if (char.IsLetter(next) || next == '_')
            {
                c.Advance();
                c.SkipWhile(x => char.IsLetterOrDigit(x) || x == '_');
                return true;
            }

            if (char.IsDigit(next) || "@#?$!*-".IndexOf(next) >= 0)
            {
                c.Advance(2);
                return true;
            }

            return false;
        }

        /// <summary>双引号字符串：\ 转义，里面的 $变量 单独着色。可以跨行。</summary>
        private static void ScanDoubleQuoted(LexerCursor c)
        {
            var segment = c.Position;
            c.Advance();

            while (!c.AtEnd)
            {
                var ch = c.Current;

                if (ch == '\\')
                {
                    c.Advance(2);
                    continue;
                }

                if (ch == '"')
                {
                    c.Advance();
                    break;
                }

                if (ch == '$' && c.Peek() != '(')
                {
                    c.Emit(TokenKind.String, segment);
                    var variableStart = c.Position;
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

        /// <summary>吃到未被 \ 转义的 terminator 之后；找不到就吃到末尾。</summary>
        private static void ScanEscapedUntil(LexerCursor c, char terminator)
        {
            while (!c.AtEnd)
            {
                if (c.Current == '\\')
                {
                    c.Advance(2);
                    continue;
                }

                if (c.Current == terminator)
                {
                    c.Advance();
                    return;
                }

                c.Advance();
            }
        }

        /// <summary>&lt;&lt; 之后的结束符：EOF、'EOF'、"EOF"、\EOF。当前行剩下的部分照常着色。</summary>
        private static PendingHeredoc TryScanHeredocDelimiter(LexerCursor c, bool stripTabs)
        {
            c.SkipWhile(x => x == ' ' || x == '\t');
            var start = c.Position;

            var quote = c.Current;
            if (quote == '\'' || quote == '"')
            {
                c.Advance();
                var nameStart = c.Position;
                c.SkipWhile(x => x != quote && x != '\n');
                var name = c.Source.Substring(nameStart, c.Position - nameStart);
                c.Advance();
                c.Emit(TokenKind.String, start);
                return name.Length > 0 ? new PendingHeredoc(name, stripTabs) : null;
            }

            if (quote == '\\')
            {
                c.Advance();
            }

            // 结束符必须以字母或下划线开头：算术里的 $((1 << 2)) 不能被当成 heredoc
            if (!char.IsLetter(c.Current) && c.Current != '_')
            {
                return null;
            }

            var wordStart = c.Position;
            c.SkipWhile(x => char.IsLetterOrDigit(x) || x == '_' || x == '-' || x == '.');

            var delimiter = c.Source.Substring(wordStart, c.Position - wordStart);
            c.Emit(TokenKind.String, start);
            return new PendingHeredoc(delimiter, stripTabs);
        }

        /// <summary>heredoc 正文：一直到某一行（<<- 时去掉行首 Tab 后）恰好等于结束符为止，含结束行。</summary>
        private static void ScanHeredocBody(LexerCursor c, PendingHeredoc heredoc)
        {
            var start = c.Position;

            while (!c.AtEnd)
            {
                var lineStart = c.Position;
                c.SkipToLineEnd();
                var line = c.Source.Substring(lineStart, c.Position - lineStart);
                if (heredoc.StripTabs)
                {
                    line = line.TrimStart('\t');
                }

                if (line == heredoc.Delimiter)
                {
                    break;
                }

                // 吃掉换行，继续下一行
                if (c.Current == '\r' && c.Peek() == '\n')
                {
                    c.Advance(2);
                }
                else
                {
                    c.Advance();
                }
            }

            c.Emit(TokenKind.String, start);

            // 多个 heredoc 挨着时，下一个从结束行的下一行开始
            if (c.Current == '\r' && c.Peek() == '\n')
            {
                c.Advance(2);
            }
            else if (c.Current == '\n')
            {
                c.Advance();
            }
        }

        private static bool IsAllDigits(string word, int from)
        {
            if (word.Length <= from)
            {
                return false;
            }

            for (var i = from; i < word.Length; i++)
            {
                if (!char.IsDigit(word[i]))
                {
                    return false;
                }
            }

            return true;
        }

        public int ScoreLikelihood(string source)
        {
            // 带闭合标签的是 XML/HTML：一行一个的 name="value" 属性不能被当成 shell 赋值
            if (LikelihoodPatterns.IsMatch(source, @"</[A-Za-z][\w:.-]*\s*>|<\?xml\b"))
            {
                return 0;
            }

            var score = 0;
            score += 10 * Count(source, @"\A#!.*\b(ba|z|k|da)?sh\b");
            score += 4 * Count(source, @"^[^\S\r\n]*(fi|done|esac)\s*(;.*)?$");
            score += 3 * Count(source, @"\$\{[\w#!]");
            score += 3 * Count(source, @"\[\[?\s+!?\s*-[a-zA-Z]\s");
            score += 2 * Count(source, @"\$[0-9@#?]");
            score += 3 * Count(source, @"^[^\S\r\n]*(export|readonly|declare|local)\s+\w+=");
            score += 2 * Count(source, @"\|\s*(grep|awk|sed|xargs|sort|uniq|head|tail|wc|cut|tr)\b");
            score += 2 * Count(source, @";\s*(then|do)\s*$");
            score += 2 * Count(source, @"^[^\S\r\n]*\w+=\S");
            return score;
        }

        private static int Count(string source, string pattern)
        {
            return LikelihoodPatterns.Count(source, pattern, RegexOptions.Multiline);
        }
    }
}
