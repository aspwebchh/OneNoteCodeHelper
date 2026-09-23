using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OneNoteCodeHelper.Highlighting.Languages
{
    /// <summary>
    /// Python 词法着色。字符串可以带 r/b/u/f 前缀、可以是三引号跨行；f-string 的 {expr} 按 Python 重新着色；
    /// 独占一行开头的三引号字符串按文档字符串（docstring）着色。
    /// </summary>
    internal sealed class PythonLanguage : ILanguage
    {
        private static readonly HashSet<string> Keywords = new HashSet<string>(StringComparer.Ordinal)
        {
            "and", "as", "assert", "async", "await", "break", "class", "continue", "def", "del", "elif", "else",
            "except", "finally", "for", "from", "global", "if", "import", "in", "is", "lambda", "nonlocal",
            "not", "or", "pass", "raise", "return", "try", "while", "with", "yield"
        };

        /// <summary>软关键字：只在「行首 + 行尾是冒号」的语句里才是关键字，match = re.match(...) 里的不算。</summary>
        private static readonly HashSet<string> SoftKeywords = new HashSet<string>(StringComparer.Ordinal)
        {
            "match", "case"
        };

        private static readonly HashSet<string> Literals = new HashSet<string>(StringComparer.Ordinal)
        {
            "True", "False", "None"
        };

        private static readonly HashSet<string> Builtins = new HashSet<string>(StringComparer.Ordinal)
        {
            "abs", "all", "any", "bool", "bytes", "callable", "chr", "dict", "dir", "enumerate", "eval", "exec",
            "filter", "float", "format", "getattr", "hasattr", "hash", "help", "hex", "id", "input", "int",
            "isinstance", "issubclass", "iter", "len", "list", "map", "max", "min", "next", "object", "open",
            "ord", "pow", "print", "range", "repr", "reversed", "round", "set", "setattr", "slice", "sorted",
            "str", "sum", "super", "tuple", "type", "vars", "zip", "self", "cls", "__name__", "__init__"
        };

        private const string PunctuationChars = "(){}[];,.:";

        public string Id => "python";

        public string DisplayName => "Python";

        public IEnumerable<Token> Tokenize(string source)
        {
            var c = new LexerCursor(source);

            while (!c.AtEnd)
            {
                var start = c.Position;
                var ch = c.Current;

                if (ch == '#')
                {
                    c.SkipToLineEnd();
                    c.Emit(TokenKind.Comment, start);
                    continue;
                }

                if (ch == '"' || ch == '\'')
                {
                    ScanString(c, start, string.Empty);
                    continue;
                }

                // 行首的 @decorator；其它位置的 @ 是矩阵乘法
                if (ch == '@' && c.IsAtLineStart() && (char.IsLetter(c.Peek()) || c.Peek() == '_'))
                {
                    c.Advance();
                    c.SkipWhile(x => char.IsLetterOrDigit(x) || x == '_' || x == '.');
                    c.Emit(TokenKind.Annotation, start);
                    continue;
                }

                if (char.IsDigit(ch) || (ch == '.' && char.IsDigit(c.Peek())))
                {
                    CommonScanners.ScanNumber(c, '_');
                    c.Emit(TokenKind.Number, start);
                    continue;
                }

                if (char.IsLetter(ch) || ch == '_')
                {
                    c.SkipWhile(x => char.IsLetterOrDigit(x) || x == '_');
                    var word = source.Substring(start, c.Position - start);

                    if ((c.Current == '"' || c.Current == '\'') && IsStringPrefix(word))
                    {
                        ScanString(c, start, word);
                        continue;
                    }

                    c.Emit(ClassifyIdentifier(c, word, start), start);
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

        private static TokenKind ClassifyIdentifier(LexerCursor c, string word, int start)
        {
            var afterMember = start > 0 && c.Source[start - 1] == '.';

            if (!afterMember && Keywords.Contains(word))
            {
                return TokenKind.Keyword;
            }

            if (!afterMember && SoftKeywords.Contains(word) && IsSoftKeywordStatement(c, start))
            {
                return TokenKind.Keyword;
            }

            if (Literals.Contains(word))
            {
                return TokenKind.Literal;
            }

            var previousWord = CommonScanners.PreviousWord(c, start);
            if (previousWord == "def")
            {
                return TokenKind.Function;
            }

            if (previousWord == "class")
            {
                return TokenKind.Type;
            }

            if (!afterMember && Builtins.Contains(word))
            {
                return TokenKind.Builtin;
            }

            if (c.NextNonWhitespaceIs('('))
            {
                return TokenKind.Function;
            }

            if (afterMember || !char.IsUpper(word[0]))
            {
                return TokenKind.Plain;
            }

            return CommonScanners.IsAllCaps(word) ? TokenKind.Constant : TokenKind.Type;
        }

        /// <summary>
        /// 当前位置在开头引号上，prefix 是已经吃掉的 r/b/u/f 前缀（可能为空），start 是整个字符串的起点。
        /// </summary>
        private void ScanString(LexerCursor c, int start, string prefix)
        {
            var quote = c.Current;
            var triple = c.Peek(1) == quote && c.Peek(2) == quote;
            var terminator = triple ? new string(quote, 3) : quote.ToString();
            var isFormat = prefix.IndexOf('f') >= 0 || prefix.IndexOf('F') >= 0;

            // 独占行首、没有 f/b 前缀的三引号字符串是 docstring
            var isDoc = triple && !isFormat && prefix.IndexOf('b') < 0 && prefix.IndexOf('B') < 0
                        && IsLineStart(c.Source, start);

            c.Advance(terminator.Length);

            if (isFormat)
            {
                CommonScanners.ScanInterpolatedBody(c, this, start, terminator, '\\', "{", multiLine: triple);
                return;
            }

            while (!c.AtEnd)
            {
                if (!triple && (c.Current == '\n' || c.Current == '\r'))
                {
                    break;
                }

                // 原始字符串里 \ 也会让紧跟的引号不结束字符串，所以统一跳过两个字符
                if (c.Current == '\\')
                {
                    c.Advance(2);
                    continue;
                }

                if (c.Matches(terminator))
                {
                    c.Advance(terminator.Length);
                    break;
                }

                c.Advance();
            }

            c.Emit(isDoc ? TokenKind.DocComment : TokenKind.String, start);
        }

        /// <summary>r、b、u、f 以及 rb、br、fr、rf 这些组合，大小写不限。</summary>
        private static bool IsStringPrefix(string word)
        {
            if (word.Length > 2)
            {
                return false;
            }

            var lower = word.ToLowerInvariant();
            return lower == "r" || lower == "b" || lower == "u" || lower == "f"
                   || lower == "rb" || lower == "br" || lower == "fr" || lower == "rf";
        }

        private static bool IsLineStart(string source, int position)
        {
            for (var i = position - 1; i >= 0; i--)
            {
                if (source[i] == '\n' || source[i] == '\r')
                {
                    return true;
                }

                if (source[i] != ' ' && source[i] != '\t')
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>match x: / case 1: —— 关键字在行首，而且这一行（去掉行尾注释）以冒号结束。</summary>
        private static bool IsSoftKeywordStatement(LexerCursor c, int start)
        {
            if (!IsLineStart(c.Source, start))
            {
                return false;
            }

            var lineEnd = c.Source.IndexOf('\n', start);
            var line = c.Source.Substring(start, (lineEnd < 0 ? c.Source.Length : lineEnd) - start);
            var comment = line.IndexOf('#');
            if (comment >= 0)
            {
                line = line.Substring(0, comment);
            }

            return line.TrimEnd().EndsWith(":", StringComparison.Ordinal);
        }

        public int ScoreLikelihood(string source)
        {
            var score = 0;
            score += 5 * Count(source, @"^\s*(async\s+)?def\s+\w+\s*\(.*\)\s*(->\s*[^:]+)?:\s*(#.*)?$");
            score += 5 * Count(source, @"^\s*from\s+[\w.]+\s+import\b");
            score += 3 * Count(source, @"^\s*import\s+[\w.]+(\s+as\s+\w+)?(\s*,\s*[\w.]+)*\s*$");
            score += 3 * Count(source, @"^\s*class\s+\w+(\(.*\))?:\s*$");
            score += 2 * Count(source, @"^\s*(if|elif|else|for|while|try|except|finally|with)\b.*:\s*(#.*)?$");
            score += 3 * Count(source, @"\belif\b");
            score += 2 * Count(source, @"\b(None|True|False)\b");
            score += 3 * Count(source, @"__\w+__");
            score += 3 * Count(source, @"\bdef\s+\w+\s*\(\s*self\b");
            score += 2 * Count(source, @"^\s*@[a-z_][\w.]*");
            score += 2 * Count(source, @"\blambda\b[^:\n]*:");
            score += Count(source, @"\bself\.\w+");
            score += Count(source, @"\bf[""'][^""'\n]*\{");
            return score;
        }

        private static int Count(string source, string pattern)
        {
            return Regex.Matches(source, pattern, RegexOptions.Multiline).Count;
        }
    }
}
