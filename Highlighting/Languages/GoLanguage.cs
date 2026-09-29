using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OneNoteCodeHelper.Highlighting.Languages
{
    /// <summary>
    /// Go 词法着色。字符串有两种：带转义的 "..." 和可以跨行、不转义的反引号原始字符串；'x' 是 rune。
    /// Go 用首字母大小写区分导出与否，大写开头的不一定是类型，所以大写启发式要避开函数调用、结构体字段和键。
    /// </summary>
    internal sealed class GoLanguage : ILanguage
    {
        private static readonly HashSet<string> Keywords = new HashSet<string>(StringComparer.Ordinal)
        {
            "break", "case", "chan", "const", "continue", "default", "defer", "else", "fallthrough", "for",
            "func", "go", "goto", "if", "import", "interface", "map", "package", "range", "return", "select",
            "struct", "switch", "type", "var"
        };

        private static readonly HashSet<string> Literals = new HashSet<string>(StringComparer.Ordinal)
        {
            "true", "false", "nil", "iota"
        };

        private static readonly HashSet<string> BuiltinTypes = new HashSet<string>(StringComparer.Ordinal)
        {
            "bool", "byte", "complex64", "complex128", "error", "float32", "float64", "int", "int8", "int16",
            "int32", "int64", "rune", "string", "uint", "uint8", "uint16", "uint32", "uint64", "uintptr", "any",
            "comparable"
        };

        /// <summary>内置函数。new 在 Go 里也是函数，不是关键字。</summary>
        private static readonly HashSet<string> BuiltinFunctions = new HashSet<string>(StringComparer.Ordinal)
        {
            "append", "cap", "clear", "close", "complex", "copy", "delete", "imag", "len", "make", "max", "min",
            "new", "panic", "print", "println", "real", "recover"
        };

        private const string PunctuationChars = "(){}[];,.";

        public string Id => "go";

        public string DisplayName => "Go";

        public IEnumerable<Token> Tokenize(string source)
        {
            var c = new LexerCursor(source);

            while (!c.AtEnd)
            {
                var start = c.Position;
                var ch = c.Current;

                if (ch == '/' && c.Peek() == '/')
                {
                    c.SkipToLineEnd();
                    c.Emit(TokenKind.Comment, start);
                    continue;
                }

                if (ch == '/' && c.Peek() == '*')
                {
                    CommonScanners.ScanBlockComment(c);
                    c.Emit(TokenKind.Comment, start);
                    continue;
                }

                if (ch == '"')
                {
                    CommonScanners.ScanQuoted(c, '"');
                    c.Emit(TokenKind.String, start);
                    continue;
                }

                // 原始字符串：没有转义，可以跨行，到下一个反引号为止
                if (ch == '`')
                {
                    c.Advance();
                    c.SkipWhile(x => x != '`');
                    c.Advance();
                    c.Emit(TokenKind.String, start);
                    continue;
                }

                if (ch == '\'')
                {
                    CommonScanners.ScanQuoted(c, '\'');
                    c.Emit(TokenKind.Char, start);
                    continue;
                }

                if (char.IsDigit(ch) || (ch == '.' && char.IsDigit(c.Peek())))
                {
                    CommonScanners.ScanNumber(c, '_');
                    c.Emit(TokenKind.Number, start);
                    continue;
                }

                if (IsIdentifierStart(ch))
                {
                    var atLineStart = c.IsAtLineStart();
                    c.SkipWhile(IsIdentifierPart);
                    var word = source.Substring(start, c.Position - start);
                    c.Emit(ClassifyIdentifier(c, word, start, atLineStart), start);
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

        private static TokenKind ClassifyIdentifier(LexerCursor c, string word, int start, bool atLineStart)
        {
            var afterMember = start > 0 && c.Source[start - 1] == '.';
            var call = c.NextNonWhitespaceIs('(');

            if (!afterMember && Keywords.Contains(word))
            {
                return TokenKind.Keyword;
            }

            if (!afterMember && Literals.Contains(word))
            {
                return TokenKind.Literal;
            }

            if (!afterMember && BuiltinTypes.Contains(word))
            {
                return TokenKind.Type;
            }

            if (!afterMember && call && BuiltinFunctions.Contains(word))
            {
                return TokenKind.Builtin;
            }

            var previousWord = CommonScanners.PreviousWord(c, start);
            if (previousWord == "func" || call)
            {
                return TokenKind.Function;
            }

            if (previousWord == "type")
            {
                return TokenKind.Type;
            }

            if (!char.IsUpper(word[0]))
            {
                return TokenKind.Plain;
            }

            // http.Client{...} 这种带包名的复合字面量；其它 pkg.Name 可能是变量、常量，不猜
            if (afterMember)
            {
                return c.NextNonWhitespaceIs('{') ? TokenKind.Type : TokenKind.Plain;
            }

            // 结构体字面量的键 Name: "x"
            if (c.NextNonWhitespaceIs(':'))
            {
                return TokenKind.Plain;
            }

            // 行首的大写名字后面还跟着东西，是结构体字段或接口方法（Name string），不是类型
            if (atLineStart && FollowedByTypeOnSameLine(c))
            {
                return TokenKind.Plain;
            }

            return CommonScanners.IsAllCaps(word) ? TokenKind.Constant : TokenKind.Type;
        }

        /// <summary>当前位置后面隔着空白、同一行里还有名字、* 或 [：字段声明 Name string、Next *Node、Items []Item。</summary>
        private static bool FollowedByTypeOnSameLine(LexerCursor c)
        {
            var i = c.Position;
            while (i < c.Source.Length && (c.Source[i] == ' ' || c.Source[i] == '\t'))
            {
                i++;
            }

            if (i == c.Position || i >= c.Source.Length)
            {
                return false;
            }

            var next = c.Source[i];
            return IsIdentifierStart(next) || next == '*' || next == '[';
        }

        private static bool IsIdentifierStart(char c) => char.IsLetter(c) || c == '_';

        private static bool IsIdentifierPart(char c) => char.IsLetterOrDigit(c) || c == '_';

        public int ScoreLikelihood(DetectionSample sample)
        {
            var code = sample.Code;
            var score = 0;

            // package main：Java 带分号，Kotlin 的包名带点
            score += 6 * Count(code, @"^package[^\S\r\n]+\w+[^\S\r\n]*$");
            score += 4 * Count(code, @"^import[^\S\r\n]*\([^\S\r\n]*$|^import[^\S\r\n]+(\w+[^\S\r\n]+)?""");
            score += 5 * Count(code, @"^func[^\S\r\n]+(\([^()\r\n]*\)[^\S\r\n]*)?\w+[^\S\r\n]*(\[[^\]\r\n]*\])?\(");
            score += 5 * Count(code, @"\berr[^\S\r\n]*!=[^\S\r\n]*nil\b");
            score += 5 * Count(code, @"^[^\S\r\n]*type[^\S\r\n]+\w+[^\S\r\n]+(struct|interface)[^\S\r\n]*\{");
            score += 4 * Count(code, @"\bfmt\.(Print|Sprint|Fprint|Errorf)\w*\(");
            score += 3 * Count(code, @"\bgo[^\S\r\n]+func\b|^[^\S\r\n]*defer[^\S\r\n]+\w");
            score += 3 * Count(code, @"\bmake\((\[\]|map\[|chan\b)|\bchan[^\S\r\n]+\w|<-");
            score += 2 * Count(code, @"\w[^\S\r\n]*:=");
            score += 2 * Count(code, @"\bfunc[^\S\r\n]*\(|\[\][\w*]|\bmap\[\w+\]");
            score += 2 * Count(code, @"\bfor[^\S\r\n]+[^;\r\n]*\brange\b");
            return score;
        }

        private static int Count(string source, string pattern)
        {
            return LikelihoodPatterns.Count(source, pattern, RegexOptions.Multiline);
        }
    }
}
