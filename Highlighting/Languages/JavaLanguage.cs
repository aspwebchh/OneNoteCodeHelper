using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OneNoteCodeHelper.Highlighting.Languages
{
    /// <summary>Java 词法着色。纯词法、不做语法分析，类型与方法名靠启发式判定。</summary>
    internal sealed class JavaLanguage : ILanguage
    {
        private static readonly HashSet<string> Keywords = new HashSet<string>(StringComparer.Ordinal)
        {
            "abstract", "assert", "boolean", "break", "byte", "case", "catch", "char", "class", "const",
            "continue", "default", "do", "double", "else", "enum", "extends", "final", "finally", "float",
            "for", "goto", "if", "implements", "import", "instanceof", "int", "interface", "long", "native",
            "new", "package", "private", "protected", "public", "return", "short", "static", "strictfp",
            "super", "switch", "synchronized", "this", "throw", "throws", "transient", "try", "void",
            "volatile", "while",
            // 上下文关键字（Java 10 起陆续引入）
            "var", "yield", "record", "sealed", "permits", "module", "requires", "exports", "opens",
            "provides", "uses"
        };

        private static readonly HashSet<string> Literals = new HashSet<string>(StringComparer.Ordinal)
        {
            "true", "false", "null"
        };

        private const string PunctuationChars = "(){}[];,.";

        private const string NumberSuffixChars = "fFdDlL";

        public string Id => "java";

        public string DisplayName => "Java";

        public IEnumerable<Token> Tokenize(string source)
        {
            var c = new LexerCursor(source);

            while (!c.AtEnd)
            {
                var start = c.Position;
                var ch = c.Current;

                // 块注释与文档注释。注意 /**/ 是空的块注释而非文档注释，必须先排除。
                if (ch == '/' && c.Peek() == '*')
                {
                    var isDoc = c.Peek(2) == '*' && c.Peek(3) != '/';
                    c.Advance(2);
                    while (!c.AtEnd && !(c.Current == '*' && c.Peek() == '/'))
                    {
                        c.Advance();
                    }

                    c.Advance(2); // 吃掉结尾 */；未闭合时 Advance 会自己停在末尾
                    c.Emit(isDoc ? TokenKind.DocComment : TokenKind.Comment, start);
                    continue;
                }

                if (ch == '/' && c.Peek() == '/')
                {
                    c.SkipToLineEnd();
                    c.Emit(TokenKind.Comment, start);
                    continue;
                }

                // 文本块（Java 15+），可跨行
                if (ch == '"' && c.Peek() == '"' && c.Peek(2) == '"')
                {
                    ScanTextBlock(c);
                    c.Emit(TokenKind.String, start);
                    continue;
                }

                if (ch == '"')
                {
                    ScanQuoted(c, '"');
                    c.Emit(TokenKind.String, start);
                    continue;
                }

                if (ch == '\'')
                {
                    ScanQuoted(c, '\'');
                    c.Emit(TokenKind.Char, start);
                    continue;
                }

                // 注解：@ 后面必须跟标识符，否则只是个普通符号
                if (ch == '@' && LexerCursor.IsIdentifierStart(c.Peek()))
                {
                    c.Advance();
                    c.SkipWhile(LexerCursor.IsIdentifierPart);
                    c.Emit(TokenKind.Annotation, start);
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

            if (c.NextNonWhitespaceIs('('))
            {
                return TokenKind.Function;
            }

            if (!char.IsUpper(word[0]))
            {
                return TokenKind.Plain;
            }

            // 全大写（允许下划线和数字）按常量看，如 MAX_SIZE；其余大写开头的按类型看。
            // 这两条启发式对 Java 的命名习惯命中率很高。
            return IsAllCaps(word) ? TokenKind.Constant : TokenKind.Type;
        }

        private static bool IsAllCaps(string word)
        {
            if (word.Length < 2)
            {
                return false;
            }

            foreach (var ch in word)
            {
                if (char.IsLower(ch))
                {
                    return false;
                }
            }

            return true;
        }

        private static void ScanTextBlock(LexerCursor c)
        {
            c.Advance(3);
            while (!c.AtEnd)
            {
                if (c.Current == '\\')
                {
                    c.Advance(2);
                    continue;
                }

                if (c.Current == '"' && c.Peek() == '"' && c.Peek(2) == '"')
                {
                    c.Advance(3);
                    return;
                }

                c.Advance();
            }
        }

        /// <summary>吃掉一段引号字面量。遇到换行就停，避免一个未闭合的引号吞掉整个文件。</summary>
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
                c.SkipWhile(ch => LexerCursor.IsHexDigit(ch) || ch == '_');
            }
            else if (c.Current == '0' && (c.Peek() == 'b' || c.Peek() == 'B'))
            {
                c.Advance(2);
                c.SkipWhile(ch => ch == '0' || ch == '1' || ch == '_');
            }
            else
            {
                c.SkipWhile(ch => char.IsDigit(ch) || ch == '_');

                if (c.Current == '.' && char.IsDigit(c.Peek()))
                {
                    c.Advance();
                    c.SkipWhile(ch => char.IsDigit(ch) || ch == '_');
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

            if (NumberSuffixChars.IndexOf(c.Current) >= 0)
            {
                c.Advance();
            }
        }

        public int ScoreLikelihood(DetectionSample sample)
        {
            var code = sample.Code;
            var score = CFamilyFeatures.Score(sample);
            score += 4 * Count(code, @"^[^\S\r\n]*import\s+(static\s+)?[\w.]+(\.\*)?\s*;");
            score += 4 * Count(code, @"^[^\S\r\n]*package\s+[\w.]+\s*;");
            score += 2 * Count(code, @"^[^\S\r\n]*@[A-Z]\w*");
            score += 4 * Count(code, @"\bSystem\.out\.print");
            score += 4 * Count(code, @"\bstatic\s+void\s+main\s*\(\s*String");
            score += 3 * Count(code, @"\bboolean\b|\bthrows\s+[A-Z]");
            score += 2 * Count(code, @"\bString(\[\])?\s+\w+\s*[=;,)]");
            score += 2 * Count(code, @"\bfinal\s+[\w<>\[\]]+\s+\w+\s*[=;]");
            score += 3 * Count(code, @"\b(extends|implements)\s+[A-Z]");

            // 泛型参数只能是包装类型；C# 写 List<string>、C++ 写 vector<int>
            score += 3 * Count(code, @"<(String|Integer|Long|Boolean|Double|Float|Character|Byte|Short|Object)(\[\])?[,>]");
            score += 2 * Count(code, @"\b(ArrayList|HashMap|TreeMap|LinkedHashMap|Optional|Stream)<");

            // 增强 for：C# 是 foreach (x in xs)，JS 是 for (x of xs)
            score += 3 * Count(code, @"\bfor\s*\(\s*(final\s+)?[\w<>\[\],.]+\s+\w+\s*:[^:]");

            // Java 的方法名是 camelCase，和 C# 那条 PascalCase 对称
            score += 3 * Count(code,
                @"\b(public|private|protected)\s+(static\s+|final\s+|abstract\s+|synchronized\s+)*[\w<>\[\],?]+\s+[a-z]\w*\s*\(");
            score += 2 * Count(code, @"\.(equals|equalsIgnoreCase|isEmpty|getClass|hashCode|size)\(");
            return score;
        }

        private static int Count(string source, string pattern)
        {
            return LikelihoodPatterns.Count(source, pattern, RegexOptions.Multiline);
        }
    }
}
