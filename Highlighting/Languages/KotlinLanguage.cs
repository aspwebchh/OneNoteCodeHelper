using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OneNoteCodeHelper.Highlighting.Languages
{
    /// <summary>
    /// Kotlin 词法着色。字符串模板有 $name 和 ${expr} 两种写法，三引号原始字符串里也能用；块注释可以嵌套。
    /// 修饰符（data、open、override……）是软关键字，后面紧跟声明时才按关键字着色。类型与常量沿用 Java 的大小写启发式。
    /// </summary>
    internal sealed class KotlinLanguage : ILanguage
    {
        private static readonly HashSet<string> Keywords = new HashSet<string>(StringComparer.Ordinal)
        {
            "as", "break", "class", "continue", "do", "else", "for", "fun", "if", "in", "interface", "is",
            "object", "package", "return", "super", "this", "throw", "try", "typealias", "typeof", "val", "var",
            "when", "while", "import", "catch", "finally", "init", "constructor", "where"
        };

        /// <summary>修饰符：后面紧跟名字时才是关键字（data class、private val），val data = ... 里只是变量名。</summary>
        private static readonly HashSet<string> Modifiers = new HashSet<string>(StringComparer.Ordinal)
        {
            "abstract", "actual", "annotation", "by", "companion", "const", "crossinline", "data", "enum",
            "expect", "external", "final", "infix", "inline", "inner", "internal", "lateinit", "noinline", "open",
            "operator", "out", "override", "private", "protected", "public", "reified", "sealed", "suspend",
            "tailrec", "value", "vararg"
        };

        private static readonly HashSet<string> Literals = new HashSet<string>(StringComparer.Ordinal)
        {
            "true", "false", "null"
        };

        private static readonly HashSet<string> TypeIntroducers = new HashSet<string>(StringComparer.Ordinal)
        {
            "class", "interface", "object", "typealias"
        };

        private const string PunctuationChars = "(){}[];,.";

        public string Id => "kotlin";

        public string DisplayName => "Kotlin";

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
                    var isDoc = CommonScanners.IsDocBlockComment(c);
                    CommonScanners.ScanNestedBlockComment(c);
                    c.Emit(isDoc ? TokenKind.DocComment : TokenKind.Comment, start);
                    continue;
                }

                // 原始字符串 """..."""：没有转义，可以跨行，模板照样生效
                if (c.Matches("\"\"\""))
                {
                    c.Advance(3);
                    CommonScanners.ScanInterpolatedBody(
                        c, this, start, "\"\"\"", '\0', "${", multiLine: true, dollarVariables: true);
                    continue;
                }

                if (ch == '"')
                {
                    c.Advance();
                    CommonScanners.ScanInterpolatedBody(
                        c, this, start, "\"", '\\', "${", multiLine: false, dollarVariables: true);
                    continue;
                }

                if (ch == '\'')
                {
                    CommonScanners.ScanQuoted(c, '\'');
                    c.Emit(TokenKind.Char, start);
                    continue;
                }

                // 注解 @Inject，以及 return@forEach 这种带标签的跳转
                if (ch == '@' && IsIdentifierStart(c.Peek()))
                {
                    c.Advance();
                    c.SkipWhile(IsIdentifierPart);
                    c.Emit(TokenKind.Annotation, start);
                    continue;
                }

                // 反引号括起来的名字：`when`、测试方法名 `returns empty list`
                if (ch == '`')
                {
                    c.Advance();
                    c.SkipWhile(x => x != '`' && x != '\n' && x != '\r');
                    if (c.Current == '`')
                    {
                        c.Advance();
                    }

                    c.Emit(TokenKind.Plain, start);
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
                    c.SkipWhile(IsIdentifierPart);
                    var word = source.Substring(start, c.Position - start);
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

            if (!afterMember && Modifiers.Contains(word) && CommonScanners.IsFollowedByWord(c))
            {
                return TokenKind.Keyword;
            }

            if (!afterMember && Literals.Contains(word))
            {
                return TokenKind.Literal;
            }

            var previousWord = CommonScanners.PreviousWord(c, start);

            // fun foo()，扩展函数 fun String.foo() 里的 String 是接收者类型
            if (previousWord == "fun")
            {
                return c.Current == '.' ? TokenKind.Type : TokenKind.Function;
            }

            if (TypeIntroducers.Contains(previousWord))
            {
                return TokenKind.Type;
            }

            if (c.NextNonWhitespaceIs('('))
            {
                return TokenKind.Function;
            }

            if (!char.IsUpper(word[0]))
            {
                return TokenKind.Plain;
            }

            return CommonScanners.IsAllCaps(word) ? TokenKind.Constant : TokenKind.Type;
        }

        private static bool IsIdentifierStart(char c) => char.IsLetter(c) || c == '_';

        private static bool IsIdentifierPart(char c) => char.IsLetterOrDigit(c) || c == '_';

        public int ScoreLikelihood(DetectionSample sample)
        {
            var code = sample.Code;
            var score = 0;
            score += 5 * Count(code, @"^[^\S\r\n]*(\w+[^\S\r\n]+)*fun[^\S\r\n]+(<[^>\r\n]*>[^\S\r\n]*)?([\w.]+\.)?\w+[^\S\r\n]*\(");
            score += 4 * Count(code, @"\bval[^\S\r\n]+\w+[^\S\r\n]*(:[^=\r\n]+)?=");
            score += 3 * Count(code, @"\b(val|var)[^\S\r\n]+\w+[^\S\r\n]*:[^\S\r\n]*[A-Z]\w*");
            score += 6 * Count(code, @"\bdata[^\S\r\n]+class\b");
            score += 4 * Count(code, @"\bcompanion[^\S\r\n]+object\b|^[^\S\r\n]*object[^\S\r\n]+\w+");
            score += 4 * Count(code, @"\bwhen[^\S\r\n]*(\([^()\r\n]*\))?[^\S\r\n]*\{");
            score += 3 * Count(code, @"\b(suspend|override|inline|private|internal|open)[^\S\r\n]+fun\b");
            score += 5 * Count(code, @"\blateinit[^\S\r\n]+var\b");
            score += 3 * Count(code, @":[^\S\r\n]*(Int|String|Boolean|Long|Double|Float|Unit|Any|Char)\??[^\S\r\n]*[,)={]");
            score += 3 * Count(code, @"\b(listOf|mutableListOf|mapOf|mutableMapOf|setOf|mutableSetOf|arrayOf|emptyList)[^\S\r\n]*[(<]");
            score += 3 * Count(code, @"^package[^\S\r\n]+\w+(\.\w+)+[^\S\r\n]*$");
            score += 3 * Count(code, @"^import[^\S\r\n]+[a-z]\w*(\.\w+)*\.[A-Z]\w*[^\S\r\n]*$");
            score += 2 * Count(code, @"\s\?:\s|!!\.");
            score += 2 * Count(code, @"(?<![.\w])println[^\S\r\n]*\(");

            // 反证：Kotlin 的语句几乎从不以分号结尾
            score -= Count(code, @";[^\S\r\n]*$");
            return score;
        }

        private static int Count(string source, string pattern)
        {
            return LikelihoodPatterns.Count(source, pattern, RegexOptions.Multiline);
        }
    }
}
