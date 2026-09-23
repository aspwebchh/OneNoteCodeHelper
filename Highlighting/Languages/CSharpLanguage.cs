using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OneNoteCodeHelper.Highlighting.Languages
{
    /// <summary>
    /// C# 词法着色。字符串的变体最多：普通、逐字 @""、插值 $""、原始 """ """ 以及它们的组合；
    /// 插值洞 {expr} 里的表达式按 C# 本身重新着色。类型与方法名的判定沿用 Java 的大小写启发式。
    /// </summary>
    internal sealed class CSharpLanguage : ILanguage
    {
        private static readonly HashSet<string> Keywords = new HashSet<string>(StringComparer.Ordinal)
        {
            "abstract", "as", "base", "bool", "break", "byte", "case", "catch", "char", "checked", "class",
            "const", "continue", "decimal", "default", "delegate", "do", "double", "else", "enum", "event",
            "explicit", "extern", "finally", "fixed", "float", "for", "foreach", "goto", "if", "implicit", "in",
            "int", "interface", "internal", "is", "lock", "long", "namespace", "new", "object", "operator", "out",
            "override", "params", "private", "protected", "public", "readonly", "ref", "return", "sbyte",
            "sealed", "short", "sizeof", "stackalloc", "static", "string", "struct", "switch", "this", "throw",
            "try", "typeof", "uint", "ulong", "unchecked", "unsafe", "ushort", "using", "virtual", "void",
            "volatile", "while",
            // 上下文关键字里常见、又很少被拿来当普通标识符的那些
            "async", "await", "var", "dynamic", "nameof", "partial", "record", "get", "set", "init", "when",
            "where", "yield", "global", "required", "scoped", "with", "nint", "nuint", "and", "or", "not"
        };

        private static readonly HashSet<string> Literals = new HashSet<string>(StringComparer.Ordinal)
        {
            "true", "false", "null"
        };

        /// <summary>这些关键字后面紧跟的名字是类型：class Foo、new Foo()、struct Foo……</summary>
        private static readonly HashSet<string> TypeIntroducers = new HashSet<string>(StringComparer.Ordinal)
        {
            "class", "struct", "interface", "enum", "record", "new", "namespace"
        };

        /// <summary>#region Foo 这类指令后面到行尾是说明文字，不按代码着色。</summary>
        private static readonly HashSet<string> TextDirectives = new HashSet<string>(StringComparer.Ordinal)
        {
            "region", "endregion", "pragma", "error", "warning", "line", "nullable"
        };

        private const string PunctuationChars = "(){}[];,.";

        public string Id => "csharp";

        public string DisplayName => "C#";

        public IEnumerable<Token> Tokenize(string source)
        {
            var c = new LexerCursor(source);

            while (!c.AtEnd)
            {
                var start = c.Position;
                var ch = c.Current;

                if (ch == '/' && c.Peek() == '/')
                {
                    // /// 是 XML 文档注释，//// 这种分隔线不算
                    var isDoc = c.Peek(2) == '/' && c.Peek(3) != '/';
                    c.SkipToLineEnd();
                    c.Emit(isDoc ? TokenKind.DocComment : TokenKind.Comment, start);
                    continue;
                }

                if (ch == '/' && c.Peek() == '*')
                {
                    var isDoc = CommonScanners.IsDocBlockComment(c);
                    CommonScanners.ScanBlockComment(c);
                    c.Emit(isDoc ? TokenKind.DocComment : TokenKind.Comment, start);
                    continue;
                }

                // 预处理指令 #if / #region，只在行首成立
                if (ch == '#' && c.IsAtLineStart() && char.IsLetter(c.Peek()))
                {
                    c.Advance();
                    c.SkipWhile(char.IsLetter);
                    var directive = source.Substring(start + 1, c.Position - start - 1);
                    c.Emit(TokenKind.Annotation, start);

                    if (TextDirectives.Contains(directive))
                    {
                        var textStart = c.Position;
                        c.SkipToLineEnd();
                        c.Emit(TokenKind.Plain, textStart);
                    }

                    continue;
                }

                if ((ch == '"' || ch == '$' || ch == '@') && TryScanString(c))
                {
                    continue;
                }

                if (ch == '\'')
                {
                    CommonScanners.ScanQuoted(c, '\'');
                    c.Emit(TokenKind.Char, start);
                    continue;
                }

                // 行首的 [Attribute] / [assembly: Attribute]
                if (ch == '[' && c.IsAtLineStart() && char.IsLetter(c.Peek()))
                {
                    c.Advance();
                    c.Emit(TokenKind.Punctuation, start);
                    var nameStart = c.Position;
                    c.SkipWhile(x => IsIdentifierPart(x) || x == '.');
                    c.Emit(TokenKind.Annotation, nameStart);
                    continue;
                }

                if (char.IsDigit(ch) || (ch == '.' && char.IsDigit(c.Peek())))
                {
                    CommonScanners.ScanNumber(c, '_');
                    c.Emit(TokenKind.Number, start);
                    continue;
                }

                // @class 这种逐字标识符
                if (ch == '@' && IsIdentifierStart(c.Peek()))
                {
                    c.Advance();
                    c.SkipWhile(IsIdentifierPart);
                    c.Emit(TokenKind.Plain, start);
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

            if (Literals.Contains(word))
            {
                return TokenKind.Literal;
            }

            if (TypeIntroducers.Contains(CommonScanners.PreviousWord(c, start)))
            {
                return TokenKind.Type;
            }

            if (c.NextNonWhitespaceIs('('))
            {
                return TokenKind.Function;
            }

            // 成员访问 obj.Name：C# 的属性也是大写开头，不能按类型着色
            if (afterMember || !char.IsUpper(word[0]))
            {
                return TokenKind.Plain;
            }

            // 属性声明 Name { get; }、赋值 Name = x、Name => ... 也是大写开头的成员，不是类型
            if (c.NextNonWhitespaceIs('{') || c.NextNonWhitespaceIs('=') || c.NextNonWhitespaceIs(';'))
            {
                return TokenKind.Plain;
            }

            return CommonScanners.IsAllCaps(word) ? TokenKind.Constant : TokenKind.Type;
        }

        /// <summary>
        /// 识别各种字符串前缀：$、@、$@、@$、$$、原始字符串 """。当前位置不是字符串开头就原地不动返回 false。
        /// </summary>
        private bool TryScanString(LexerCursor c)
        {
            var start = c.Position;
            var prefix = 0;
            var dollars = 0;
            var verbatim = false;

            while (c.Peek(prefix) == '$')
            {
                dollars++;
                prefix++;
            }

            if (c.Peek(prefix) == '@')
            {
                verbatim = true;
                prefix++;
                while (c.Peek(prefix) == '$')
                {
                    dollars++;
                    prefix++;
                }
            }

            if (c.Peek(prefix) != '"')
            {
                return false;
            }

            var quotes = 0;
            while (c.Peek(prefix + quotes) == '"')
            {
                quotes++;
            }

            // 原始字符串 """...""" (C# 11)，引号数可以更多，结束符与开头引号数相同
            if (!verbatim && quotes >= 3)
            {
                var terminator = new string('"', quotes);
                c.Advance(prefix + quotes);
                if (dollars > 0)
                {
                    CommonScanners.ScanInterpolatedBody(
                        c, this, start, terminator, '\0', new string('{', dollars), multiLine: true);
                    return true;
                }

                while (!c.AtEnd && !c.Matches(terminator))
                {
                    c.Advance();
                }

                c.Advance(quotes);
                c.Emit(TokenKind.String, start);
                return true;
            }

            if (!verbatim && dollars == 0)
            {
                CommonScanners.ScanQuoted(c, '"');
                c.Emit(TokenKind.String, start);
                return true;
            }

            c.Advance(prefix + 1);

            if (dollars > 0)
            {
                CommonScanners.ScanInterpolatedBody(
                    c, this, start, "\"", verbatim ? '\0' : '\\', "{", multiLine: verbatim, doubledTerminator: verbatim);
                return true;
            }

            // 逐字字符串：没有转义，"" 表示一个引号，可以跨行
            while (!c.AtEnd)
            {
                if (c.Current == '"')
                {
                    if (c.Peek() == '"')
                    {
                        c.Advance(2);
                        continue;
                    }

                    c.Advance();
                    break;
                }

                c.Advance();
            }

            c.Emit(TokenKind.String, start);
            return true;
        }

        private static bool IsIdentifierStart(char c) => char.IsLetter(c) || c == '_';

        private static bool IsIdentifierPart(char c) => char.IsLetterOrDigit(c) || c == '_';

        public int ScoreLikelihood(DetectionSample sample)
        {
            var code = sample.Code;
            var score = CFamilyFeatures.Score(sample);
            score += 4 * Count(code, @"^[^\S\r\n]*using\s+(static\s+)?[\w.]+\s*;");
            score += 4 * Count(code, @"^[^\S\r\n]*namespace\s+[A-Z][\w.]*\s*[;{]?\s*$");
            score += 4 * Count(code, @"\{\s*get\s*;|\bget\s*\{|\bset\s*\{|\bget\s*=>");
            score += 3 * Count(code, @"\bConsole\.(Write|Read)");
            score += 3 * Count(code, @"\bforeach\s*\(\s*var\b");
            score += 3 * Count(code, @"\b(async\s+Task|Task<|IEnumerable<|IReadOnly\w*<|IList<|Dictionary<)");
            score += 3 * Count(code, @"\b(internal|sealed|readonly|nameof)\b");
            score += 3 * Count(code, @"\bis\s+(not\s+)?null\b");

            // C# 的方法名是 PascalCase，Java 是 camelCase
            score += 3 * Count(code,
                @"\b(public|private|protected|internal)\s+(static\s+|override\s+|virtual\s+|async\s+)*[\w<>\[\],?]+\s+[A-Z]\w*\s*\(");
            score += 2 * Count(code, CFamilyFeatures.PrimitiveTypeArgument);
            score += 2 * Count(code,
                @"\.(Where|Select|SelectMany|OrderBy|OrderByDescending|GroupBy|ToList|ToArray|ToDictionary" +
                @"|FirstOrDefault|SingleOrDefault|LastOrDefault)\s*\(");
            score += 2 * Count(code, @"\bstring\s+\w+\s*[=;,)]|\bstring\.[A-Z]");
            score += 2 * Count(code, @"^[^\S\r\n]*\[[A-Z]\w*(\(.*\))?\]\s*$");
            score += 2 * Count(sample.Raw, @"\$@?""");
            score += 2 * Count(code, @"\boverride\b");
            score += 2 * Count(code, @"^[^\S\r\n]*#(region|endregion)\b");
            score += 2 * Count(sample.Raw, @"^[^\S\r\n]*///");
            score += Count(code, @"\.[A-Z][a-z]\w*\(");
            score += Count(code, @"\bvar\s+\w+\s*=");
            score += Count(code, @"\bbool\b");
            return score;
        }

        private static int Count(string source, string pattern)
        {
            return LikelihoodPatterns.Count(source, pattern, RegexOptions.Multiline);
        }
    }
}
