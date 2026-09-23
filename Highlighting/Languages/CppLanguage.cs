using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OneNoteCodeHelper.Highlighting.Languages
{
    /// <summary>
    /// C 与 C++ 共用一套词法着色（C 基本是 C++ 词法的子集）。
    /// 要特别处理的是行首的预处理指令、#include &lt;...&gt;、原始字符串 R"x(...)x" 以及 C++14 的数字分隔符 1'000。
    /// </summary>
    internal sealed class CppLanguage : ILanguage
    {
        private static readonly HashSet<string> Keywords = new HashSet<string>(StringComparer.Ordinal)
        {
            // C
            "auto", "break", "case", "char", "const", "continue", "default", "do", "double", "else", "enum",
            "extern", "float", "for", "goto", "if", "inline", "int", "long", "register", "restrict", "return",
            "short", "signed", "sizeof", "static", "struct", "switch", "typedef", "union", "unsigned", "void",
            "volatile", "while", "_Alignas", "_Alignof", "_Atomic", "_Bool", "_Complex", "_Generic",
            "_Noreturn", "_Static_assert", "_Thread_local",
            // C++
            "alignas", "alignof", "asm", "bool", "catch", "char8_t", "char16_t", "char32_t", "class", "concept",
            "consteval", "constexpr", "constinit", "const_cast", "co_await", "co_return", "co_yield", "decltype",
            "delete", "dynamic_cast", "explicit", "export", "final", "friend", "mutable", "namespace", "new",
            "noexcept", "operator", "override", "private", "protected", "public", "reinterpret_cast",
            "requires", "static_assert", "static_cast", "template", "this", "thread_local", "throw", "try",
            "typeid", "typename", "using", "virtual", "wchar_t"
        };

        private static readonly HashSet<string> Literals = new HashSet<string>(StringComparer.Ordinal)
        {
            "true", "false", "nullptr", "NULL"
        };

        private static readonly HashSet<string> Builtins = new HashSet<string>(StringComparer.Ordinal)
        {
            "std", "cout", "cin", "cerr", "clog", "endl"
        };

        /// <summary>
        /// 标准库的模板类型。list、map 这类名字也常被当成普通变量名，
        /// 所以只在 std:: 之后或紧跟 &lt; 时才按类型着色。
        /// </summary>
        private static readonly HashSet<string> StandardTemplates = new HashSet<string>(StringComparer.Ordinal)
        {
            "string", "wstring", "string_view", "vector", "map", "multimap", "set", "multiset", "unordered_map",
            "unordered_set", "list", "forward_list", "deque", "array", "pair", "tuple", "stack", "queue",
            "priority_queue", "shared_ptr", "unique_ptr", "weak_ptr", "optional", "variant", "function",
            "thread", "mutex", "atomic", "span", "basic_string", "initializer_list"
        };

        private static readonly HashSet<string> TypeIntroducers = new HashSet<string>(StringComparer.Ordinal)
        {
            "class", "struct", "union", "enum", "typename", "namespace", "new"
        };

        private static readonly HashSet<string> StringPrefixes = new HashSet<string>(StringComparer.Ordinal)
        {
            "L", "u", "U", "u8", "R", "LR", "uR", "UR", "u8R"
        };

        private const string PunctuationChars = "(){}[];,.";

        public string Id => "cpp";

        public string DisplayName => "C/C++";

        public IEnumerable<Token> Tokenize(string source)
        {
            var c = new LexerCursor(source);

            while (!c.AtEnd)
            {
                var start = c.Position;
                var ch = c.Current;

                if (ch == '/' && c.Peek() == '/')
                {
                    // /// 与 //! 是 Doxygen 文档注释
                    var isDoc = (c.Peek(2) == '/' && c.Peek(3) != '/') || c.Peek(2) == '!';
                    c.SkipToLineEnd();
                    c.Emit(isDoc ? TokenKind.DocComment : TokenKind.Comment, start);
                    continue;
                }

                if (ch == '/' && c.Peek() == '*')
                {
                    var isDoc = CommonScanners.IsDocBlockComment(c) || c.Peek(2) == '!';
                    CommonScanners.ScanBlockComment(c);
                    c.Emit(isDoc ? TokenKind.DocComment : TokenKind.Comment, start);
                    continue;
                }

                // 预处理指令：# 必须在行首，# 和指令名之间允许空白
                if (ch == '#' && c.IsAtLineStart())
                {
                    c.Advance();
                    c.SkipWhile(x => x == ' ' || x == '\t');
                    var nameStart = c.Position;
                    c.SkipWhile(char.IsLetter);
                    var directive = source.Substring(nameStart, c.Position - nameStart);
                    c.Emit(TokenKind.Annotation, start);

                    // #include <stdio.h> 里的尖括号路径按字符串着色
                    if (directive == "include" || directive == "import" || directive == "include_next")
                    {
                        c.SkipWhile(x => x == ' ' || x == '\t');
                        if (c.Current == '<')
                        {
                            var pathStart = c.Position;
                            c.SkipWhile(x => x != '>' && x != '\n' && x != '\r');
                            c.Advance();
                            c.Emit(TokenKind.String, pathStart);
                        }
                    }

                    continue;
                }

                if (ch == '"')
                {
                    CommonScanners.ScanQuoted(c, '"');
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
                    CommonScanners.ScanNumber(c, '\'');
                    c.Emit(TokenKind.Number, start);
                    continue;
                }

                if (char.IsLetter(ch) || ch == '_')
                {
                    c.SkipWhile(x => char.IsLetterOrDigit(x) || x == '_');
                    var word = source.Substring(start, c.Position - start);

                    // L"..."、u8"..."、R"(...)" 这类带前缀的字符串与字符
                    if ((c.Current == '"' || c.Current == '\'') && StringPrefixes.Contains(word))
                    {
                        ScanPrefixedLiteral(c, start, word);
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
            var previous = start > 0 ? c.Source[start - 1] : '\0';
            var afterMember = previous == '.' || (previous == '>' && start > 1 && c.Source[start - 2] == '-');
            var afterScope = previous == ':' && start > 1 && c.Source[start - 2] == ':';

            if (!afterMember && Keywords.Contains(word))
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

            if (StandardTemplates.Contains(word) && (afterScope || c.Current == '<'))
            {
                return TokenKind.Type;
            }

            if (TypeIntroducers.Contains(CommonScanners.PreviousWord(c, start)))
            {
                return TokenKind.Type;
            }

            if (c.NextNonWhitespaceIs('('))
            {
                return TokenKind.Function;
            }

            if (afterMember)
            {
                return TokenKind.Plain;
            }

            // size_t、uint32_t、pthread_t 这类 _t 结尾的类型
            if (word.Length > 2 && word.EndsWith("_t", StringComparison.Ordinal))
            {
                return TokenKind.Type;
            }

            if (CommonScanners.IsAllCaps(word))
            {
                return TokenKind.Constant;
            }

            return char.IsUpper(word[0]) ? TokenKind.Type : TokenKind.Plain;
        }

        /// <summary>前缀已经吃掉，当前位置在引号上。R 开头的是原始字符串 R"delim(...)delim"，可以跨行。</summary>
        private static void ScanPrefixedLiteral(LexerCursor c, int start, string prefix)
        {
            if (c.Current == '\'')
            {
                CommonScanners.ScanQuoted(c, '\'');
                c.Emit(TokenKind.Char, start);
                return;
            }

            if (!prefix.EndsWith("R", StringComparison.Ordinal))
            {
                CommonScanners.ScanQuoted(c, '"');
                c.Emit(TokenKind.String, start);
                return;
            }

            // 分隔符最长 16 个字符，不能含空白和括号
            var length = 1;
            while (length <= 17 && c.Peek(length) != '(' && c.Peek(length) != '\0'
                   && !char.IsWhiteSpace(c.Peek(length)) && c.Peek(length) != ')')
            {
                length++;
            }

            if (c.Peek(length) != '(')
            {
                CommonScanners.ScanQuoted(c, '"');
                c.Emit(TokenKind.String, start);
                return;
            }

            var terminator = ")" + c.Source.Substring(c.Position + 1, length - 1) + "\"";
            c.Advance(length + 1);
            while (!c.AtEnd && !c.Matches(terminator))
            {
                c.Advance();
            }

            c.Advance(terminator.Length);
            c.Emit(TokenKind.String, start);
        }

        public int ScoreLikelihood(DetectionSample sample)
        {
            var code = sample.Code;
            var score = CFamilyFeatures.Score(sample);
            score += 5 * Count(code, @"^[^\S\r\n]*#\s*include\s*[<""]");
            score += 3 * Count(code, @"^[^\S\r\n]*#\s*(define|ifdef|ifndef|endif|pragma|undef)\b");
            score += 5 * Count(code, @"\bstd::");
            score += 3 * Count(code, @"\b(cout|cerr)\s*<<|\bcin\s*>>");
            score += 4 * Count(code, @"\btemplate\s*<");
            score += 4 * Count(code, @"\bint\s+main\s*\(");
            score += 3 * Count(code, @"\b(nullptr|typedef|constexpr)\b");
            score += 2 * Count(code, @"\b(printf|scanf|malloc|calloc|free|memcpy|memset|strlen|strcpy)\s*\(");
            score += 2 * Count(code, @"\b(unsigned|size_t|u?int\d+_t|const\s+char\s*\*)");
            score += 2 * Count(code, @"\b\w+::~?\w+\s*\(");
            score += 2 * Count(code, @"\bstruct\s+\w+\s*\*");
            score += 3 * Count(code, @"^[^\S\r\n]*(public|private|protected)\s*:");
            score += 2 * Count(code, @"\bconst\s+[\w:<>]+\s*&");
            score += 2 * Count(code, CFamilyFeatures.PrimitiveTypeArgument);
            score += Count(code, @"\w->\w");
            return score;
        }

        private static int Count(string source, string pattern)
        {
            return LikelihoodPatterns.Count(source, pattern, RegexOptions.Multiline);
        }
    }
}
