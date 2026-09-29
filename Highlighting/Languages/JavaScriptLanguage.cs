using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OneNoteCodeHelper.Highlighting.Languages
{
    /// <summary>
    /// JavaScript 词法着色。两个难点：模板字符串 `...${expr}...` 里的表达式要按 JS 重新着色；
    /// / 既可能是除号也可能是正则字面量的开头，要看前一个有意义的符号来判断。
    /// TypeScript 共用这套词法，只在给标识符分类时多认类型相关的写法，见 <see cref="TypeScriptLanguage"/>。
    /// </summary>
    internal sealed class JavaScriptLanguage : ILanguage
    {
        private static readonly HashSet<string> Keywords = new HashSet<string>(StringComparer.Ordinal)
        {
            "break", "case", "catch", "class", "const", "continue", "debugger", "default", "delete", "do",
            "else", "enum", "export", "extends", "finally", "for", "function", "if", "implements", "import",
            "in", "instanceof", "interface", "let", "new", "of", "package", "private", "protected", "public",
            "return", "static", "super", "switch", "this", "throw", "try", "typeof", "var", "void", "while",
            "with", "yield", "async", "await", "from", "as"
        };

        private static readonly HashSet<string> Literals = new HashSet<string>(StringComparer.Ordinal)
        {
            "true", "false", "null", "undefined", "NaN", "Infinity"
        };

        private static readonly HashSet<string> Builtins = new HashSet<string>(StringComparer.Ordinal)
        {
            "console", "window", "document", "globalThis", "process", "require", "module", "exports", "JSON",
            "Math", "arguments"
        };

        /// <summary>这些关键字之后出现的 / 是正则字面量的开头，而不是除号。</summary>
        private static readonly HashSet<string> RegexAfterKeywords = new HashSet<string>(StringComparer.Ordinal)
        {
            "return", "typeof", "instanceof", "in", "of", "new", "delete", "void", "throw", "case", "do", "else",
            "yield", "await"
        };

        private static readonly HashSet<string> TypeIntroducers = new HashSet<string>(StringComparer.Ordinal)
        {
            "class", "new", "extends", "implements", "instanceof"
        };

        /// <summary>TypeScript 的上下文关键字：后面紧跟名字时才算（type Foo、readonly items），type = 1 里只是变量名。</summary>
        private static readonly HashSet<string> TypeScriptContextualKeywords = new HashSet<string>(StringComparer.Ordinal)
        {
            "type", "namespace", "module", "declare", "abstract", "readonly", "keyof", "infer", "is", "asserts",
            "satisfies", "override", "unique", "accessor"
        };

        private static readonly HashSet<string> TypeScriptPrimitiveTypes = new HashSet<string>(StringComparer.Ordinal)
        {
            "string", "number", "boolean", "any", "unknown", "never", "object", "symbol", "bigint"
        };

        private static readonly HashSet<string> TypeScriptTypeIntroducers = new HashSet<string>(StringComparer.Ordinal)
        {
            "interface", "type", "enum", "namespace"
        };

        /// <summary>这些词后面的名字是变量，哪怕叫 number、string：const number = 5。</summary>
        private static readonly HashSet<string> Declarators = new HashSet<string>(StringComparer.Ordinal)
        {
            "const", "let", "var"
        };

        private const string PunctuationChars = "(){}[];,.";

        public string Id => "javascript";

        public string DisplayName => "JavaScript";

        public IEnumerable<Token> Tokenize(string source)
        {
            return Tokenize(source, this, typeScript: false);
        }

        /// <summary>JavaScript 与 TypeScript 共用的词法。self 是模板字符串插值洞里用来重新着色的语言。</summary>
        internal static IEnumerable<Token> Tokenize(string source, ILanguage self, bool typeScript)
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
                    CommonScanners.ScanBlockComment(c);
                    c.Emit(isDoc ? TokenKind.DocComment : TokenKind.Comment, start);
                    continue;
                }

                if (ch == '/' && RegexAllowedAt(c, start))
                {
                    var end = FindRegexEnd(source, start);
                    if (end > 0)
                    {
                        c.Advance(end - start);
                        c.SkipWhile(char.IsLetter); // 标志位 gimsuy
                        c.Emit(TokenKind.String, start);
                        continue;
                    }
                }

                if (ch == '`')
                {
                    c.Advance();
                    CommonScanners.ScanInterpolatedBody(c, self, start, "`", '\\', "${", multiLine: true);
                    continue;
                }

                if (ch == '"' || ch == '\'')
                {
                    CommonScanners.ScanQuoted(c, ch);
                    c.Emit(TokenKind.String, start);
                    continue;
                }

                // 装饰器 @decorator
                if (ch == '@' && LexerCursor.IsIdentifierStart(c.Peek()))
                {
                    c.Advance();
                    c.SkipWhile(x => LexerCursor.IsIdentifierPart(x) || x == '.');
                    c.Emit(TokenKind.Annotation, start);
                    continue;
                }

                if (char.IsDigit(ch) || (ch == '.' && char.IsDigit(c.Peek())))
                {
                    CommonScanners.ScanNumber(c, '_');
                    c.Emit(TokenKind.Number, start);
                    continue;
                }

                if (LexerCursor.IsIdentifierStart(ch))
                {
                    c.SkipWhile(LexerCursor.IsIdentifierPart);
                    var word = source.Substring(start, c.Position - start);
                    c.Emit(ClassifyIdentifier(c, word, start, typeScript), start);
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

        private static TokenKind ClassifyIdentifier(LexerCursor c, string word, int start, bool typeScript)
        {
            var afterMember = start > 0 && c.Source[start - 1] == '.';

            if (!afterMember && Keywords.Contains(word))
            {
                return TokenKind.Keyword;
            }

            if (typeScript && !afterMember && TypeScriptContextualKeywords.Contains(word)
                && CommonScanners.IsFollowedByWord(c))
            {
                return TokenKind.Keyword;
            }

            if (!afterMember && Literals.Contains(word))
            {
                return TokenKind.Literal;
            }

            if (!afterMember && Builtins.Contains(word))
            {
                return TokenKind.Builtin;
            }

            var previousWord = CommonScanners.PreviousWord(c, start);
            if (previousWord == "function")
            {
                return TokenKind.Function;
            }

            if (TypeIntroducers.Contains(previousWord)
                || (typeScript && TypeScriptTypeIntroducers.Contains(previousWord)))
            {
                return TokenKind.Type;
            }

            // string、number 做类型标注；{ string: 1 } 的键、string(x) 这种调用、const number = 5 不算
            if (typeScript && !afterMember && TypeScriptPrimitiveTypes.Contains(word)
                && !c.NextNonWhitespaceIs('(') && !c.NextNonWhitespaceIs(':') && !Declarators.Contains(previousWord))
            {
                return TokenKind.Type;
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
        /// 前一个有意义的符号是运算符、左括号、逗号或 return 这类关键字时，/ 开始的是正则；
        /// 前面是标识符、数字、右括号时是除号。
        /// </summary>
        private static bool RegexAllowedAt(LexerCursor c, int position)
        {
            var previous = CommonScanners.PreviousNonWhitespace(c, position);
            if (previous == '\0')
            {
                return true;
            }

            if (LexerCursor.IsIdentifierPart(previous))
            {
                return RegexAfterKeywords.Contains(CommonScanners.PreviousWord(c, position));
            }

            return previous != ')' && previous != ']' && previous != '"' && previous != '\'' && previous != '`';
        }

        /// <summary>从 / 开始找正则字面量的结束 /（字符类 [...] 里的 / 不算），返回结束 / 之后的位置；同一行找不到返回 -1。</summary>
        private static int FindRegexEnd(string source, int start)
        {
            var inClass = false;
            for (var i = start + 1; i < source.Length; i++)
            {
                var ch = source[i];
                switch (ch)
                {
                    case '\n':
                    case '\r':
                        return -1;
                    case '\\':
                        i++;
                        break;
                    case '[':
                        inClass = true;
                        break;
                    case ']':
                        inClass = false;
                        break;
                    case '/':
                        if (!inClass)
                        {
                            return i == start + 1 ? -1 : i + 1;
                        }

                        break;
                }
            }

            return -1;
        }

        /// <summary>JSX 也按 JavaScript 识别。TypeScript 在这个分数上再加它独有的写法，见 <see cref="TypeScriptLanguage"/>。</summary>
        public int ScoreLikelihood(DetectionSample sample)
        {
            var code = sample.Code;
            var score = CFamilyFeatures.Score(sample);
            score += 10 * Count(sample.Raw, @"\A#!.*\bnode\b");
            score += 3 * Count(code, @"^[^\S\r\n]*(export\s+)?(const|let)\s+[\w${}\[\], ]+\s*=");
            score += 4 * Count(code, @"\bfunction\s*\*?\s*[\w$]*\s*\([^)]*\)\s*\{");
            score += 4 * Count(code, @"\bconsole\.(log|error|warn|info|debug)\s*\(");
            score += 3 * Count(code, @"\b(document|window)\.\w+");
            score += 4 * Count(code,
                @"\brequire\s*\(\s*['""]|\bmodule\.exports\b" +
                @"|^[^\S\r\n]*export\s+(default|const|let|function|class|async|interface|type|enum|abstract)\b" +
                @"|^[^\S\r\n]*import\s+.*\bfrom\s+['""]|^[^\S\r\n]*import\s+['""]");
            score += 3 * Count(code, @"===|!==");
            score += 2 * Count(code, @"\bundefined\b");
            score += 2 * Count(sample.Raw, @"`[^`]*\$\{");
            score += 2 * Count(code, @"\.then\s*\(|\bawait\s+fetch\b|\baddEventListener\s*\(");
            score += Count(code, @"=>");
            score += Count(code, @"\bvar\s+\w+\s*=");

            // class 里不写返回类型的方法：speak() {、async load(id) {。Java/C# 的方法前面总有返回类型；
            // 排除掉的是各语言带括号的语句块（C# 的 foreach/using/lock、Java 的 synchronized）
            score += 3 * Count(code, @"\bconstructor\s*\(");
            score += 2 * Count(code,
                @"^[^\S\r\n]*(async\s+|static\s+|get\s+|set\s+)*" +
                @"(?!(if|for|foreach|while|switch|catch|function|return|else|using|lock|synchronized|fixed|with)\b)" +
                @"[A-Za-z_$][\w$]*[^\S\r\n]*\([^()\r\n]*\)[^\S\r\n]*\{");
            return score;
        }

        private static int Count(string source, string pattern)
        {
            return LikelihoodPatterns.Count(source, pattern, RegexOptions.Multiline);
        }
    }
}
