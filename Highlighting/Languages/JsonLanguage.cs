using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OneNoteCodeHelper.Highlighting.Languages
{
    /// <summary>
    /// JSON 词法着色。顺带兼容 JSONC（tsconfig、VS Code settings.json 里的 // 和 /* */ 注释）
    /// 与 JSON5 的单引号字符串、裸键。字符串后面紧跟冒号的是键，否则是值。
    /// </summary>
    internal sealed class JsonLanguage : ILanguage
    {
        private static readonly HashSet<string> Literals = new HashSet<string>(StringComparer.Ordinal)
        {
            "true", "false", "null"
        };

        private const string PunctuationChars = "{}[],:";

        public string Id => "json";

        public string DisplayName => "JSON";

        public IEnumerable<Token> Tokenize(string source)
        {
            var c = new LexerCursor(source);

            while (!c.AtEnd)
            {
                var start = c.Position;
                var ch = c.Current;

                if (char.IsWhiteSpace(ch))
                {
                    c.SkipWhile(char.IsWhiteSpace);
                    c.Emit(TokenKind.Plain, start);
                    continue;
                }

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

                if (ch == '"' || ch == '\'')
                {
                    CommonScanners.ScanQuoted(c, ch);
                    c.Emit(c.NextNonWhitespaceIs(':') ? TokenKind.Attribute : TokenKind.String, start);
                    continue;
                }

                if (char.IsDigit(ch)
                    || (ch == '.' && char.IsDigit(c.Peek()))
                    || ((ch == '-' || ch == '+') && (char.IsDigit(c.Peek()) || (c.Peek() == '.' && char.IsDigit(c.Peek(2))))))
                {
                    if (ch == '-' || ch == '+')
                    {
                        c.Advance();
                    }

                    CommonScanners.ScanNumber(c, '_');
                    c.Emit(TokenKind.Number, start);
                    continue;
                }

                if (LexerCursor.IsIdentifierStart(ch))
                {
                    c.SkipWhile(LexerCursor.IsIdentifierPart);
                    var word = source.Substring(start, c.Position - start);
                    c.Emit(Literals.Contains(word) ? TokenKind.Literal
                        : word == "NaN" || word == "Infinity" ? TokenKind.Number
                        : c.NextNonWhitespaceIs(':') ? TokenKind.Attribute
                        : TokenKind.Plain, start);
                    continue;
                }

                c.Advance();
                c.Emit(PunctuationChars.IndexOf(ch) >= 0 ? TokenKind.Punctuation : TokenKind.Plain, start);
            }

            return c.Finish();
        }

        public int ScoreLikelihood(DetectionSample sample)
        {
            var source = sample.Raw;

            // 必须以 { 或 [ 开头（前面可以有 // 注释）：JS、Python 里的对象字面量前面总有 const x =、return 之类，
            // 这条把它们挡在外面
            if (!LikelihoodPatterns.IsMatch(source, @"\A\s*(//[^\r\n]*\s*)*[\[{]"))
            {
                return 0;
            }

            // 带引号的键。不锚定行首，压缩成一行的 JSON 也能认出来
            var score = 3 * Count(source, @"""[^""\r\n]*""\s*:");

            // 整段都能按 JSON 切分，[1, 2, 3] 这种没有键的也认；[1, 2].map(...) 里的 .map 切不出来，不算
            if (IsAllJsonTokens(source))
            {
                score += 4;
            }

            return score;
        }

        /// <summary>除空白外，每个 token 都是 JSON 里合法的东西（含 JSONC 注释与 JSON5 的裸键）。</summary>
        private bool IsAllJsonTokens(string source)
        {
            foreach (var token in Tokenize(source))
            {
                if (token.Kind != TokenKind.Plain)
                {
                    continue;
                }

                for (var i = token.Start; i < token.End; i++)
                {
                    if (!char.IsWhiteSpace(source[i]))
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        private static int Count(string source, string pattern)
        {
            return LikelihoodPatterns.Count(source, pattern, RegexOptions.Multiline);
        }
    }
}
