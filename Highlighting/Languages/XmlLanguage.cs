using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OneNoteCodeHelper.Highlighting.Languages
{
    /// <summary>XML 词法着色，实际切分由 <see cref="MarkupLexer"/> 完成。</summary>
    internal sealed class XmlLanguage : ILanguage
    {
        public string Id => "xml";

        public string DisplayName => "XML";

        public IEnumerable<Token> Tokenize(string source)
        {
            return MarkupLexer.Tokenize(source, html: false);
        }

        public int ScoreLikelihood(string source)
        {
            var score = 0;
            score += 10 * Count(source, @"^[^\S\r\n]*<\?xml\b");
            score += 4 * Count(source, @"\bxmlns(:\w+)?\s*=");
            score += 2 * Count(source, @"</[A-Za-z_][\w:.-]*\s*>");
            score += Count(source, @"/>");
            return score;
        }

        private static int Count(string source, string pattern)
        {
            return LikelihoodPatterns.Count(source, pattern, RegexOptions.Multiline);
        }
    }
}
