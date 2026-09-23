using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OneNoteCodeHelper.Highlighting.Languages
{
    /// <summary>
    /// HTML 词法着色，实际切分由 <see cref="MarkupLexer"/> 完成。
    /// &lt;style&gt; 里的内容按 CSS 着色，&lt;script&gt; 里的内容按 JavaScript 着色。
    /// </summary>
    internal sealed class HtmlLanguage : ILanguage
    {
        public string Id => "html";

        public string DisplayName => "HTML";

        public IEnumerable<Token> Tokenize(string source)
        {
            return MarkupLexer.Tokenize(source, html: true);
        }

        public int ScoreLikelihood(string source)
        {
            var score = 0;
            score += 10 * Count(source, @"<!DOCTYPE\s+html", RegexOptions.IgnoreCase);

            // 只认小写的常见 HTML 标签名：XAML、Android 布局这类 XML 用的是 PascalCase（<Button>、<Label>），
            // 不能让它们被算成 HTML。标签名后面必须是空白、/ 或 >，免得 sort <input.txt 这种重定向也被算进来。
            score += 3 * Count(source,
                @"</?(html|head|body|div|span|p|a|ul|ol|li|table|thead|tbody|tr|td|th|script|style|link|meta|title" +
                @"|h[1-6]|img|br|hr|form|input|button|label|select|option|textarea|nav|header|footer|section" +
                @"|article|main|aside|iframe|pre|code|strong|em|b|i|small|canvas|svg|video|audio|source)(?=[\s/>])",
                RegexOptions.None);
            return score;
        }

        private static int Count(string source, string pattern, RegexOptions options)
        {
            return LikelihoodPatterns.Count(source, pattern, RegexOptions.Multiline | options);
        }
    }
}
