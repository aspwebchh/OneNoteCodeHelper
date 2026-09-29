using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OneNoteCodeHelper.Highlighting.Languages
{
    /// <summary>
    /// TypeScript 词法着色。词法和 JavaScript 完全一样，只是多认 type、readonly 这类上下文关键字，
    /// string、number 这些基本类型按类型着色。TSX 也按这里处理。
    /// </summary>
    internal sealed class TypeScriptLanguage : ILanguage
    {
        private static readonly JavaScriptLanguage JavaScript = new JavaScriptLanguage();

        public string Id => "typescript";

        public string DisplayName => "TypeScript";

        public IEnumerable<Token> Tokenize(string source)
        {
            return JavaScriptLanguage.Tokenize(source, this, typeScript: true);
        }

        /// <summary>
        /// JavaScript 的分数加上 TS 独有的写法。这样 TS 永远不低于 JS：出现 TS 写法就认成 TS，
        /// 一条没有就和 JS 打平，按登记顺序取 JS。两者在同一族，打平不会变成「无法确定」。
        /// </summary>
        public int ScoreLikelihood(DetectionSample sample)
        {
            var code = sample.Code;
            var score = JavaScript.ScoreLikelihood(sample);

            // 类型标注跟在名字后面，后面接 ; , ) = { |。行尾结束的不算，那是 YAML 的 type: string
            score += 3 * Count(code,
                @"[\w)?][^\S\r\n]*:[^\S\r\n]*(string|number|boolean|any|void|unknown|never)(\[\])?[^\S\r\n]*[;,)={|]");
            // 不含 string：C# 的 List<string> 也长这样
            score += 2 * Count(code, @"<(number|boolean|void|any|unknown)[,>]");
            score += 3 * Count(code, @"^[^\S\r\n]*(export\s+)?type\s+\w+(<[^>\r\n]*>)?\s*=");
            score += 4 * Count(code, @"^[^\S\r\n]*(import|export)\s+type\b");
            score += 4 * Count(code, @"\bdeclare\s+(module|global|const|let|function|class|namespace)\b");
            score += 3 * Count(code, @"\bas\s+const\b|\bkeyof\s+\w|\b(Partial|Required|Readonly|Record|Pick|Omit|ReturnType)<");

            // 带类型标注的类成员：private orders: Order[]。Java 写成 private Order[] orders
            score += 3 * Count(code, @"^[^\S\r\n]*(private|public|protected|readonly)\s+(readonly\s+)?\w+[?!]?\s*:");

            // 接口成员和返回值用自定义类型：items: Item[];、): Order | undefined {
            score += 2 * Count(code, @"^[^\S\r\n]*\w+\??[^\S\r\n]*:[^\S\r\n]*[A-Z]\w*(<[^>\r\n]*>)?(\[\])?[^\S\r\n]*;");
            score += 2 * Count(code,
                @"\)[^\S\r\n]*:[^\S\r\n]*[A-Z]\w*(<[^>\r\n]*>)?(\[\])?[^\S\r\n]*(\||\{|=>)");
            return score;
        }

        private static int Count(string source, string pattern)
        {
            return LikelihoodPatterns.Count(source, pattern, RegexOptions.Multiline);
        }
    }
}
