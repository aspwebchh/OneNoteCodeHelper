using System;
using System.Collections.Generic;
using System.Linq;
using OneNoteCodeHelper.Highlighting.Languages;

namespace OneNoteCodeHelper.Highlighting
{
    /// <summary>
    /// 已支持语言的登记处。加新语言只需在 All 里加一行。
    /// </summary>
    internal static class LanguageRegistry
    {
        /// <summary>「自动识别」在界面与配置里的标识。</summary>
        internal const string AutoDetectId = "auto";

        /// <summary>最高分低于这个数就不认：只命中一两条弱特征（比如一个分号）说明不了什么。</summary>
        private const int MinimumScore = 3;

        /// <summary>顺序即下拉顺序；自动识别打平时也取靠前的那个。</summary>
        internal static IReadOnlyList<ILanguage> All { get; } = new ILanguage[]
        {
            new JavaLanguage(),
            new CSharpLanguage(),
            new CppLanguage(),
            new JavaScriptLanguage(),
            new PythonLanguage(),
            new SqlLanguage(),
            new LuaLanguage(),
            new PowerShellLanguage(),
            new BatLanguage(),
            new BashLanguage(),
            new XmlLanguage(),
            new HtmlLanguage(),
            new CssLanguage(),
            new JsonLanguage(),
            new YamlLanguage(),
            new TextLanguage()
        };

        /// <summary>
        /// 高亮效果几乎一样的语言归成一族：彼此认错的代价很小，打平时不必因此放弃识别。
        /// C 系共用注释、字符串、数字的写法，只是关键字有出入；XML 和 HTML 共用同一个标记词法。
        /// </summary>
        private static readonly string[][] Families =
        {
            new[] { "java", "csharp", "cpp", "javascript" },
            new[] { "xml", "html" }
        };

        private static readonly string[] MarkupFamily = Families[1];

        internal static ILanguage Find(string id)
        {
            return All.FirstOrDefault(l => string.Equals(l.Id, id, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// 按打分猜测语言。要求最高分够高、且明显高于次高分，否则判为无法确定并返回 null，
        /// 由调用方提示用户手动选——猜错语言比不猜更糟。
        /// 例外是同族的语言：把整族当成一个候选跟族外比，比得过就取族内分最高的那个。
        /// </summary>
        internal static ILanguage Detect(string source)
        {
            if (string.IsNullOrWhiteSpace(source))
            {
                return null;
            }

            var ranked = Rank(source);
            var best = ranked[0];
            if (best.Score < MinimumScore)
            {
                return null;
            }

            var family = FamilyOf(best.Language);
            var runnerUp = ranked
                .Skip(1)
                .Where(x => family == null || FamilyOf(x.Language) != family)
                .Select(x => Math.Max(x.Score, 0))
                .FirstOrDefault();

            return best.Score >= runnerUp * 2 || best.Score - runnerUp >= 4 ? best.Language : null;
        }

        /// <summary>
        /// 各候选语言的得分，从高到低；同分时按 <see cref="All"/> 的顺序。
        /// 整段是一份标记文档时只在 XML/HTML 之间选，里面的脚本和样式再像别的语言也不算。
        /// </summary>
        internal static IReadOnlyList<(ILanguage Language, int Score)> Rank(string source)
        {
            var sample = new DetectionSample(source);
            var candidates = sample.IsMarkupDocument
                ? All.Where(l => FamilyOf(l) == MarkupFamily)
                : All;

            return candidates
                .Select(language => (Language: language, Score: language.ScoreLikelihood(sample)))
                .OrderByDescending(x => x.Score)
                .ToList();
        }

        private static string[] FamilyOf(ILanguage language)
        {
            return Families.FirstOrDefault(f => Array.IndexOf(f, language.Id) >= 0);
        }

        /// <summary>
        /// 解析配置/界面里的语言标识：拿到 "auto" 或空值时走自动识别。
        /// 识别不出来返回 null。
        /// </summary>
        internal static ILanguage Resolve(string id, string source)
        {
            if (string.IsNullOrEmpty(id) || string.Equals(id, AutoDetectId, StringComparison.OrdinalIgnoreCase))
            {
                return Detect(source);
            }

            return Find(id) ?? Detect(source);
        }
    }
}
