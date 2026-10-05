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
            new TypeScriptLanguage(),
            new PythonLanguage(),
            new GoLanguage(),
            new KotlinLanguage(),
            new RustLanguage(),
            new PhpLanguage(),
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
        /// Go、Kotlin、Rust、PHP 的关键字差别大，认错了颜色会明显不对，不进族。
        /// </summary>
        private static readonly string[][] Families =
        {
            new[] { "java", "csharp", "cpp", "javascript", "typescript" },
            new[] { "xml", "html" }
        };

        private static readonly string[] MarkupFamily = Families[1];

        internal static ILanguage Find(string id)
        {
            return All.FirstOrDefault(l => string.Equals(l.Id, id, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Markdown 围栏上写的语言名常用简写（js、sh、c#），这里按常见别名对上已有的语言。
        /// 没有对应高亮的（Markdown、Mermaid、日志）按纯文本。
        /// </summary>
        private static readonly Dictionary<string, string> Aliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["js"] = "javascript", ["jsx"] = "javascript", ["mjs"] = "javascript", ["cjs"] = "javascript", ["node"] = "javascript",
            ["ts"] = "typescript", ["tsx"] = "typescript",
            ["sh"] = "bash", ["shell"] = "bash", ["zsh"] = "bash", ["console"] = "bash", ["shellscript"] = "bash",
            ["cs"] = "csharp", ["c#"] = "csharp",
            ["c"] = "cpp", ["c++"] = "cpp", ["cc"] = "cpp", ["cxx"] = "cpp", ["h"] = "cpp", ["hpp"] = "cpp",
            ["py"] = "python", ["python3"] = "python",
            ["ps1"] = "powershell", ["pwsh"] = "powershell", ["ps"] = "powershell",
            ["cmd"] = "bat", ["batch"] = "bat",
            ["yml"] = "yaml",
            ["xaml"] = "xml", ["svg"] = "xml", ["xsd"] = "xml", ["xsl"] = "xml", ["xslt"] = "xml", ["csproj"] = "xml",
            ["xhtml"] = "html",
            ["scss"] = "css", ["less"] = "css",
            ["mysql"] = "sql", ["postgresql"] = "sql", ["postgres"] = "sql", ["psql"] = "sql", ["sqlite"] = "sql",
            ["tsql"] = "sql", ["plsql"] = "sql",
            ["golang"] = "go",
            ["kt"] = "kotlin", ["kts"] = "kotlin",
            ["rs"] = "rust",
            ["htm"] = "html", ["vue"] = "html",
            ["jsonc"] = "json", ["json5"] = "json",
            ["txt"] = "text", ["plaintext"] = "text", ["plain"] = "text", ["log"] = "text",
            ["md"] = "text", ["markdown"] = "text", ["mermaid"] = "text"
        };

        /// <summary>按 id 或常用别名找语言，找不到返回 null。</summary>
        internal static ILanguage FindByName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return null;
            name = name.Trim();
            return Find(name) ?? (Aliases.TryGetValue(name, out var id) ? Find(id) : null);
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
        /// PHP 例外：PHP 文件和模板也以 &lt;?php 或 HTML 标签开头。
        /// </summary>
        internal static IReadOnlyList<(ILanguage Language, int Score)> Rank(string source)
        {
            var sample = new DetectionSample(source);
            var candidates = sample.IsMarkupDocument
                ? All.Where(l => FamilyOf(l) == MarkupFamily || l is PhpLanguage)
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
