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

        /// <summary>自动识别最多看这么多字符，约四五百行代码。</summary>
        private const int DetectionSampleLength = 16 * 1024;

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
            new TextLanguage()
        };

        internal static ILanguage Find(string id)
        {
            return All.FirstOrDefault(l => string.Equals(l.Id, id, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// 按打分猜测语言。要求最高分必须为正、且明显高于次高分，否则判为无法确定并返回 null，
        /// 由调用方提示用户手动选——猜错语言比不猜更糟。
        /// </summary>
        internal static ILanguage Detect(string source)
        {
            if (string.IsNullOrWhiteSpace(source))
            {
                return null;
            }

            source = TakeDetectionSample(source);
            var ranked = All
                .Select(language => new { Language = language, Score = language.ScoreLikelihood(source) })
                .OrderByDescending(x => x.Score)
                .ToList();

            var best = ranked[0];
            if (best.Score <= 0)
            {
                return null;
            }

            var runnerUp = ranked.Count > 1 ? ranked[1].Score : 0;
            return best.Score >= runnerUp * 2 || best.Score - runnerUp >= 4 ? best.Language : null;
        }

        /// <summary>
        /// 只拿开头一段来打分，在行尾截断。几百行足够看出是什么语言；整份几千行的文件全量跑上百条正则
        /// 要大半秒，而「插入代码」窗口每改一次都要重新识别。
        /// </summary>
        private static string TakeDetectionSample(string source)
        {
            if (source.Length <= DetectionSampleLength)
            {
                return source;
            }

            var cut = source.LastIndexOf('\n', DetectionSampleLength - 1);
            return source.Substring(0, cut > 0 ? cut + 1 : DetectionSampleLength);
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
