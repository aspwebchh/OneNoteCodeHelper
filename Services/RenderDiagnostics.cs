using System;
using System.Linq;
using System.Xml.Linq;
using OneNoteCodeHelper.Highlighting;
using OneNoteCodeHelper.Highlighting.Themes;

namespace OneNoteCodeHelper.Services
{
    /// <summary>
    /// 不接触 OneNote 就能跑通整条渲染链路的入口，供排查问题和验证新语言用。
    /// Tools/selftest.ps1 通过反射调用这里。
    /// </summary>
    internal static class RenderDiagnostics
    {
        /// <summary>渲染成逐行 HTML，行间用 \n 连接。</summary>
        internal static string RenderLines(string code, string languageId, string themeId)
        {
            var language = ResolveLanguage(code, languageId);
            var theme = CodeThemes.Find(themeId);
            var settings = new AddInSettings { ThemeId = theme.Id };

            var lines = OneNoteHtmlEncoder.BuildLines(
                code, language.Tokenize(code), theme, settings.TabWidth);

            return string.Join("\n", lines);
        }

        /// <summary>渲染成完整的 one:Table XML。</summary>
        internal static string RenderTableXml(string code, string languageId, string themeId)
        {
            var language = ResolveLanguage(code, languageId);
            var theme = CodeThemes.Find(themeId);
            var settings = new AddInSettings { ThemeId = theme.Id };

            return CodeBlockBuilder.BuildTable(code, language, theme, settings)
                .ToString(SaveOptions.None);
        }

        /// <summary>
        /// 生成 PageEditor 实际会发给 UpdatePageContent 的那份 XML，用于核对命名空间前缀、
        /// OE 数量、字体样式这些只有在完整上下文里才看得出的东西。
        /// </summary>
        internal static string RenderPageChangesXml(string code, string languageId, string themeId)
        {
            var language = ResolveLanguage(code, languageId);
            var theme = CodeThemes.Find(themeId);
            var settings = new AddInSettings { ThemeId = theme.Id };

            var table = CodeBlockBuilder.BuildTable(code, language, theme, settings);
            var outline = CodeBlockBuilder.BuildOutline(table, 36, 86, settings.CodeBlockWidth);

            return PageEditor.BuildPageChanges("{TEST-PAGE-ID}", outline);
        }

        /// <summary>把 token 序列转成 "kind:文本" 的可读形式，便于核对词法分析结果。</summary>
        internal static string DumpTokens(string code, string languageId)
        {
            var language = ResolveLanguage(code, languageId);

            return string.Join("\n", language.Tokenize(code)
                .Where(t => t.Kind != TokenKind.Plain)
                .Select(t => t.Kind + ": " + code.Substring(t.Start, t.Length).Replace("\n", "\\n")));
        }

        /// <summary>检查 token 是否完整、无重叠地覆盖了整个源码——这是 ILanguage 的核心契约。</summary>
        internal static string VerifyCoverage(string code, string languageId)
        {
            var language = ResolveLanguage(code, languageId);
            var position = 0;

            foreach (var token in language.Tokenize(code))
            {
                if (token.Start != position)
                {
                    return $"FAIL: 位置 {position} 处出现断裂或重叠，下一个 token 从 {token.Start} 开始";
                }

                if (token.Length <= 0)
                {
                    return $"FAIL: 位置 {token.Start} 处出现空 token";
                }

                position = token.End;
            }

            return position == code.Length
                ? "OK: 完整覆盖 " + position + " 个字符"
                : $"FAIL: 只覆盖到 {position}，源码共 {code.Length} 个字符";
        }

        /// <summary>返回自动识别的结果名，识别不出来返回 "(无法确定)"。</summary>
        internal static string DetectLanguageName(string code)
        {
            return LanguageRegistry.Detect(code)?.DisplayName ?? "(无法确定)";
        }

        /// <summary>自动识别时各语言的得分，形如 "java=12 csharp=5"，不列 0 分的。Tools/detect-test.ps1 调这里。</summary>
        internal static string DescribeDetectionScores(string code)
        {
            if (string.IsNullOrWhiteSpace(code))
            {
                return string.Empty;
            }

            return string.Join(" ", LanguageRegistry.Rank(code)
                .Where(x => x.Score != 0)
                .Select(x => x.Language.Id + "=" + x.Score));
        }

        private static ILanguage ResolveLanguage(string code, string languageId)
        {
            var language = LanguageRegistry.Resolve(languageId, code);
            if (language == null)
            {
                throw new InvalidOperationException("无法确定语言：" + languageId);
            }

            return language;
        }
    }
}
