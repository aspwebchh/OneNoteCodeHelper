using System;
using System.Globalization;
using System.IO;
using System.Xml.Serialization;
using OneNoteCodeHelper.Highlighting;
using OneNoteCodeHelper.Highlighting.Themes;

namespace OneNoteCodeHelper.Services
{
    /// <summary>
    /// 用户设置。必须是 public 且有无参构造，XmlSerializer 才能处理。
    /// 这里用 XmlSerializer 而不是 System.Text.Json：net48 没有内置 Json，而本项目不引 NuGet。
    /// </summary>
    public sealed class AddInSettings
    {
        private const double MinFontSize = 5;

        private const double MaxFontSize = 72;

        /// <summary>字号下拉里的备选（磅）。插入窗口和功能区共用，两边都还能手输列表外的值。</summary>
        internal static readonly double[] FontSizePresets = { 8, 9, 9.5, 10, 10.5, 11, 12, 14, 16, 18, 20 };

        /// <summary>
        /// 当前的设置版本，见 <see cref="Upgrade"/>。
        /// 2：代码框改用 GitHub 风配色，默认不显示边框。
        /// </summary>
        internal const int CurrentSettingsVersion = 2;

        /// <summary>
        /// 这份设置是哪个版本存的。不给初始值：旧版本存的文件没有这一项，读出来是 0，
        /// 由 <see cref="Upgrade"/> 按版本补上默认值的变化。
        /// </summary>
        public int SettingsVersion { get; set; }

        /// <summary>配色方案 id，见 <see cref="CodeThemes"/>。</summary>
        public string ThemeId { get; set; } = CodeThemes.Light.Id;

        /// <summary>默认语言 id，"auto" 表示自动识别。</summary>
        public string LanguageId { get; set; } = LanguageRegistry.AutoDetectId;

        public string FontFamily { get; set; } = "Consolas";

        public double FontSize { get; set; } = 9.5;

        /// <summary>一个 Tab 展开成几个空格。OneNote 里制表符不可靠，必须先展开。</summary>
        public int TabWidth { get; set; } = 4;

        /// <summary>插入新代码框时的宽度（磅）。</summary>
        public double CodeBlockWidth { get; set; } = 520;

        /// <summary>代码框是否显示边框。默认只靠底色区分：OneNote 的表格边框颜色改不了，深灰细线显得生硬。</summary>
        public bool ShowBorders { get; set; }

        /// <summary>
        /// Agent 窗口「功能」下拉选的是哪一项：<see cref="AiConfigStore.AgentFunctionName"/>，
        /// 或 ai-settings.xml 里 Function 的 name。旧版本存的 AiFunction 元素不再读取，升级后默认回到 Agent。
        /// </summary>
        public string AgentFunction { get; set; } = AiConfigStore.AgentFunctionName;

        /// <summary>Agent 窗口用哪个模型，存模型 id（ai-settings.xml 里 Model 的 id）。</summary>
        public string AiModel { get; set; } = AiConfigStore.DefaultModelId;

        /// <summary>思考强度（none/low/medium/high/max），见 <see cref="AiEfforts"/>。</summary>
        public string AiEffort { get; set; } = AiEfforts.DefaultId;

        internal CodeTheme Theme => CodeThemes.Find(ThemeId);

        internal AddInSettings Clone()
        {
            return new AddInSettings
            {
                SettingsVersion = SettingsVersion,
                ThemeId = ThemeId,
                LanguageId = LanguageId,
                FontFamily = FontFamily,
                FontSize = FontSize,
                TabWidth = TabWidth,
                CodeBlockWidth = CodeBlockWidth,
                ShowBorders = ShowBorders,
                AgentFunction = AgentFunction,
                AiModel = AiModel,
                AiEffort = AiEffort
            };
        }

        /// <summary>把明显不合理的值拉回可用范围，避免手改配置文件后把渲染搞崩。</summary>
        internal void Normalize()
        {
            if (string.IsNullOrWhiteSpace(FontFamily))
            {
                FontFamily = "Consolas";
            }

            FontSize = NormalizeFontSize(FontSize);
            TabWidth = (int)Clamp(TabWidth, 1, 16);
            CodeBlockWidth = Clamp(CodeBlockWidth, 100, 2000);

            if (CodeThemes.Find(ThemeId) is CodeTheme theme)
            {
                ThemeId = theme.Id;
            }

            if (!string.Equals(LanguageId, LanguageRegistry.AutoDetectId, StringComparison.OrdinalIgnoreCase)
                && LanguageRegistry.Find(LanguageId) == null)
            {
                LanguageId = LanguageRegistry.AutoDetectId;
            }

            // 功能和模型的名字要对照 ai-settings.xml，那边随时会改，用的时候再兜底（找不到就取第一项）。
            AiEffort = AiEfforts.Normalize(AiEffort);
        }

        /// <summary>
        /// 把旧版本存的设置升到当前版本，返回是否有改动。默认值变了的项在这里改一次，
        /// 之后用户再改回去就一直保留。
        /// </summary>
        internal bool Upgrade()
        {
            if (SettingsVersion >= CurrentSettingsVersion)
            {
                return false;
            }

            // 版本 2：旧版本默认带边框，保存的设置里几乎都是 true，这里统一关掉一次。
            ShowBorders = false;
            SettingsVersion = CurrentSettingsVersion;
            return true;
        }

        internal static string FormatFontSize(double size)
        {
            return size.ToString("0.#", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// 解析用户输入的字号，允许带「pt」或「磅」后缀。超范围的夹回去，
        /// 并取整到半磅——OneNote 自己的字号就是这个粒度，也让结果能和下拉里的项对上。
        /// </summary>
        internal static bool TryParseFontSize(string text, out double size)
        {
            size = 0;
            var trimmed = text?.Trim() ?? string.Empty;

            if (trimmed.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
            {
                trimmed = trimmed.Substring(0, trimmed.Length - 2).TrimEnd();
            }
            else if (trimmed.EndsWith("磅", StringComparison.Ordinal))
            {
                trimmed = trimmed.Substring(0, trimmed.Length - 1).TrimEnd();
            }

            if (!double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
                || double.IsNaN(value) || double.IsInfinity(value))
            {
                return false;
            }

            size = NormalizeFontSize(value);
            return true;
        }

        private static double NormalizeFontSize(double size)
        {
            return Math.Round(Clamp(size, MinFontSize, MaxFontSize) * 2, MidpointRounding.AwayFromZero) / 2;
        }

        private static double Clamp(double value, double min, double max)
        {
            return value < min ? min : value > max ? max : value;
        }
    }

    /// <summary>设置的读写。任何异常都降级为「用默认值」，不能让配置问题拖垮外接程序。</summary>
    internal static class SettingsStore
    {
        private static readonly XmlSerializer Serializer = new XmlSerializer(typeof(AddInSettings));

        internal static string SettingsPath { get; } = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "OneNoteCodeHelper",
            "settings.xml");

        internal static AddInSettings Load()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    AddInSettings settings;
                    using (var stream = File.OpenRead(SettingsPath))
                    {
                        settings = Serializer.Deserialize(stream) as AddInSettings;
                    }

                    if (settings != null)
                    {
                        settings.Normalize();

                        // 升级后马上存回去，免得每次启动都当成旧文件再升一遍
                        if (settings.Upgrade())
                        {
                            Save(settings);
                            AddInLog.Info($"设置升级到版本 {AddInSettings.CurrentSettingsVersion}：代码框默认不显示边框。");
                        }

                        return settings;
                    }
                }
            }
            catch (Exception ex)
            {
                AddInLog.Warn("读取设置失败，改用默认值。", ex);
            }

            // 默认值本身就是当前版本，不能留 0，否则用户打开边框后下次读取又会被 Upgrade 关掉。
            return new AddInSettings { SettingsVersion = AddInSettings.CurrentSettingsVersion };
        }

        internal static void Save(AddInSettings settings)
        {
            if (settings == null)
            {
                return;
            }

            try
            {
                settings.Normalize();

                var directory = Path.GetDirectoryName(SettingsPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                using (var stream = File.Create(SettingsPath))
                {
                    Serializer.Serialize(stream, settings);
                }
            }
            catch (Exception ex)
            {
                AddInLog.Warn("保存设置失败。", ex);
            }
        }
    }
}
