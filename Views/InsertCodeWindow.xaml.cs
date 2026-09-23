using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Interop;
using OneNoteCodeHelper.Highlighting;
using OneNoteCodeHelper.Highlighting.Themes;
using OneNoteCodeHelper.Services;

namespace OneNoteCodeHelper.Views
{
    /// <summary>
    /// 「插入代码」窗口：粘贴代码 → 选语言与主题 → 看预览 → 插入到当前页。
    ///
    /// 这个窗口在 COM 代理进程的独立 STA 线程运行，用 ShowDialog 提供消息循环。
    /// </summary>
    public partial class InsertCodeWindow : Window
    {
        /// <summary>下拉里代表「自动识别」的那一项。</summary>
        private sealed class LanguageChoice
        {
            internal LanguageChoice(string id, string label, ILanguage language)
            {
                Id = id;
                Label = label;
                Language = language;
            }

            internal string Id { get; }

            public string Label { get; }

            internal ILanguage Language { get; }

            public override string ToString() => Label;
        }

        /// <summary>
        /// 字体下拉的备选。OneNote 不认 CSS 字体栈（实测只取第一个名字），
        /// 所以代码里有中文时只能整体换成中英文都等宽的字体，新宋体是 Windows 自带里最合适的。
        /// 下拉可编辑，想用别的字体直接敲就行。
        /// </summary>
        private static readonly string[] MonospaceFonts =
        {
            "Consolas",
            "NSimSun",
            "Cascadia Mono",
            "Courier New"
        };

        private readonly PageEditor _editor;

        internal InsertCodeWindow(PageEditor editor, AddInSettings settings, IntPtr ownerHandle)
        {
            InitializeComponent();

            _editor = editor;
            Settings = settings.Clone();

            // 认 OneNote 主窗口做属主，免得窗口跑到 OneNote 后面去。
            if (ownerHandle != IntPtr.Zero)
            {
                new WindowInteropHelper(this).Owner = ownerHandle;
            }
            else
            {
                WindowStartupLocation = WindowStartupLocation.CenterScreen;
            }

            var languageChoices = BuildLanguageChoices();
            LanguageBox.ItemsSource = languageChoices;
            LanguageBox.DisplayMemberPath = nameof(LanguageChoice.Label);
            LanguageBox.SelectedItem = languageChoices
                .FirstOrDefault(c => string.Equals(c.Id, Settings.LanguageId, StringComparison.OrdinalIgnoreCase));

            ThemeBox.ItemsSource = CodeThemes.All.ToList();
            ThemeBox.DisplayMemberPath = nameof(CodeTheme.DisplayName);
            ThemeBox.SelectedItem = Settings.Theme;

            FontBox.ItemsSource = MonospaceFonts;
            FontBox.Text = Settings.FontFamily;

            Loaded += (_, __) => CodeBox.Focus();
            UpdatePreview();
        }

        /// <summary>窗口里改过的设置。关闭后由调用方决定要不要持久化。</summary>
        internal AddInSettings Settings { get; }

        /// <summary>用户是否改过语言或主题。</summary>
        internal bool SettingsChanged { get; private set; }

        private static List<LanguageChoice> BuildLanguageChoices()
        {
            var choices = new List<LanguageChoice>
            {
                new LanguageChoice(LanguageRegistry.AutoDetectId, "自动识别", null)
            };

            choices.AddRange(LanguageRegistry.All.Select(l => new LanguageChoice(l.Id, l.DisplayName, l)));
            return choices;
        }

        private void OnCodeChanged(object sender, RoutedEventArgs e) => UpdatePreview();

        private void OnOptionChanged(object sender, RoutedEventArgs e)
        {
            if (LanguageBox.SelectedItem is LanguageChoice choice && choice.Id != Settings.LanguageId)
            {
                Settings.LanguageId = choice.Id;
                SettingsChanged = true;
            }

            if (ThemeBox.SelectedItem is CodeTheme theme && theme.Id != Settings.ThemeId)
            {
                Settings.ThemeId = theme.Id;
                SettingsChanged = true;
            }

            var font = FontBox.Text?.Trim();
            if (!string.IsNullOrEmpty(font) && font != Settings.FontFamily)
            {
                Settings.FontFamily = font;
                SettingsChanged = true;
            }

            UpdatePreview();
        }

        /// <summary>按当前选项解析出要用的语言；返回 null 表示自动识别没识别出来。</summary>
        private ILanguage ResolveLanguage()
        {
            var choice = LanguageBox.SelectedItem as LanguageChoice;
            return choice?.Language ?? LanguageRegistry.Detect(CodeBox.Text);
        }

        private void UpdatePreview()
        {
            // 预览失败不该弹窗打断打字，出错就把原因写到状态栏。
            try
            {
                var code = CodeBox.Text;
                var theme = Settings.Theme;

                if (string.IsNullOrWhiteSpace(code))
                {
                    Preview.Document = null;
                    DetectHint.Text = string.Empty;
                    SetStatus("等待粘贴代码…");
                    InsertButton.IsEnabled = false;
                    return;
                }

                var language = ResolveLanguage();
                if (language == null)
                {
                    Preview.Document = null;
                    DetectHint.Text = string.Empty;
                    SetStatus("无法自动判断这段代码的语言，请在左上角手动选择 Java 或 Lua。");
                    InsertButton.IsEnabled = false;
                    return;
                }

                DetectHint.Text = IsAutoSelected() ? $"已识别为 {language.DisplayName}" : string.Empty;

                var document = CodePreviewRenderer.Build(code, language, theme, Settings);
                CodePreviewRenderer.ApplyTheme(document, theme);
                Preview.Document = document;

                SetStatus($"{document.Blocks.Count} 行，将按 {language.DisplayName} 高亮。");
                InsertButton.IsEnabled = true;
            }
            catch (Exception ex)
            {
                AddInLog.Warn("生成预览失败。", ex);
                SetStatus("生成预览失败：" + ex.Message);
                InsertButton.IsEnabled = false;
            }
        }

        private bool IsAutoSelected()
        {
            return LanguageBox.SelectedItem is LanguageChoice choice && choice.Language == null;
        }

        private void OnInsert(object sender, RoutedEventArgs e)
        {
            var language = ResolveLanguage();
            if (language == null)
            {
                SetStatus("请先选择语言。");
                return;
            }

            try
            {
                InsertButton.IsEnabled = false;
                var result = _editor.InsertCode(CodeBox.Text, language, Settings);

                if (result.Success)
                {
                    AddInLog.Info(result.Message);
                    DialogResult = true;
                    Close();
                    return;
                }

                SetStatus(result.Message);
                InsertButton.IsEnabled = true;
            }
            catch (Exception ex)
            {
                AddInLog.Error("插入代码失败。", ex);
                MessageBox.Show(this, "插入失败：" + ex.Message, "OneNote 代码高亮",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                InsertButton.IsEnabled = true;
            }
        }

        private void OnCancel(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void SetStatus(string text) => StatusText.Text = text;
    }
}
