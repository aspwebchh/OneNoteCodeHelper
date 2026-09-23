using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Threading;
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

        /// <summary>
        /// 预览最多画这么多行。FlowDocument 没有虚拟化，几千行光建文档加排版就要好几秒；
        /// 插入到 OneNote 的始终是全部代码，不受这个限制。
        /// </summary>
        private const int PreviewLineLimit = 300;

        /// <summary>停止输入这么久之后才刷新预览，免得每敲一个字都重做一遍识别和排版。</summary>
        private static readonly TimeSpan PreviewDelay = TimeSpan.FromMilliseconds(250);

        private readonly PageEditor _editor;

        private readonly DispatcherTimer _previewTimer;

        /// <summary>每发起一次刷新加一。后台识别完回来时编号已经变了，说明期间又改过，结果作废。</summary>
        private int _previewVersion;

        internal InsertCodeWindow(PageEditor editor, AddInSettings settings, IntPtr ownerHandle)
        {
            _previewTimer = new DispatcherTimer { Interval = PreviewDelay };
            _previewTimer.Tick += (_, __) => UpdatePreview();
            Closed += (_, __) => _previewTimer.Stop();

            InitializeComponent();

            _editor = editor;
            Settings = settings.Clone();

            // 必须最先设：下面给各个下拉赋初值会触发 OnOptionChanged，那时复选框若还是默认的
            // 未勾选，就会把设置里的 ShowBorders 误改成 false。
            BorderBox.IsChecked = Settings.ShowBorders;

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

        /// <summary>用户是否改过语言、主题、字体或边框。</summary>
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

        private void OnCodeChanged(object sender, RoutedEventArgs e)
        {
            _previewTimer.Stop();
            _previewTimer.Start();
        }

        private void OnOptionChanged(object sender, RoutedEventArgs e)
        {
            var needsRebuild = false;
            var themeChanged = false;

            if (LanguageBox.SelectedItem is LanguageChoice choice && choice.Id != Settings.LanguageId)
            {
                Settings.LanguageId = choice.Id;
                SettingsChanged = true;
                needsRebuild = true;
            }

            if (ThemeBox.SelectedItem is CodeTheme theme && theme.Id != Settings.ThemeId)
            {
                Settings.ThemeId = theme.Id;
                SettingsChanged = true;
                themeChanged = true;
            }

            // 从下拉里选字体时，SelectionChanged 触发那一刻 FontBox.Text 还是旧值，要从事件参数里取新选的那项
            var picked = sender == FontBox && e is SelectionChangedEventArgs selection && selection.AddedItems.Count > 0
                ? selection.AddedItems[0] as string
                : null;
            var font = (picked ?? FontBox.Text)?.Trim();
            if (!string.IsNullOrEmpty(font) && font != Settings.FontFamily)
            {
                Settings.FontFamily = font;
                SettingsChanged = true;
                needsRebuild = true;
            }

            var showBorders = BorderBox.IsChecked == true;
            if (showBorders != Settings.ShowBorders)
            {
                Settings.ShowBorders = showBorders;
                SettingsChanged = true;
                needsRebuild = true;
            }

            if (needsRebuild)
            {
                UpdatePreview();
            }
            else if (themeChanged && Preview.Document != null)
            {
                // 只换了主题：给已经排好的 Run 重新上色就行，不必重新识别、分词、排版
                CodePreviewRenderer.ApplyTheme(Preview.Document, Settings.Theme);
            }
        }

        /// <summary>按当前选项解析出要用的语言；返回 null 表示自动识别没识别出来。</summary>
        private ILanguage ResolveLanguage()
        {
            var choice = LanguageBox.SelectedItem as LanguageChoice;
            return choice?.Language ?? LanguageRegistry.Detect(CodeBox.Text);
        }

        /// <summary>
        /// 刷新预览。自动识别放到后台线程做，识别完再回到界面线程画，
        /// 这样哪怕碰上识别很慢的输入，窗口也还能打字、能关。
        /// </summary>
        private void UpdatePreview()
        {
            _previewTimer.Stop();
            var version = ++_previewVersion;
            var code = CodeBox.Text;

            if (string.IsNullOrWhiteSpace(code))
            {
                ClearPreview("等待粘贴代码…");
                return;
            }

            var chosen = (LanguageBox.SelectedItem as LanguageChoice)?.Language;
            if (chosen != null)
            {
                ShowPreview(code, chosen);
                return;
            }

            var dispatcher = Dispatcher;
            Task.Run(() =>
            {
                ILanguage detected = null;
                try
                {
                    detected = LanguageRegistry.Detect(code);
                }
                catch (Exception ex)
                {
                    AddInLog.Warn("自动识别语言失败。", ex);
                }

                dispatcher.BeginInvoke(new Action(() =>
                {
                    if (version == _previewVersion)
                    {
                        ShowPreview(code, detected);
                    }
                }));
            });
        }

        private void ShowPreview(string code, ILanguage language)
        {
            // 预览失败不该弹窗打断打字，出错就把原因写到状态栏。
            try
            {
                if (language == null)
                {
                    ClearPreview("无法自动判断这段代码的语言，请在左上角手动选择语言。");
                    return;
                }

                DetectHint.Text = IsAutoSelected() ? $"已识别为 {language.DisplayName}" : string.Empty;

                // 和插入时一样去掉末尾空白，预览的行数才和插入后对得上
                code = code.TrimEnd();
                var lineCount = CountLines(code);
                var truncated = lineCount > PreviewLineLimit;

                var theme = Settings.Theme;
                var document = CodePreviewRenderer.Build(
                    truncated ? TakeLines(code, PreviewLineLimit) : code, language, theme, Settings);
                CodePreviewRenderer.ApplyTheme(document, theme);
                Preview.Document = document;

                SetStatus(truncated
                    ? $"{lineCount} 行，将按 {language.DisplayName} 高亮。预览只显示前 {PreviewLineLimit} 行，插入的是全部。"
                    : $"{lineCount} 行，将按 {language.DisplayName} 高亮。");
                InsertButton.IsEnabled = true;
            }
            catch (Exception ex)
            {
                AddInLog.Warn("生成预览失败。", ex);
                SetStatus("生成预览失败：" + ex.Message);
                InsertButton.IsEnabled = false;
            }
        }

        private void ClearPreview(string status)
        {
            Preview.Document = null;
            DetectHint.Text = string.Empty;
            SetStatus(status);
            InsertButton.IsEnabled = false;
        }

        /// <summary>行数，\r\n、\n、\r 都算一个换行，和插入时的拆行规则一致。</summary>
        private static int CountLines(string code)
        {
            var lines = 1;
            for (var i = 0; i < code.Length; i++)
            {
                var ch = code[i];
                if (ch != '\n' && ch != '\r')
                {
                    continue;
                }

                if (ch == '\r' && i + 1 < code.Length && code[i + 1] == '\n')
                {
                    i++;
                }

                lines++;
            }

            return lines;
        }

        /// <summary>取前 count 行，不含第 count 行末尾的换行。</summary>
        private static string TakeLines(string code, int count)
        {
            var lines = 0;
            for (var i = 0; i < code.Length; i++)
            {
                var ch = code[i];
                if (ch != '\n' && ch != '\r')
                {
                    continue;
                }

                if (++lines == count)
                {
                    return code.Substring(0, i);
                }

                if (ch == '\r' && i + 1 < code.Length && code[i + 1] == '\n')
                {
                    i++;
                }
            }

            return code;
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
