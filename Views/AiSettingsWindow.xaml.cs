using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using OneNoteCodeHelper.Interop;
using OneNoteCodeHelper.Services;
using OneNoteCodeHelper.Services.Agent;

namespace OneNoteCodeHelper.Views
{
    /// <summary>
    /// ai-settings.xml 的图形界面：接口、模型、文字功能、Agent 四页，保存时写回同一个文件（见 <see cref="AiConfigStore.Save"/>）。
    /// 界面上的值不直接改配置对象，点「保存」时先校验，再整份交给 AiConfigStore。
    ///
    /// 从 Agent 窗口右上角的「AI 配置」打开，在 Agent 窗口的线程上模态显示。关窗后 Agent 窗口按文件修改时间重新读配置，这边保存后不用通知它。
    /// </summary>
    public partial class AiSettingsWindow : Window
    {
        private const string Caption = "AI 配置";

        private const int ApiPageIndex = 0, ModelsPageIndex = 1, FunctionsPageIndex = 2, AgentPageIndex = 3;

        /// <summary>Agent 数值项的显示名，键是 <see cref="AgentOptions.Numbers"/> 里的节点名。</summary>
        private static readonly Dictionary<string, string> NumberLabels = new Dictionary<string, string>
        {
            ["MaxTurns"] = "最多轮数",
            ["MaxToolCalls"] = "最多工具调用次数",
            ["TimeoutSeconds"] = "任务总时限（秒）",
            ["MaxPageChars"] = "页面字数上限",
            ["MaxRequestChars"] = "请求字符上限"
        };

        /// <summary>Agent 开关的显示名和所在分组（0 格式工具、1 结构工具、2 接口兼容），键是 <see cref="AgentOptions.Switches"/> 里的节点名。</summary>
        private static readonly Dictionary<string, (int Group, string Label)> SwitchLabels = new Dictionary<string, (int, string)>
        {
            ["EnableNativeHeadings"] = (0, "原生标题样式"),
            ["EnableParagraphSpacing"] = (0, "段前段后间距"),
            ["EnableMixedOutlines"] = (0, "整理图文混排的文本框"),
            ["EnableCodeHighlight"] = (0, "把代码转成代码框"),
            ["EnableLists"] = (0, "项目符号和编号列表"),
            ["EnableTags"] = (0, "待办、重要、问题标记"),
            ["EnableTableStyles"] = (0, "表格边框和底色"),
            ["EnableMarkdownCleanup"] = (0, "去除 Markdown 符号"),
            ["EnableBlankLineRemoval"] = (1, "删除多余的空行"),
            ["EnableIndent"] = (1, "调整缩进层级"),
            ["EnableMoves"] = (1, "移动段落、合并文本框"),
            ["EnableInsert"] = (1, "插入摘要、目录、小标题和空行"),
            ["EnableTextTables"] = (1, "把文字转成表格"),
            ["SendThinking"] = (2, "发送思考参数"),
            ["ReplayReasoning"] = (2, "回传思考内容"),
            ["StreamUsage"] = (2, "请求用量统计")
        };

        private readonly ObservableCollection<string> _models = new ObservableCollection<string>();
        private readonly ObservableCollection<FunctionItem> _functions = new ObservableCollection<FunctionItem>();
        private readonly Dictionary<string, TextBox> _numberBoxes = new Dictionary<string, TextBox>();
        private readonly Dictionary<string, CheckBox> _switchBoxes = new Dictionary<string, CheckBox>();
        private readonly FrameworkElement[] _pages;
        /// <summary>界面上一次和文件一致时的内容，和当前的比较判断有没有没保存的修改。</summary>
        private string _snapshot;
        /// <summary>界面上一次和文件一致时文件的修改时间，保存前据此判断文件是不是在别处被改过。</summary>
        private DateTime _stamp;
        /// <summary>点了「取消」或已经用文本编辑器打开：关窗时不再问要不要保存。</summary>
        private bool _discard;
        /// <summary>正在把选中的功能填进编辑区，这时的 TextChanged 不算用户改动。</summary>
        private bool _loadingFunction;

        internal AiSettingsWindow(AiConfig config, Exception loadError, IntPtr owner)
        {
            InitializeComponent();
            _pages = new FrameworkElement[] { ApiPage, ModelsPage, FunctionsPage, AgentPage };
            PathText.Text = AiConfigStore.ConfigPath;
            TimeoutHint.Text = $"单次请求从发出到收完的总时限，{AiConfigStore.MinTimeoutSeconds}–{AiConfigStore.MaxTimeoutSeconds}。另有固定的 60 秒无数据判定，不在这里配。";
            MaxTokensHint.Text = $"单次请求最多输出多少 token（含思考过程），0 表示用接口的默认值，最大 {AiConfigStore.MaxMaxTokens}。";
            ModelList.ItemsSource = _models;
            FunctionList.ItemsSource = _functions;
            FontPicker.ItemsSource = ParagraphStyles.Fonts;
            AgentDefaultRequestBox.MaxLength = AgentOptions.MaxRequestLength;
            _models.CollectionChanged += (_, __) => UpdateModelButtons();
            _functions.CollectionChanged += (_, __) => UpdateFunctionButtons();
            BuildAgentOptions();
            Fill(config);
            ShowLoadError(loadError);
            _stamp = AiConfigStore.Stamp();
            PageList.SelectedIndex = ApiPageIndex;

            AppIcon.Source = WindowIcons.Load("AiConfig");
            if (AppIcon.Source != null) Icon = AppIcon.Source;
            var interop = new WindowInteropHelper(this);
            if (owner != IntPtr.Zero) interop.Owner = owner;
            SourceInitialized += (_, __) =>
            {
                NativeMethods.TrySetCaptionColor(interop.Handle, ((SolidColorBrush)Background).Color);
                if (owner != IntPtr.Zero) NativeMethods.CenterOver(interop.Handle, owner);
            };
            // 开着窗口时可能在文本编辑器里改过文件；界面没改动时切回来就重新读。
            Activated += (_, __) => RefreshFromFile();
            Closing += (_, e) =>
            {
                if (!_discard && IsDirty && !ConfirmLeave("AI 配置有修改还没保存，要保存吗？")) e.Cancel = true;
                if (!e.Cancel) NativeMethods.ReturnForeground(interop.Handle, owner);
            };
            PreviewKeyDown += (_, e) =>
            {
                if (e.Key == Key.Escape && !FontPicker.IsDropDownOpen) Close();
            };
        }

        private bool IsDirty => Snapshot() != _snapshot;

        private string ApiKeyValue => ShowKeyBox.IsChecked == true ? ApiKeyText.Text : ApiKeyBox.Password;

        /// <summary>切到第 index 页（接口、模型、文字功能、Agent）。</summary>
        internal void ShowPage(int index) => PageList.SelectedIndex = index;

        /// <summary>按 AgentOptions 的两张表生成 Agent 页的数值框和开关，表里加了新项界面自动跟上。</summary>
        private void BuildAgentOptions()
        {
            foreach (var option in AgentOptions.Numbers)
            {
                var row = new Grid { Margin = new Thickness(0, 0, 0, 8) };
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
                row.ColumnDefinitions.Add(new ColumnDefinition());
                var box = new TextBox { Style = (Style)FindResource("InputBox"), ToolTip = option.Key };
                var range = new TextBlock
                {
                    Margin = new Thickness(12, 0, 0, 0),
                    VerticalAlignment = VerticalAlignment.Center,
                    Style = (Style)FindResource("Hint"),
                    Text = $"{option.Min}–{option.Max}，默认 {option.Default}"
                };
                Grid.SetColumn(box, 1);
                Grid.SetColumn(range, 2);
                row.Children.Add(new TextBlock { Text = Label(NumberLabels, option.Key), Style = (Style)FindResource("FieldLabel") });
                row.Children.Add(box);
                row.Children.Add(range);
                AgentNumbersPanel.Children.Add(row);
                _numberBoxes[option.Key] = box;
            }

            var groups = new[] { FormatSwitchesPanel, StructureSwitchesPanel, CompatSwitchesPanel };
            foreach (var option in AgentOptions.Switches)
            {
                // 表里新加、这里还没起名的开关放进格式工具，直接显示节点名，不会漏掉。
                var info = SwitchLabels.TryGetValue(option.Key, out var known) ? known : (0, option.Key);
                var box = new CheckBox { Content = info.Item2, ToolTip = option.Key, Margin = new Thickness(0, 0, 12, 8) };
                groups[info.Item1].Children.Add(box);
                _switchBoxes[option.Key] = box;
            }
        }

        private static string Label(Dictionary<string, string> labels, string key) => labels.TryGetValue(key, out var label) ? label : key;

        /// <summary>把配置填进界面，并以此为「没有改动」的基准。</summary>
        private void Fill(AiConfig config)
        {
            ApiUrlBox.Text = config.ApiUrl;
            ApiKeyBox.Password = config.ApiKey;
            ApiKeyText.Text = config.ApiKey;
            TimeoutBox.Text = config.TimeoutSeconds.ToString(CultureInfo.InvariantCulture);
            MaxTokensBox.Text = config.MaxTokens.ToString(CultureInfo.InvariantCulture);

            _models.Clear();
            foreach (var model in config.Models) _models.Add(model.Id);
            ModelList.SelectedIndex = _models.Count > 0 ? 0 : -1;

            var selected = FunctionList.SelectedIndex;
            _functions.Clear();
            foreach (var function in config.Functions)
                _functions.Add(new FunctionItem { Name = function.Name, Prompt = function.Prompt, RemoveExtraBlankLines = function.RemoveExtraBlankLines });
            FunctionList.SelectedIndex = _functions.Count == 0 ? -1 : Math.Min(Math.Max(0, selected), _functions.Count - 1);

            foreach (var option in AgentOptions.Numbers)
                _numberBoxes[option.Key].Text = option.Get(config.Agent).ToString(CultureInfo.InvariantCulture);
            foreach (var option in AgentOptions.Switches)
                _switchBoxes[option.Key].IsChecked = option.Get(config.Agent);
            FontPicker.SelectedItem = ParagraphStyles.Fonts.FirstOrDefault(f => f == config.Agent.FontFamily) ?? AgentOptions.DefaultFontFamily;
            AgentDefaultRequestBox.Text = config.Agent.DefaultRequest;

            UpdateModelButtons();
            UpdateFunctionButtons();
            _snapshot = Snapshot();
        }

        /// <summary>界面上全部可改的内容拼成的一个字符串，只用来比较有没有改动。</summary>
        private string Snapshot()
        {
            var parts = new List<string>
            {
                ApiUrlBox.Text.Trim(), ApiKeyValue.Trim(), TimeoutBox.Text.Trim(), MaxTokensBox.Text.Trim(),
                FontPicker.SelectedItem as string, AgentOptions.NormalizeRequest(AgentDefaultRequestBox.Text), "M" + _models.Count
            };
            parts.AddRange(_models);
            parts.Add("F" + _functions.Count);
            parts.AddRange(_functions.Select(f => f.Name.Trim() + "\u0002" + NormalizePrompt(f.Prompt) + "\u0002" + f.RemoveExtraBlankLines));
            parts.AddRange(_numberBoxes.Values.Select(b => b.Text.Trim()));
            parts.AddRange(_switchBoxes.Values.Select(b => (b.IsChecked == true).ToString()));
            return string.Join("\u0001", parts);
        }

        /// <summary>文本框里按回车换出来的是 \r\n，存进文件、和文件比较都按 \n。</summary>
        private static string NormalizePrompt(string prompt) => (prompt ?? string.Empty).Replace("\r\n", "\n").Trim();

        private void OnRestoreAgentDefaultRequest(object sender, RoutedEventArgs e)
        {
            AgentDefaultRequestBox.Text = AgentOptions.DefaultRequestText;
            ClearError();
        }

        private void ShowLoadError(Exception error)
        {
            LoadErrorBox.Visibility = error == null ? Visibility.Collapsed : Visibility.Visible;
            LoadErrorText.Text = error == null
                ? string.Empty
                : "配置文件读不了，下面显示的是默认值：" + error.Message + "\n保存时会先把原文件备份成 ai-settings.xml.bak，再重新写一份。";
        }

        /// <summary>文件在别处被改过、界面又没有改动时，重新读文件。读不了（编辑器还在写、XML 写坏了）就显示提示，界面不动。</summary>
        private void RefreshFromFile()
        {
            var stamp = AiConfigStore.Stamp();
            if (stamp == _stamp || IsDirty) return;
            if (!AiConfigStore.TryLoad(out var config, out var error))
            {
                ShowLoadError(error);
                return;
            }

            Fill(config);
            ShowLoadError(null);
            _stamp = stamp;
            AddInLog.Info("AI 配置窗口已重新读取配置文件。");
        }

        private void OnPageChanged(object sender, SelectionChangedEventArgs e)
        {
            if (PageList.SelectedIndex < 0) return;
            for (var i = 0; i < _pages.Length; i++)
                _pages[i].Visibility = i == PageList.SelectedIndex ? Visibility.Visible : Visibility.Collapsed;
        }

        private void OnShowKeyChanged(object sender, RoutedEventArgs e)
        {
            var show = ShowKeyBox.IsChecked == true;
            if (show) ApiKeyText.Text = ApiKeyBox.Password;
            else ApiKeyBox.Password = ApiKeyText.Text;
            ApiKeyText.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
            ApiKeyBox.Visibility = show ? Visibility.Collapsed : Visibility.Visible;
        }

        #region 模型

        private void OnAddModel(object sender, RoutedEventArgs e)
        {
            ClearError();
            var id = NewModelBox.Text.Trim();
            if (id.Length == 0)
            {
                NewModelBox.Focus();
                return;
            }

            var existing = _models.FirstOrDefault(m => string.Equals(m, id, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                ModelList.SelectedItem = existing;
                ShowError(ModelsPageIndex, NewModelBox, "已经有模型「" + existing + "」了。");
                return;
            }

            _models.Add(id);
            ModelList.SelectedIndex = _models.Count - 1;
            ModelList.ScrollIntoView(id);
            NewModelBox.Clear();
            NewModelBox.Focus();
        }

        private void OnNewModelKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter) return;
            e.Handled = true;
            OnAddModel(sender, e);
        }

        private void OnRemoveModel(object sender, RoutedEventArgs e)
        {
            var index = ModelList.SelectedIndex;
            if (index < 0) return;
            _models.RemoveAt(index);
            ModelList.SelectedIndex = Math.Min(index, _models.Count - 1);
        }

        private void OnMoveModelUp(object sender, RoutedEventArgs e) => MoveSelected(ModelList, _models, -1);

        private void OnMoveModelDown(object sender, RoutedEventArgs e) => MoveSelected(ModelList, _models, 1);

        private void OnModelSelectionChanged(object sender, SelectionChangedEventArgs e) => UpdateModelButtons();

        private void UpdateModelButtons()
        {
            var index = ModelList.SelectedIndex;
            RemoveModelButton.IsEnabled = index >= 0;
            ModelUpButton.IsEnabled = index > 0;
            ModelDownButton.IsEnabled = index >= 0 && index < _models.Count - 1;
        }

        #endregion

        #region 文字功能

        private void OnAddFunction(object sender, RoutedEventArgs e)
        {
            ClearError();
            var name = "新功能";
            for (var i = 2; _functions.Any(f => string.Equals(f.Name.Trim(), name, StringComparison.OrdinalIgnoreCase)); i++)
                name = "新功能 " + i;
            var item = new FunctionItem { Name = name, Prompt = string.Empty };
            _functions.Add(item);
            SelectFunction(item);
            FunctionNameBox.Focus();
            FunctionNameBox.SelectAll();
        }

        private void OnRemoveFunction(object sender, RoutedEventArgs e)
        {
            var index = FunctionList.SelectedIndex;
            if (index < 0) return;
            _functions.RemoveAt(index);
            FunctionList.SelectedIndex = Math.Min(index, _functions.Count - 1);
        }

        private void OnMoveFunctionUp(object sender, RoutedEventArgs e) => MoveSelected(FunctionList, _functions, -1);

        private void OnMoveFunctionDown(object sender, RoutedEventArgs e) => MoveSelected(FunctionList, _functions, 1);

        /// <summary>
        /// 把配置里没有的内置功能按默认提示词加到最后。插件只在配置文件不存在时写默认值，
        /// 所以后来新增的内置功能（比如「智能校正」）不会自动进旧配置，这里补。
        /// </summary>
        private void OnAddBuiltInFunctions(object sender, RoutedEventArgs e)
        {
            ClearError();
            FunctionItem first = null;
            foreach (var function in MissingBuiltInFunctions().ToList())
            {
                var item = new FunctionItem { Name = function.Name, Prompt = function.Prompt, RemoveExtraBlankLines = function.RemoveExtraBlankLines };
                _functions.Add(item);
                first = first ?? item;
            }

            if (first != null) SelectFunction(first);
        }

        private IEnumerable<AiFunction> MissingBuiltInFunctions() =>
            AiConfigStore.Default.Functions.Where(d => !_functions.Any(f => string.Equals(f.Name.Trim(), d.Name, StringComparison.OrdinalIgnoreCase)));

        /// <summary>换了选中的功能：把它的名称、提示词和删空行填进右边的编辑区。</summary>
        private void OnFunctionSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var item = FunctionList.SelectedItem as FunctionItem;
            _loadingFunction = true;
            try
            {
                FunctionNameBox.Text = item?.Name ?? string.Empty;
                FunctionPromptBox.Text = item?.Prompt ?? string.Empty;
                FunctionBlankLinesBox.IsChecked = item?.RemoveExtraBlankLines == true;
            }
            finally
            {
                _loadingFunction = false;
            }

            UpdateFunctionButtons();
        }

        /// <summary>编辑区改了什么就写回选中的那一项，列表里的名字跟着变。</summary>
        private void OnFunctionEdited(object sender, RoutedEventArgs e)
        {
            if (_loadingFunction || !(FunctionList.SelectedItem is FunctionItem item)) return;
            item.Name = FunctionNameBox.Text;
            item.Prompt = FunctionPromptBox.Text;
            item.RemoveExtraBlankLines = FunctionBlankLinesBox.IsChecked == true;
            AddBuiltInButton.IsEnabled = MissingBuiltInFunctions().Any();
        }

        private void UpdateFunctionButtons()
        {
            var index = FunctionList.SelectedIndex;
            RemoveFunctionButton.IsEnabled = index >= 0;
            FunctionUpButton.IsEnabled = index > 0;
            FunctionDownButton.IsEnabled = index >= 0 && index < _functions.Count - 1;
            AddBuiltInButton.IsEnabled = MissingBuiltInFunctions().Any();
            FunctionEditor.IsEnabled = FunctionList.SelectedItem != null;
        }

        private void SelectFunction(FunctionItem item)
        {
            FunctionList.SelectedItem = item;
            FunctionList.ScrollIntoView(item);
        }

        #endregion

        private static void MoveSelected<T>(ListBox list, ObservableCollection<T> items, int offset)
        {
            var index = list.SelectedIndex;
            var target = index + offset;
            if (index < 0 || target < 0 || target >= items.Count) return;
            items.Move(index, target);
            list.SelectedIndex = target;
            list.ScrollIntoView(items[target]);
        }

        #region 保存

        private void OnSave(object sender, RoutedEventArgs e)
        {
            if (Save())
            {
                Close();
            }
        }

        private void OnCancel(object sender, RoutedEventArgs e) => CloseWithoutSaving();

        /// <summary>放弃没保存的修改直接关窗，不问。「取消」和 OneNote 关闭时用。</summary>
        internal void CloseWithoutSaving()
        {
            _discard = true;
            Close();
        }

        /// <summary>
        /// 校验界面内容并写回配置文件。文件在窗口打开后被别处改过时先确认。
        /// 返回 false 表示没有保存：校验不过（已切到出错的地方并提示）、用户取消，或写文件失败。
        /// </summary>
        private bool Save()
        {
            ClearError();
            if (!TryBuildConfig(out var config)) return false;

            if (AiConfigStore.Stamp() != _stamp &&
                MessageBox.Show(this,
                    "配置文件在这个窗口打开之后被改过（比如在文本编辑器里）。\n\n" +
                    "继续保存会用窗口里的内容覆盖接口、模型、文字功能和 Agent 设置；文件里窗口没有的内容仍然保留。",
                    Caption, MessageBoxButton.OKCancel, MessageBoxImage.Warning) != MessageBoxResult.OK)
            {
                return false;
            }

            try
            {
                AiConfigStore.Save(config);
            }
            catch (Exception ex)
            {
                AddInLog.Error("保存 AI 配置失败。", ex);
                ShowError(null, null, "保存失败：" + ex.Message);
                return false;
            }

            _stamp = AiConfigStore.Stamp();
            _snapshot = Snapshot();
            ShowLoadError(null);
            return true;
        }

        /// <summary>
        /// 按界面内容拼出一份配置。校验规则和读配置时的兜底一致：数值要在读取时的范围内（超出的话读的时候会被悄悄截断），
        /// 名称不能重复（下拉按名字不区分大小写找），也不能和 Agent 那一项同名（窗口里会被跳过）。
        /// </summary>
        internal bool TryBuildConfig(out AiConfig config)
        {
            config = null;

            var url = ApiUrlBox.Text.Trim();
            if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
                return ShowError(ApiPageIndex, ApiUrlBox, "接口地址要填以 http:// 或 https:// 开头的完整地址。");
            if (!TryReadNumber(TimeoutBox, AiConfigStore.MinTimeoutSeconds, AiConfigStore.MaxTimeoutSeconds, out var timeout))
                return ShowError(ApiPageIndex, TimeoutBox, $"请求超时要填 {AiConfigStore.MinTimeoutSeconds}–{AiConfigStore.MaxTimeoutSeconds} 之间的整数。");
            if (!TryReadNumber(MaxTokensBox, AiConfigStore.MinMaxTokens, AiConfigStore.MaxMaxTokens, out var maxTokens))
                return ShowError(ApiPageIndex, MaxTokensBox, $"最大输出 token 要填 {AiConfigStore.MinMaxTokens}–{AiConfigStore.MaxMaxTokens} 之间的整数。");

            if (_models.Count == 0)
                return ShowError(ModelsPageIndex, NewModelBox, "至少要有一个模型。");

            if (_functions.Count == 0)
                return ShowError(FunctionsPageIndex, null, "至少要有一个文字功能，可以点「补回内置功能」。");
            var functions = new List<AiFunction>();
            foreach (var item in _functions)
            {
                var name = item.Name.Trim();
                var prompt = NormalizePrompt(item.Prompt);
                var number = functions.Count + 1;
                if (name.Length == 0)
                    return ShowFunctionError(item, FunctionNameBox, $"第 {number} 个文字功能还没有名称。");
                if (string.Equals(name, AiConfigStore.AgentFunctionName, StringComparison.OrdinalIgnoreCase))
                    return ShowFunctionError(item, FunctionNameBox, "「" + AiConfigStore.AgentFunctionName + "」是 Agent 自己的名字，文字功能换一个名字。");
                if (functions.Any(f => string.Equals(f.Name, name, StringComparison.OrdinalIgnoreCase)))
                    return ShowFunctionError(item, FunctionNameBox, "有两个文字功能都叫「" + name + "」，改一个名字。");
                if (prompt.Length == 0)
                    return ShowFunctionError(item, FunctionPromptBox, "「" + name + "」的提示词是空的。");
                functions.Add(new AiFunction(name, prompt, item.RemoveExtraBlankLines));
            }

            var agent = new AgentOptions();
            var defaultRequest = AgentOptions.NormalizeRequest(AgentDefaultRequestBox.Text);
            if (defaultRequest.Length == 0 || defaultRequest.Length > AgentOptions.MaxRequestLength)
                return ShowError(AgentPageIndex, AgentDefaultRequestBox, $"默认需求要填 1–{AgentOptions.MaxRequestLength} 字的内容。");
            agent.DefaultRequest = defaultRequest;
            foreach (var option in AgentOptions.Numbers)
            {
                var box = _numberBoxes[option.Key];
                if (!TryReadNumber(box, option.Min, option.Max, out var value))
                    return ShowError(AgentPageIndex, box, $"「{Label(NumberLabels, option.Key)}」要填 {option.Min}–{option.Max} 之间的整数。");
                option.Set(agent, value);
            }

            foreach (var option in AgentOptions.Switches)
                option.Set(agent, _switchBoxes[option.Key].IsChecked == true);
            agent.FontFamily = FontPicker.SelectedItem as string ?? AgentOptions.DefaultFontFamily;

            config = new AiConfig(url, ApiKeyValue.Trim(), timeout, maxTokens,
                _models.Select(m => new AiModel(m)).ToList(), functions) { Agent = agent };
            return true;
        }

        private static bool TryReadNumber(TextBox box, int min, int max, out int value) =>
            int.TryParse(box.Text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out value) && value >= min && value <= max;

        private bool ShowFunctionError(FunctionItem item, Control focus, string message)
        {
            SelectFunction(item);
            return ShowError(FunctionsPageIndex, focus, message);
        }

        /// <summary>切到出错的页、把焦点放到出错的框上，底部红字说明。总是返回 false，方便校验里直接 return。</summary>
        private bool ShowError(int? page, Control focus, string message)
        {
            if (page != null) ShowPage(page.Value);
            StatusText.Text = message;
            StatusText.Visibility = Visibility.Visible;
            if (focus != null)
            {
                // 刚切过去的页还没排版，等排完再给焦点。
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    focus.Focus();
                    (focus as TextBox)?.SelectAll();
                }), System.Windows.Threading.DispatcherPriority.Input);
            }

            return false;
        }

        private void ClearError()
        {
            StatusText.Text = string.Empty;
            StatusText.Visibility = Visibility.Collapsed;
        }

        /// <summary>有没保存的修改时问一句：保存、不保存、取消。返回 true 表示可以继续（已保存或放弃修改）。</summary>
        private bool ConfirmLeave(string question)
        {
            var answer = MessageBox.Show(this, question, Caption, MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
            return answer == MessageBoxResult.No || (answer == MessageBoxResult.Yes && Save());
        }

        #endregion

        /// <summary>
        /// 用系统默认的程序打开 ai-settings.xml（没有就先生成默认的），然后关掉本窗口，免得两边同时改同一个文件。
        /// .xml 没有关联任何程序时退回记事本。
        /// </summary>
        private void OnOpenInEditor(object sender, RoutedEventArgs e)
        {
            ClearError();
            if (IsDirty && !ConfirmLeave("窗口里的修改还没保存，先保存再打开吗？\n\n选「否」放弃这些修改。")) return;

            try
            {
                var path = AiConfigStore.EnsureFile();
                try
                {
                    Process.Start(new ProcessStartInfo(path) { UseShellExecute = true })?.Dispose();
                }
                catch (Win32Exception ex)
                {
                    AddInLog.Warn(".xml 没有可用的默认程序，改用记事本打开。", ex);
                    Process.Start(new ProcessStartInfo("notepad.exe", "\"" + path + "\"") { UseShellExecute = false })?.Dispose();
                }
            }
            catch (Exception ex)
            {
                AddInLog.Error("打开 AI 配置文件失败。", ex);
                ShowError(null, null, "打开配置文件失败：" + ex.Message);
                return;
            }

            _discard = true;
            Close();
        }

        /// <summary>文字功能列表里的一项。右边编辑区改动时由 OnFunctionEdited 写进来，改名字时通知列表刷新显示。</summary>
        private sealed class FunctionItem : INotifyPropertyChanged
        {
            private string _name = string.Empty;

            public event PropertyChangedEventHandler PropertyChanged;

            public string Name
            {
                get => _name;
                set
                {
                    _name = value ?? string.Empty;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Name)));
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(DisplayName)));
                }
            }

            public string Prompt { get; set; } = string.Empty;

            public bool RemoveExtraBlankLines { get; set; }

            /// <summary>列表里显示的名字，名称还空着时给个占位。</summary>
            public string DisplayName => string.IsNullOrWhiteSpace(Name) ? "（未命名）" : Name.Trim();
        }
    }
}
