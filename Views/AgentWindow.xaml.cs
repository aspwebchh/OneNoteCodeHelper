using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using OneNoteCodeHelper.Interop;
using OneNoteCodeHelper.Services;
using OneNoteCodeHelper.Services.Agent;

namespace OneNoteCodeHelper.Views
{
    /// <summary>
    /// AI 助手窗口：「功能」下拉选 Agent 自定义排版，或 AI 配置里的文字功能（智能校正等），
    /// 模型、思考强度也在这里选。页面和选区在打开窗口时固定。
    ///
    /// 在 COM 代理进程的独立 STA 线程上 ShowDialog；真正的活（读页面、调接口、写回）在线程池上跑。
    /// </summary>
    public partial class AgentWindow : Window
    {
        private const string ThinkingPlaceholder = "等待模型输出…";

        private const string AgentIntro =
            "Agent 会读取所选范围并调用格式工具调整字体、标题、间距等，完成后显示回读核验的结果；本窗口内可以撤销最近一次修改。";

        private readonly IOneNotePageAccess _api;
        private readonly string _pageId;
        private readonly HashSet<string> _selection;
        /// <summary>选区里的空行，文字功能删空行时只删这些。</summary>
        private readonly HashSet<string> _selectedBlankLines;
        /// <summary>功能、模型、思考强度的当前选择，以及代码框的主题、字号等（取自打开窗口时的功能区设置）。</summary>
        private readonly AddInSettings _settings;
        private AiConfig _config;
        /// <summary>读 _config 时 ai-settings.xml 的修改时间，变了才重新读。</summary>
        private DateTime _configStamp;
        private readonly CancellationTokenSource _lifetime = new CancellationTokenSource();
        private readonly Stopwatch _elapsed = new Stopwatch();
        private CancellationTokenSource _run;
        private Task _job;
        /// <summary>Agent 最近一次执行或撤销的结果，撤销用。文字功能不能撤销，执行时清掉。</summary>
        private AgentReport _report;
        /// <summary>处理期间每秒刷新一次「已用时」。</summary>
        private DispatcherTimer _ticker;
        /// <summary>本次执行的步骤列表，下一次执行或撤销开始时清空。</summary>
        private readonly ObservableCollection<StepRow> _steps = new ObservableCollection<StepRow>();
        /// <summary>「已用时」前面的实时说明：Agent 是第几轮，文字功能是 AI 在思考 / 输出 / 已返回几段。null 不显示。</summary>
        private string _liveDetail;
        private bool _busy;
        private bool _closed;
        /// <summary>正在往下拉里填选项，这时的 SelectionChanged 不算用户改选。</summary>
        private bool _loadingChoices;

        internal AgentWindow(IOneNotePageAccess api, string pageId, string selectionXml, AiConfig config,
            AddInSettings settings, IntPtr owner)
        {
            InitializeComponent();
            _api = api; _pageId = pageId; _config = config; _settings = settings;
            _configStamp = ConfigStamp();
            var page = AgentPageSnapshot.ParsePage(selectionXml);
            _selection = AgentPageSnapshot.SelectedIds(page);
            _selectedBlankLines = PageEditor.FindSelectedBlankLines(page);
            var name = (string)page.Attribute("name");
            PageText.Text = "目标页面：" + (string.IsNullOrWhiteSpace(name) ? "当前页" : name);
            SelectionScope.IsEnabled = _selection.Count > 0;
            StepsList.ItemsSource = _steps;
            LoadChoices();

            // 屏幕太矮时收一收，状态卡片里的滚动区跟着变矮。
            var limit = SystemParameters.WorkArea.Height - 16;
            if (Height > limit) Height = Math.Max(MinHeight, limit);

            AppIcon.Source = LoadIcon("Agent");
            // 不设的话标题栏上是宿主进程（dllhost）的图标。
            if (AppIcon.Source != null) Icon = AppIcon.Source;
            var interop = new WindowInteropHelper(this);
            if (owner != IntPtr.Zero) interop.Owner = owner;
            SourceInitialized += (_, __) =>
            {
                // 标题栏刷成窗口底色，和下面连成一整块（Windows 11 才有效果）。
                NativeMethods.TrySetCaptionColor(interop.Handle, ((SolidColorBrush)Background).Color);
                if (owner != IntPtr.Zero) NativeMethods.CenterOver(interop.Handle, owner);
            };
            // 用户可能开着窗口去改了 AI 配置，切回来时重新读。
            Activated += (_, __) => RefreshConfig();
            Closing += (_, __) => _lifetime.Cancel();
            Closed += (_, __) => { _closed = true; StopRunning(); };
            // 下拉展开时 Esc 只收起下拉，不关窗。
            PreviewKeyDown += (_, e) =>
            {
                if (e.Key == Key.Escape && !FunctionPicker.IsDropDownOpen && !ModelPicker.IsDropDownOpen && !EffortPicker.IsDropDownOpen)
                    Close();
            };
        }

        /// <summary>用户在窗口里改过功能、模型或思考强度，关窗后由调用方存进设置。</summary>
        internal bool ChoicesChanged { get; private set; }

        internal string FunctionName => _settings.AgentFunction;

        internal string ModelId => _settings.AiModel;

        internal string Effort => _settings.AiEffort;

        internal void CancelForShutdown()
        {
            _lifetime.Cancel();
            if (!Dispatcher.HasShutdownStarted && !Dispatcher.HasShutdownFinished) Dispatcher.BeginInvoke(new Action(Close));
        }

        internal void WaitForJob()
        {
            try { _job?.GetAwaiter().GetResult(); } catch (Exception) { /* UI 负责显示；关闭时仅收尾。 */ }
        }

        /// <summary>按 _config 填三个下拉，尽量保留原来的选择，找不到就选第一项。</summary>
        private void LoadChoices()
        {
            _loadingChoices = true;
            try
            {
                var functions = new List<string> { AiConfigStore.AgentFunctionName };
                foreach (var function in _config.Functions)
                {
                    if (string.Equals(function.Name, AiConfigStore.AgentFunctionName, StringComparison.OrdinalIgnoreCase))
                    {
                        AddInLog.Warn("AI 配置里有和 Agent 同名的功能，窗口里跳过：" + function.Name);
                        continue;
                    }

                    functions.Add(function.Name);
                }

                FunctionPicker.ItemsSource = functions;
                FunctionPicker.SelectedItem = functions.FirstOrDefault(n =>
                    string.Equals(n, _settings.AgentFunction, StringComparison.OrdinalIgnoreCase)) ?? functions[0];
                ModelPicker.ItemsSource = _config.Models.Select(m => m.Id).ToList();
                ModelPicker.SelectedItem = _config.FindModel(_settings.AiModel).Id;
                EffortPicker.ItemsSource = AiEfforts.Ids;
                EffortPicker.SelectedItem = AiEfforts.Normalize(_settings.AiEffort);
            }
            finally
            {
                _loadingChoices = false;
            }

            // 兜底选中的项和设置里存的不一样时，以界面上的为准，但不算用户改过。
            ReadChoices();
            ShowFunction();
        }

        private void ReadChoices()
        {
            _settings.AgentFunction = (string)FunctionPicker.SelectedItem;
            _settings.AiModel = (string)ModelPicker.SelectedItem;
            _settings.AiEffort = (string)EffortPicker.SelectedItem;
        }

        private void OnChoiceChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_loadingChoices || FunctionPicker.SelectedItem == null || ModelPicker.SelectedItem == null || EffortPicker.SelectedItem == null)
                return;
            ReadChoices();
            ChoicesChanged = true;
            if (ReferenceEquals(sender, FunctionPicker)) ShowFunction();
        }

        /// <summary>选中的文字功能，选的是 Agent 时为 null。</summary>
        private AiFunction TextFunction()
        {
            return string.Equals(_settings.AgentFunction, AiConfigStore.AgentFunctionName, StringComparison.Ordinal)
                ? null
                : _config.FindFunction(_settings.AgentFunction);
        }

        /// <summary>按所选功能切换需求框 / 提示词框和说明。两个框同样高，切换时下面的内容不动。</summary>
        private void ShowFunction()
        {
            var function = TextFunction();
            RequestText.Visibility = function == null ? Visibility.Visible : Visibility.Collapsed;
            PromptText.Visibility = function == null ? Visibility.Collapsed : Visibility.Visible;
            PromptText.Text = function?.Prompt ?? string.Empty;
            RequestLabel.Text = function == null ? "需求" : "提示词";
            RequestHint.Text = function == null ? "只调整格式，保留正文、代码和段落结构" : "在「AI 配置」中修改";
            IntroText.Text = function == null
                ? AgentIntro
                : "AI 按「" + function.Name + "」逐段修改文字，结果直接写回并列出改动清单" +
                  (function.RemoveExtraBlankLines ? "，顺带删掉多余的空行" : "") +
                  "。代码框不会交给 AI；处理期间你改过的段落会跳过，不会覆盖。";
        }

        /// <summary>ai-settings.xml 改过就重新读。读不了（编辑器还在写、XML 写坏了）先留着旧配置，下次再读。</summary>
        private void RefreshConfig()
        {
            if (_busy || _closed) return;
            var stamp = ConfigStamp();
            if (stamp == _configStamp || !AiConfigStore.TryLoad(out var config, out _)) return;
            _configStamp = stamp;
            var previous = _config;
            _config = config;
            if (config.HasSameChoices(previous)) ShowFunction();
            else LoadChoices();
            AddInLog.Info("Agent 窗口已重新读取 AI 配置。");
        }

        private static DateTime ConfigStamp()
        {
            try { return File.GetLastWriteTimeUtc(AiConfigStore.ConfigPath); }
            catch (Exception) { return DateTime.MinValue; }
        }

        private async void OnExecute(object sender, RoutedEventArgs e)
        {
            if (_busy) return;
            RefreshConfig();
            var config = _config;
            var function = TextFunction();
            var request = RequestText.Text.Trim();
            if (function == null && request.Length == 0) { ShowOutcome(WarningIcon, "请先输入需求。", null); return; }
            if (string.IsNullOrWhiteSpace(config.ApiKey)) { ShowOutcome(WarningIcon, "请先在「AI 配置」中填写 ApiKey。", "保存后回到本窗口再点「执行」。"); return; }
            var model = config.FindModel(_settings.AiModel).Id;
            var effort = AiEfforts.Normalize(_settings.AiEffort);
            var selectionOnly = SelectionScope.IsChecked == true;
            var selection = selectionOnly ? new HashSet<string>(_selection) : null;

            if (function != null)
            {
                var blankLines = selectionOnly ? new HashSet<string>(_selectedBlankLines) : null;
                var optimizer = new AiOptimizer(_api, config, function, model, effort);
                await StartJob(token => optimizer.RunAsync(_pageId, selection, blankLines, CreateProgress<AiProgress>(ShowAiProgress), token),
                    "正在读取固定目标页…", true, ShowAiReport);
                return;
            }

            var codeSettings = _settings.Clone();
            await StartJob(token =>
            {
                var xml = _api.GetPageContent(_pageId, Microsoft.Office.Interop.OneNote.PageInfo.piBasic);
                var snapshot = new AgentPageSnapshot(xml, selection, config.Agent);
                if (!snapshot.Blocks.Any(b => b.Editable || b.CodeCandidate)) throw new AiException("目标范围没有可编辑文字，或所在文本框包含尚未启用的混合内容。");
                if (!AgentTools.FontInstalled(config.Agent.FontFamily))
                    config.Agent.FontFamily = ParagraphStyles.Fonts.FirstOrDefault(AgentTools.FontInstalled)
                        ?? throw new AiException("Agent 默认字体未安装，请在 AI 配置的 Agent/FontFamily 中选择已安装字体。");
                var runner = new AgentRunner(new AgentChatClient(config, model, effort), new AgentCommitter(_api), codeSettings);
                return runner.RunAsync(snapshot, request, CreateProgress<AgentProgress>(ShowAgentProgress), token);
            }, "正在读取固定目标页…", false, report => { _report = report; ShowReport(report, false); });
        }

        private IProgress<T> CreateProgress<T>(Action<T, bool> show)
        {
            // 在后台调用此工厂也只通过 Dispatcher 更新 UI，不捕获线程池上下文。流式进度已在各自的客户端里限频。
            var run = _run;
            return new DirectProgress<T>(progress =>
            {
                if (_closed) return;
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (!_closed && _busy && ReferenceEquals(_run, run)) show(progress, run.IsCancellationRequested);
                }));
            });
        }

        private sealed class DirectProgress<T> : IProgress<T>
        {
            private readonly Action<T> _report;
            internal DirectProgress(Action<T> report) { _report = report; }
            public void Report(T value) => _report(value);
        }

        private void ShowAgentProgress(AgentProgress progress, bool cancelling)
        {
            // 已点取消：状态行和说明留给「正在取消…」，步骤照常更新，写回中的那步要显示出结果。
            if (progress.Step != null) ShowStep(progress.Step);
            if (cancelling) return;
            if (progress.Turn > 0) { _liveDetail = $"第 {progress.Turn} 轮"; ShowElapsed(); }
            if (progress.Status != null) StatusText.Text = progress.Status;
            if (progress.Thinking != null) ShowThinking(progress.Thinking);
        }

        private void ShowAiProgress(AiProgress progress, bool cancelling)
        {
            if (cancelling) return;
            StatusText.Text = progress.Message;
            _liveDetail = progress.Detail;
            ShowElapsed();
            ShowThinking(progress.Thinking);
        }

        /// <summary>思考摘录。处理期间摘录框一直占着位置，没有文字时显示占位，不收起。</summary>
        private void ShowThinking(string text)
        {
            var empty = string.IsNullOrEmpty(text);
            ThinkingText.Text = empty ? ThinkingPlaceholder : text;
            ThinkingText.Foreground = (Brush)FindResource(empty ? "Faint" : "Muted");
        }

        /// <summary>按 Id 新增或替换一行步骤；新增时滚到最下面。</summary>
        private void ShowStep(AgentStep step)
        {
            var row = new StepRow(step.Id, step.Text, step.State, StepBrush(step.State));
            for (var i = 0; i < _steps.Count; i++)
            {
                if (_steps[i].Id != step.Id) continue;
                _steps[i] = row;
                return;
            }
            _steps.Add(row);
            StepsPanel.Visibility = Visibility.Visible;
            ContentScroll.ScrollToEnd();
        }

        /// <summary>任务没做完就结束了（失败或取消）：还显示「执行中」的步骤改成失败。</summary>
        private void InterruptSteps()
        {
            for (var i = 0; i < _steps.Count; i++)
                if (_steps[i].State == AgentStepState.Running)
                    _steps[i] = new StepRow(_steps[i].Id, _steps[i].Text, AgentStepState.Failed, StepBrush(AgentStepState.Failed));
        }

        private Brush StepBrush(AgentStepState state)
        {
            switch (state)
            {
                case AgentStepState.Done: return (Brush)FindResource("Success");
                case AgentStepState.Failed: return (Brush)FindResource("Danger");
                case AgentStepState.Note: return (Brush)FindResource("Faint");
                default: return (Brush)FindResource("Primary");
            }
        }

        /// <summary>步骤列表里的一行，只读；状态变了就整行替换，不用属性通知。</summary>
        private sealed class StepRow
        {
            internal StepRow(int id, string text, AgentStepState state, Brush brush) { Id = id; Text = text; State = state; Brush = brush; }
            public int Id { get; }
            public string Text { get; }
            public AgentStepState State { get; }
            public Brush Brush { get; }
            public string Glyph
            {
                get
                {
                    switch (State)
                    {
                        case AgentStepState.Done: return "✓";
                        case AgentStepState.Failed: return "✕";
                        case AgentStepState.Note: return "·";
                        default: return "…";
                    }
                }
            }
        }

        /// <summary>在线程池上跑 job，结果交给 show。textFunction 决定取消、失败时的说明和日志写法。</summary>
        private async Task StartJob<T>(Func<CancellationToken, Task<T>> job, string status, bool textFunction, Action<T> show)
        {
            _run?.Dispose();
            _run = CancellationTokenSource.CreateLinkedTokenSource(_lifetime.Token);
            var token = _run.Token;
            // 下一次执行开始后，上一次的撤销记录作废。
            _report = null;
            SetBusy(true);
            StartRunning(status);
            try
            {
                var task = Task.Run(() => job(token));
                _job = task;
                var result = await task;
                if (_closed) return;
                show(result);
            }
            catch (OperationCanceledException)
            {
                if (!_closed) { InterruptSteps(); ShowOutcome(InfoIcon, "已取消", textFunction ? "页面没有改动。" : "未提交本次草稿。"); }
            }
            catch (Exception ex)
            {
                string message;
                if (!textFunction)
                {
                    message = ex is AiException ? ex.Message : "处理失败。请检查页面状态后重试。";
                    // 不记录模型原文、工具参数或笔记内容。
                    AddInLog.Info("Agent 失败，类型=" + ex.GetType().Name);
                }
                else if (ex is AiException)
                {
                    message = ex.Message;
                    AddInLog.Warn("AI 文字功能失败：" + ex.Message, ex.InnerException);
                }
                else
                {
                    message = $"处理失败：{ex.Message}\n\n详情见日志：{AddInLog.LogPath}";
                    AddInLog.Error("AI 文字功能失败。", ex);
                }

                if (!_closed) { InterruptSteps(); ShowFailure(message); }
            }
            finally { if (!_closed) SetBusy(false); }
        }

        private async void OnUndo(object sender, RoutedEventArgs e)
        {
            if (_busy || _report == null || !_report.CanUndo) return;
            var previous = _report;
            var options = _config.Agent;
            await StartJob(token => Task.FromResult(new AgentCommitter(_api).Undo(_pageId, previous, options, token)),
                "正在撤销…", false, report => { _report = report; ShowReport(report, true); });
            // 只撤销最近一次执行；不把撤销的逆操作继续暴露为撤销。
            if (_report != null) { _report.Undo.Clear(); _report.CodeUndo.Clear(); }
            if (!_closed) UndoButton.Visibility = Visibility.Collapsed;
        }

        private void OnCancel(object sender, RoutedEventArgs e)
        {
            _run?.Cancel();
            // 停掉「已用时」刷新，免得盖掉下面这行说明；总用时照常计。
            _ticker?.Stop();
            ShowThinking(null);
            StatusText.Text = "正在取消…";
            ShowDetail("如果已开始写回，将先核对实际结果。");
            CancelButton.IsEnabled = false;
        }

        private void SetBusy(bool value)
        {
            _busy = value;
            ExecuteButton.IsEnabled = !value;
            CancelButton.IsEnabled = value;
            CancelButton.Visibility = value ? Visibility.Visible : Visibility.Collapsed;
            UndoButton.Visibility = !value && _report?.CanUndo == true ? Visibility.Visible : Visibility.Collapsed;
            FunctionPicker.IsEnabled = !value;
            ModelPicker.IsEnabled = !value;
            EffortPicker.IsEnabled = !value;
            RequestText.IsEnabled = !value;
            PageScope.IsEnabled = !value;
            SelectionScope.IsEnabled = !value && _selection.Count > 0;
        }

        /// <summary>开始转圈和计时，收起说明、上一次的结果和步骤，摘录框从现在起一直占着位置。</summary>
        private void StartRunning(string status)
        {
            ShowIcon(SpinnerIcon);
            StatusText.Text = status;
            IntroText.Visibility = Visibility.Collapsed;
            ResultPanel.Visibility = Visibility.Collapsed;
            ErrorBox.Visibility = Visibility.Collapsed;
            _steps.Clear();
            StepsPanel.Visibility = Visibility.Collapsed;
            ThinkingBox.Visibility = Visibility.Visible;
            ShowThinking(null);
            _liveDetail = null;
            _elapsed.Restart();
            ShowElapsed();
            if (_ticker == null)
            {
                _ticker = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
                _ticker.Tick += (_, __) => ShowElapsed();
            }
            _ticker.Start();
            SpinnerRotation.BeginAnimation(RotateTransform.AngleProperty,
                new DoubleAnimation(0, 360, TimeSpan.FromSeconds(0.9)) { RepeatBehavior = RepeatBehavior.Forever });
        }

        private void StopRunning()
        {
            _ticker?.Stop();
            _elapsed.Stop();
            SpinnerRotation.BeginAnimation(RotateTransform.AngleProperty, null);
        }

        private void ShowReport(AgentReport report, bool undo)
        {
            var summary = $"用时 {ElapsedSeconds()} 秒";
            switch (report.Status)
            {
                case "Verified":
                    ShowOutcome(SuccessIcon, undo ? "已撤销" : "已完成", summary);
                    break;
                case "PartiallyApplied":
                    ShowOutcome(WarningIcon, undo ? "部分撤销" : "部分完成", summary);
                    break;
                case "CommitOutcomeUnknown":
                    ShowOutcome(WarningIcon, "写回未确认", summary);
                    break;
                default:
                    if (report.Conflicts > 0) ShowOutcome(WarningIcon, "改动与页面冲突，未写入", summary);
                    else ShowOutcome(InfoIcon, "没有写入修改", summary);
                    break;
            }

            ShowResult("结果", report.Message);
        }

        private void ShowAiReport(AiReport report)
        {
            if (!report.Success) { ShowFailure(report.Error); return; }

            var summary = $"{report.Scope} · 用时 {ElapsedSeconds()} 秒";
            if (report.Changed) ShowOutcome(SuccessIcon, "已完成", summary);
            else if (report.Conflicted > 0) ShowOutcome(WarningIcon, "没有写回任何改动", summary);
            else ShowOutcome(InfoIcon, "没有需要修改的地方", summary);

            // 只删了空行、AI 没改字时不显示清单；改了字却没写说明时，清单位置只放一行提示。
            ShowResult("改动清单", report.Changes.Count > 0
                ? string.Join("\n", report.Changes.Select(c => "• " + c))
                : report.Applied > 0 ? "AI 没有给出改动说明。" : null);
        }

        private void ShowResult(string label, string text)
        {
            ResultLabel.Text = label;
            ResultText.Text = text ?? string.Empty;
            ResultPanel.Visibility = string.IsNullOrWhiteSpace(text) ? Visibility.Collapsed : Visibility.Visible;
            ContentScroll.ScrollToEnd();
        }

        private void ShowFailure(string message)
        {
            ShowOutcome(ErrorIcon, "没能完成", $"用时 {ElapsedSeconds()} 秒");
            ErrorText.Text = message;
            ErrorBox.Visibility = Visibility.Visible;
            ContentScroll.ScrollToEnd();
        }

        /// <summary>停掉转圈，换成结果图标和标题，收起摘录框；结果框和错误框先收起，由调用方按需打开。步骤列表保留。</summary>
        private void ShowOutcome(FrameworkElement icon, string status, string detail)
        {
            StopRunning();
            ShowIcon(icon);
            StatusText.Text = status;
            ShowDetail(detail);
            ThinkingBox.Visibility = Visibility.Collapsed;
            ResultPanel.Visibility = Visibility.Collapsed;
            ErrorBox.Visibility = Visibility.Collapsed;
        }

        private void ShowElapsed() => ShowDetail((_liveDetail != null ? _liveDetail + " · " : "") + $"已用时 {(int)_elapsed.Elapsed.TotalSeconds} 秒");

        /// <summary>说明行始终占着一行，没有内容时留空，状态行的高度不变。</summary>
        private void ShowDetail(string text) => DetailText.Text = text ?? string.Empty;

        private void ShowIcon(FrameworkElement icon)
        {
            foreach (var candidate in new FrameworkElement[] { SpinnerIcon, SuccessIcon, WarningIcon, InfoIcon, ErrorIcon })
            {
                candidate.Visibility = candidate == icon ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        private int ElapsedSeconds() => (int)Math.Round(_elapsed.Elapsed.TotalSeconds);

        /// <summary>读嵌入资源里的图标（和功能区用的是同一张），找不到就不显示。</summary>
        private static ImageSource LoadIcon(string name)
        {
            try
            {
                var assembly = typeof(AgentWindow).Assembly;
                var resource = assembly.GetManifestResourceNames()
                    .FirstOrDefault(n => n.EndsWith("." + name + ".png", StringComparison.OrdinalIgnoreCase));
                if (resource == null) return null;

                using (var stream = assembly.GetManifestResourceStream(resource))
                {
                    var image = new BitmapImage();
                    image.BeginInit();
                    image.CacheOption = BitmapCacheOption.OnLoad;
                    image.StreamSource = stream;
                    image.EndInit();
                    image.Freeze();
                    return image;
                }
            }
            catch (Exception ex)
            {
                AddInLog.Warn("读取窗口图标失败：" + name, ex);
                return null;
            }
        }
    }
}
