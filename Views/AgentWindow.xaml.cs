using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using OneNoteCodeHelper.Interop;
using OneNoteCodeHelper.Services;
using OneNoteCodeHelper.Services.Agent;

namespace OneNoteCodeHelper.Views
{
    public partial class AgentWindow : Window
    {
        private readonly IOneNotePageAccess _api;
        private readonly string _pageId;
        private readonly HashSet<string> _selection;
        private readonly AiConfig _config;
        private readonly string _model;
        private readonly string _effort;
        /// <summary>代码框的主题、字号等，取自打开窗口时的功能区设置。</summary>
        private readonly AddInSettings _codeSettings;
        private readonly CancellationTokenSource _lifetime = new CancellationTokenSource();
        private readonly Stopwatch _elapsed = new Stopwatch();
        private CancellationTokenSource _run;
        private Task<AgentReport> _job;
        private AgentReport _report;
        /// <summary>处理期间每秒刷新一次「已用时」。</summary>
        private DispatcherTimer _ticker;
        /// <summary>本次执行的步骤列表，下一次执行或撤销开始时清空。</summary>
        private readonly ObservableCollection<StepRow> _steps = new ObservableCollection<StepRow>();
        /// <summary>模型当前是第几轮，0 表示还没开始。</summary>
        private int _turn;
        private bool _busy;
        private bool _closed;

        internal AgentWindow(IOneNotePageAccess api, string pageId, string selectionXml, AiConfig config, string model, string effort,
            AddInSettings codeSettings, IntPtr owner)
        {
            InitializeComponent();
            _api = api; _pageId = pageId; _config = config; _model = model; _effort = effort; _codeSettings = codeSettings;
            var page = AgentPageSnapshot.ParsePage(selectionXml);
            _selection = AgentPageSnapshot.SelectedIds(page);
            PageText.Text = (string)page.Attribute("name") ?? "当前页";
            ModelText.Text = model + " · 思考 " + effort;
            SelectionScope.IsEnabled = _selection.Count > 0;
            StepsList.ItemsSource = _steps;
            AppIcon.Source = AiProgressWindow.LoadIcon("Agent");
            // 不设的话标题栏上是宿主进程（dllhost）的图标。
            if (AppIcon.Source != null) Icon = AppIcon.Source;
            var interop = new WindowInteropHelper(this);
            if (owner != IntPtr.Zero) interop.Owner = owner;
            SourceInitialized += (_, __) =>
            {
                // 标题栏底色和居中方式同「AI 优化」进度窗。
                NativeMethods.TrySetCaptionColor(interop.Handle, ((SolidColorBrush)Background).Color);
                if (owner != IntPtr.Zero) NativeMethods.CenterOver(interop.Handle, owner);
            };
            Closing += (_, __) => _lifetime.Cancel();
            Closed += (_, __) => { _closed = true; StopRunning(); };
            PreviewKeyDown += (_, e) => { if (e.Key == Key.Escape) Close(); };
        }

        internal void CancelForShutdown()
        {
            _lifetime.Cancel();
            if (!Dispatcher.HasShutdownStarted && !Dispatcher.HasShutdownFinished) Dispatcher.BeginInvoke(new Action(Close));
        }

        internal void WaitForJob()
        {
            try { _job?.GetAwaiter().GetResult(); } catch (Exception) { /* UI 负责显示；关闭时仅收尾。 */ }
        }

        private async void OnExecute(object sender, RoutedEventArgs e)
        {
            if (_busy) return;
            var request = RequestText.Text.Trim();
            if (request.Length == 0) { ShowOutcome(WarningIcon, "请先输入需求。", null); return; }
            if (string.IsNullOrWhiteSpace(_config.ApiKey)) { ShowOutcome(WarningIcon, "请先在「AI 配置」中填写 ApiKey。", "保存配置后重开 Agent 窗口生效。"); return; }
            var selection = SelectionScope.IsChecked == true ? new HashSet<string>(_selection) : null;
            await StartJob(token =>
            {
                var xml = _api.GetPageContent(_pageId, Microsoft.Office.Interop.OneNote.PageInfo.piBasic);
                var snapshot = new AgentPageSnapshot(xml, selection, _config.Agent);
                if (!snapshot.Blocks.Any(b => b.Editable || b.CodeCandidate)) throw new AiException("目标范围没有可编辑文字，或所在文本框包含尚未启用的混合内容。");
                if (!AgentTools.FontInstalled(_config.Agent.FontFamily))
                    _config.Agent.FontFamily = ParagraphStyles.Fonts.FirstOrDefault(AgentTools.FontInstalled)
                        ?? throw new AiException("Agent 默认字体未安装，请在 AI 配置的 Agent/FontFamily 中选择已安装字体。");
                var runner = new AgentRunner(new AgentChatClient(_config, _model, _effort), new AgentCommitter(_api), _codeSettings);
                return runner.RunAsync(snapshot, request, CreateProgress(), token);
            }, false);
        }

        private IProgress<AgentProgress> CreateProgress()
        {
            // 在后台调用此工厂也只通过 Dispatcher 更新 UI，不捕获线程池上下文。流式进度已在 AgentChatClient 里限频。
            var run = _run;
            return new DirectProgress(progress =>
            {
                if (_closed) return;
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (!_closed && _busy && ReferenceEquals(_run, run)) ShowProgress(progress, run.IsCancellationRequested);
                }));
            });
        }

        private sealed class DirectProgress : IProgress<AgentProgress>
        {
            private readonly Action<AgentProgress> _report;
            internal DirectProgress(Action<AgentProgress> report) { _report = report; }
            public void Report(AgentProgress value) => _report(value);
        }

        private void ShowProgress(AgentProgress progress, bool cancelling)
        {
            // 已点取消：状态行和说明留给「正在取消…」，步骤照常更新，写回中的那步要显示出结果。
            if (progress.Step != null) ShowStep(progress.Step);
            if (cancelling) return;
            if (progress.Turn > 0) { _turn = progress.Turn; ShowElapsed(); }
            if (progress.Status != null) StatusText.Text = progress.Status;
            if (progress.Thinking != null) ShowThinking(progress.Thinking);
        }

        /// <summary>思考摘录，空串收起摘录框。</summary>
        private void ShowThinking(string text)
        {
            ThinkingText.Text = text;
            ThinkingBox.Visibility = text.Length == 0 ? Visibility.Collapsed : Visibility.Visible;
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
            StepsScroll.ScrollToEnd();
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

        private async Task StartJob(Func<CancellationToken, Task<AgentReport>> job, bool undo)
        {
            _run?.Dispose();
            _run = CancellationTokenSource.CreateLinkedTokenSource(_lifetime.Token);
            SetBusy(true);
            StartRunning(undo ? "正在撤销…" : "正在读取固定目标页…");
            try
            {
                _job = Task.Run(() => job(_run.Token));
                _report = await _job;
                if (_closed) return;
                ShowReport(_report, undo);
            }
            catch (OperationCanceledException)
            {
                _report = null;
                if (!_closed) { InterruptSteps(); ShowOutcome(InfoIcon, "已取消", "未提交本次草稿。"); }
            }
            catch (Exception ex)
            {
                _report = null;
                if (!_closed) { InterruptSteps(); ShowFailure(ex is AiException ? ex.Message : "处理失败。请检查页面状态后重试。"); }
                // 不记录模型原文、工具参数或笔记内容。
                AddInLog.Info("Agent 失败，类型=" + ex.GetType().Name);
            }
            finally { if (!_closed) SetBusy(false); }
        }

        private async void OnUndo(object sender, RoutedEventArgs e)
        {
            if (_busy || _report == null || !_report.CanUndo) return;
            var previous = _report;
            await StartJob(token => Task.FromResult(new AgentCommitter(_api).Undo(_pageId, previous, _config.Agent, token)), true);
            // 只撤销最近一次执行；不把撤销的逆操作继续暴露为撤销。
            if (_report != null) { _report.Undo.Clear(); _report.CodeUndo.Clear(); }
            if (!_closed) UndoButton.Visibility = Visibility.Collapsed;
        }

        private void OnCancel(object sender, RoutedEventArgs e)
        {
            _run?.Cancel();
            // 停掉「已用时」刷新，免得盖掉下面这行说明；总用时照常计。
            _ticker?.Stop();
            ShowThinking("");
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
            RequestText.IsEnabled = !value;
            PageScope.IsEnabled = !value;
            SelectionScope.IsEnabled = !value && _selection.Count > 0;
        }

        /// <summary>开始转圈和计时，收起上一次的结果和步骤。</summary>
        private void StartRunning(string status)
        {
            ShowIcon(SpinnerIcon);
            StatusText.Text = status;
            ResultPanel.Visibility = Visibility.Collapsed;
            ErrorBox.Visibility = Visibility.Collapsed;
            _steps.Clear();
            StepsPanel.Visibility = Visibility.Collapsed;
            ShowThinking("");
            _turn = 0;
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

            ResultText.Text = report.Message;
            ResultPanel.Visibility = string.IsNullOrWhiteSpace(report.Message) ? Visibility.Collapsed : Visibility.Visible;
        }

        private void ShowFailure(string message)
        {
            ShowOutcome(ErrorIcon, "没能完成", $"用时 {ElapsedSeconds()} 秒");
            ErrorText.Text = message;
            ErrorBox.Visibility = Visibility.Visible;
        }

        /// <summary>停掉转圈，换成结果图标和标题；结果框和错误框先收起，由调用方按需打开。步骤列表保留。</summary>
        private void ShowOutcome(FrameworkElement icon, string status, string detail)
        {
            StopRunning();
            ShowIcon(icon);
            StatusText.Text = status;
            ShowDetail(detail);
            ShowThinking("");
            ResultPanel.Visibility = Visibility.Collapsed;
            ErrorBox.Visibility = Visibility.Collapsed;
        }

        private void ShowElapsed() => ShowDetail((_turn > 0 ? $"第 {_turn} 轮 · " : "") + $"已用时 {(int)_elapsed.Elapsed.TotalSeconds} 秒");

        private void ShowDetail(string text)
        {
            DetailText.Text = text ?? "";
            DetailText.Visibility = text == null ? Visibility.Collapsed : Visibility.Visible;
        }

        private void ShowIcon(FrameworkElement icon)
        {
            foreach (var candidate in new FrameworkElement[] { SpinnerIcon, SuccessIcon, WarningIcon, InfoIcon, ErrorIcon })
            {
                candidate.Visibility = candidate == icon ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        private int ElapsedSeconds() => (int)Math.Round(_elapsed.Elapsed.TotalSeconds);
    }
}
