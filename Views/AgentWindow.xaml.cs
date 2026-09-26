using System;
using System.Collections.Generic;
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
        private readonly CancellationTokenSource _lifetime = new CancellationTokenSource();
        private readonly Stopwatch _elapsed = new Stopwatch();
        private CancellationTokenSource _run;
        private Task<AgentReport> _job;
        private AgentReport _report;
        /// <summary>处理期间每秒刷新一次「已用时」。</summary>
        private DispatcherTimer _ticker;
        private bool _busy;
        private bool _closed;

        internal AgentWindow(IOneNotePageAccess api, string pageId, string selectionXml, AiConfig config, string model, string effort, IntPtr owner)
        {
            InitializeComponent();
            _api = api; _pageId = pageId; _config = config; _model = model; _effort = effort;
            var page = AgentPageSnapshot.ParsePage(selectionXml);
            _selection = AgentPageSnapshot.SelectedIds(page);
            PageText.Text = (string)page.Attribute("name") ?? "当前页";
            ModelText.Text = model + " · 思考 " + effort;
            SelectionScope.IsEnabled = _selection.Count > 0;
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
                if (!snapshot.Blocks.Any(b => b.Editable)) throw new AiException("目标范围没有可编辑文字，或所在文本框包含尚未启用的混合内容。");
                if (!AgentTools.FontInstalled(_config.Agent.FontFamily))
                    _config.Agent.FontFamily = ParagraphStyles.Fonts.FirstOrDefault(AgentTools.FontInstalled)
                        ?? throw new AiException("Agent 默认字体未安装，请在 AI 配置的 Agent/FontFamily 中选择已安装字体。");
                var runner = new AgentRunner(new AgentChatClient(_config, _model, _effort), new AgentCommitter(_api));
                return runner.RunAsync(snapshot, request, CreateProgress(), token);
            }, false);
        }

        private IProgress<string> CreateProgress()
        {
            // 在后台调用此工厂也只通过 Dispatcher 更新 UI，不捕获线程池上下文。
            var last = Stopwatch.StartNew();
            var run = _run;
            return new DirectProgress(text =>
            {
                if (_closed || (text.StartsWith("模型处理中", StringComparison.Ordinal) && last.ElapsedMilliseconds < 200)) return;
                last.Restart();
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (!_closed && _busy && ReferenceEquals(_run, run)) StatusText.Text = text;
                }));
            });
        }

        private sealed class DirectProgress : IProgress<string>
        {
            private readonly Action<string> _report;
            internal DirectProgress(Action<string> report) { _report = report; }
            public void Report(string value) => _report(value);
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
                if (!_closed) ShowOutcome(InfoIcon, "已取消", "未提交本次草稿。");
            }
            catch (Exception ex)
            {
                _report = null;
                if (!_closed) ShowFailure(ex is AiException ? ex.Message : "处理失败。请检查页面状态后重试。");
                // 不记录模型原文、工具参数或笔记内容。
                AddInLog.Info("Agent 失败，类型=" + ex.GetType().Name);
            }
            finally { if (!_closed) SetBusy(false); }
        }

        private async void OnUndo(object sender, RoutedEventArgs e)
        {
            if (_busy || _report == null || _report.Undo.Count == 0) return;
            var previous = _report;
            await StartJob(token => Task.FromResult(new AgentCommitter(_api).Undo(_pageId, previous, _config.Agent, token)), true);
            // 只撤销最近一次执行；不把撤销的逆操作继续暴露为撤销。
            if (_report != null) _report.Undo.Clear();
            if (!_closed) UndoButton.Visibility = Visibility.Collapsed;
        }

        private void OnCancel(object sender, RoutedEventArgs e)
        {
            _run?.Cancel();
            // 停掉「已用时」刷新，免得盖掉下面这行说明；总用时照常计。
            _ticker?.Stop();
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
            UndoButton.Visibility = !value && _report?.Undo.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
            RequestText.IsEnabled = !value;
            PageScope.IsEnabled = !value;
            SelectionScope.IsEnabled = !value && _selection.Count > 0;
        }

        /// <summary>开始转圈和计时，收起上一次的结果。</summary>
        private void StartRunning(string status)
        {
            ShowIcon(SpinnerIcon);
            StatusText.Text = status;
            ResultPanel.Visibility = Visibility.Collapsed;
            ErrorBox.Visibility = Visibility.Collapsed;
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

        /// <summary>停掉转圈，换成结果图标和标题；结果框和错误框先收起，由调用方按需打开。</summary>
        private void ShowOutcome(FrameworkElement icon, string status, string detail)
        {
            StopRunning();
            ShowIcon(icon);
            StatusText.Text = status;
            ShowDetail(detail);
            ResultPanel.Visibility = Visibility.Collapsed;
            ErrorBox.Visibility = Visibility.Collapsed;
        }

        private void ShowElapsed() => ShowDetail($"已用时 {(int)_elapsed.Elapsed.TotalSeconds} 秒");

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
