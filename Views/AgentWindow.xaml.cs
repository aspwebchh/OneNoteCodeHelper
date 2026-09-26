using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
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
        private CancellationTokenSource _run;
        private Task<AgentReport> _job;
        private AgentReport _report;
        private bool _busy;
        private bool _closed;

        internal AgentWindow(IOneNotePageAccess api, string pageId, string selectionXml, AiConfig config, string model, string effort, IntPtr owner)
        {
            InitializeComponent();
            _api = api; _pageId = pageId; _config = config; _model = model; _effort = effort;
            var page = AgentPageSnapshot.ParsePage(selectionXml);
            _selection = AgentPageSnapshot.SelectedIds(page);
            PageText.Text = "目标页面：" + ((string)page.Attribute("name") ?? "当前页");
            ModelText.Text = model + " · 思考 " + effort;
            SelectionScope.IsEnabled = _selection.Count > 0;
            if (owner != IntPtr.Zero) new WindowInteropHelper(this).Owner = owner;
            Closing += (_, __) => _lifetime.Cancel();
            Closed += (_, __) => _closed = true;
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
            if (request.Length == 0) { StatusText.Text = "请先输入需求。"; return; }
            if (string.IsNullOrWhiteSpace(_config.ApiKey)) { StatusText.Text = "请先在「AI 配置」中填写 ApiKey。"; return; }
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
            });
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

        private async Task StartJob(Func<CancellationToken, Task<AgentReport>> job)
        {
            _run?.Dispose();
            _run = CancellationTokenSource.CreateLinkedTokenSource(_lifetime.Token);
            SetBusy(true);
            ResultText.Text = "";
            StatusText.Text = "正在读取固定目标页…";
            var timer = Stopwatch.StartNew();
            try
            {
                _job = Task.Run(() => job(_run.Token));
                _report = await _job;
                if (_closed) return;
                ResultText.Text = _report.Message;
                StatusText.Text = $"处理结束 · {timer.Elapsed.TotalSeconds:0} 秒";
            }
            catch (OperationCanceledException)
            {
                _report = null;
                if (!_closed) { StatusText.Text = "已取消"; ResultText.Text = "未提交本次草稿。"; }
            }
            catch (Exception ex)
            {
                _report = null;
                if (!_closed)
                {
                    StatusText.Text = "未能完成";
                    ResultText.Text = ex is AiException ? ex.Message : "处理失败。请检查页面状态后重试。";
                }
                // 不记录模型原文、工具参数或笔记内容。
                AddInLog.Info("Agent 失败，类型=" + ex.GetType().Name);
            }
            finally { if (!_closed) SetBusy(false); }
        }

        private async void OnUndo(object sender, RoutedEventArgs e)
        {
            if (_busy || _report == null || _report.Undo.Count == 0) return;
            var previous = _report;
            await StartJob(token => Task.FromResult(new AgentCommitter(_api).Undo(_pageId, previous, _config.Agent, token)));
            // 只撤销最近一次执行；不把撤销的逆操作继续暴露为撤销。
            if (_report != null) _report.Undo.Clear();
            if (!_closed) UndoButton.IsEnabled = false;
        }
        private void OnCancel(object sender, RoutedEventArgs e)
        {
            _run?.Cancel();
            StatusText.Text = "正在取消；如果已开始写回，将先核对实际结果…";
            CancelButton.IsEnabled = false;
        }
        private void SetBusy(bool value)
        {
            _busy = value;
            ExecuteButton.IsEnabled = !value;
            CancelButton.IsEnabled = value;
            UndoButton.IsEnabled = !value && _report?.Undo.Count > 0;
            RequestText.IsEnabled = !value;
            PageScope.IsEnabled = !value;
            SelectionScope.IsEnabled = !value && _selection.Count > 0;
        }
    }
}
