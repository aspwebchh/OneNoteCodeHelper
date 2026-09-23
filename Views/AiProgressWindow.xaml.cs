using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;
using OneNoteCodeHelper.Interop;
using OneNoteCodeHelper.Services;

namespace OneNoteCodeHelper.Views
{
    /// <summary>
    /// 「AI 优化」的进度小窗：显示进度、可以取消；做完显示结果，没有需要留意的就自己关掉。
    ///
    /// 和插入窗口一样在 COM 代理进程的独立 STA 线程上 ShowDialog。真正的活（读页面、调接口、写回）
    /// 在线程池上跑，这个窗口只管显示。
    /// </summary>
    public partial class AiProgressWindow : Window
    {
        /// <summary>顺利做完、没有额外提示时，结果留这么久再自动关窗。</summary>
        private static readonly TimeSpan AutoCloseDelay = TimeSpan.FromSeconds(2.5);

        private readonly Func<IProgress<AiProgress>, CancellationToken, Task<EditResult>> _job;

        private readonly CancellationTokenSource _cancellation = new CancellationTokenSource();

        private Task<EditResult> _jobTask;

        private DispatcherTimer _autoCloseTimer;

        /// <summary>任务已经有了结果（成功、失败或取消），关窗不再算取消。</summary>
        private bool _finished;

        private bool _closed;

        internal AiProgressWindow(string caption,
            Func<IProgress<AiProgress>, CancellationToken, Task<EditResult>> job, IntPtr ownerHandle)
        {
            InitializeComponent();

            _job = job;
            CaptionText.Text = caption;

            // 属主和居中的做法同插入窗口，见 NativeMethods.CenterOver。
            if (ownerHandle != IntPtr.Zero)
            {
                var interop = new WindowInteropHelper(this) { Owner = ownerHandle };
                SourceInitialized += (_, __) => NativeMethods.CenterOver(interop.Handle, ownerHandle);
            }

            Loaded += (_, __) => Start();

            // 还没做完就关窗等于取消。关窗不能拦着用户，后台任务的收尾由调用方 WaitForJob 等。
            Closing += (_, __) =>
            {
                if (!_finished)
                {
                    _cancellation.Cancel();
                }
            };

            Closed += (_, __) =>
            {
                _closed = true;
                _autoCloseTimer?.Stop();
            };
        }

        /// <summary>
        /// 等后台任务结束。窗口可能在任务还没跑完时就被关掉了，调用方要等它收尾才能放行下一次「AI 优化」，
        /// 否则两次写回可能撞在同一页上。任务的异常已经在 <see cref="Start"/> 里处理过，这里只管等。
        /// </summary>
        internal void WaitForJob(TimeSpan timeout)
        {
            try
            {
                _jobTask?.Wait(timeout);
            }
            catch (Exception)
            {
                // 同上，已经处理过。
            }
        }

        private async void Start()
        {
            // Progress 记下的是当前（界面线程）的同步上下文，后台报的进度会自动回到界面线程。
            var progress = new Progress<AiProgress>(ShowProgress);
            _jobTask = Task.Run(() => _job(progress, _cancellation.Token));

            EditResult result = null;
            string failure = null;

            try
            {
                result = await _jobTask;
            }
            catch (OperationCanceledException)
            {
                AddInLog.Info("用户取消了 AI 优化，页面没有改动。");
            }
            catch (AiException ex)
            {
                AddInLog.Warn("AI 优化失败：" + ex.Message, ex.InnerException);
                failure = ex.Message;
            }
            catch (Exception ex)
            {
                AddInLog.Error("AI 优化失败。", ex);
                failure = $"AI 优化失败：{ex.Message}\n\n详情见日志：{AddInLog.LogPath}";
            }

            // 用户可能早就把窗口关了，这时只留日志，不能再碰窗口。
            if (_closed)
            {
                return;
            }

            if (result == null && failure == null)
            {
                _finished = true;
                Close();
            }
            else if (result != null && result.Success)
            {
                ShowSuccess(result.Message);
            }
            else
            {
                ShowFailure(failure ?? result.Message);
            }
        }

        private void ShowProgress(AiProgress progress)
        {
            if (_finished || _cancellation.IsCancellationRequested)
            {
                return;
            }

            StatusText.Text = progress.Message;

            if (progress.Total > 0)
            {
                ProgressMeter.IsIndeterminate = false;
                ProgressMeter.Maximum = progress.Total;
                ProgressMeter.Value = progress.Done;
            }
            else
            {
                ProgressMeter.IsIndeterminate = true;
            }
        }

        private void ShowSuccess(string message)
        {
            _finished = true;
            StatusText.Text = message;
            ProgressMeter.IsIndeterminate = false;
            ProgressMeter.Maximum = 1;
            ProgressMeter.Value = 1;
            ActionButton.Content = "关闭";
            ActionButton.IsEnabled = true;

            // 结果里带了「跳过」「没采用」这类要留意的话就不自动关，让用户看清楚。
            if (message != null && message.IndexOf('\n') < 0)
            {
                _autoCloseTimer = new DispatcherTimer { Interval = AutoCloseDelay };
                _autoCloseTimer.Tick += (_, __) =>
                {
                    _autoCloseTimer.Stop();
                    Close();
                };
                _autoCloseTimer.Start();
            }
        }

        private void ShowFailure(string message)
        {
            _finished = true;
            StatusText.Text = message;
            StatusText.Foreground = (System.Windows.Media.Brush)FindResource("Danger");
            ProgressMeter.IsIndeterminate = false;
            ProgressMeter.Value = 0;
            ProgressMeter.Visibility = Visibility.Collapsed;
            ActionButton.Content = "关闭";
            ActionButton.IsEnabled = true;
        }

        private void OnAction(object sender, RoutedEventArgs e)
        {
            if (_finished)
            {
                Close();
                return;
            }

            // 先不关窗：要是已经在写回，取消不了，等它做完把结果显示出来，免得用户以为页面没动。
            _cancellation.Cancel();
            StatusText.Text = "正在取消…";
            ActionButton.IsEnabled = false;
        }
    }
}
