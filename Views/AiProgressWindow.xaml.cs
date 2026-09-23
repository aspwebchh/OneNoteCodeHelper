using System;
using System.Diagnostics;
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
    /// 「AI 优化」的进度小窗：显示进度、可以取消。顺利改完直接关掉，改动在页面上就看得到；
    /// 失败或结果里有要留意的话时留着，用户看完点标题栏的叉关。
    ///
    /// 和插入窗口一样在 COM 代理进程的独立 STA 线程上 ShowDialog。真正的活（读页面、调接口、写回）
    /// 在线程池上跑，这个窗口只管显示。
    /// </summary>
    public partial class AiProgressWindow : Window
    {
        private readonly Func<IProgress<AiProgress>, CancellationToken, Task<EditResult>> _job;

        private readonly CancellationTokenSource _cancellation = new CancellationTokenSource();

        private Task<EditResult> _jobTask;

        /// <summary>等 AI 期间每秒刷新一次「已用时」。</summary>
        private DispatcherTimer _ticker;

        private readonly Stopwatch _elapsed = new Stopwatch();

        /// <summary>最近一次进度里的实时情况，为 null 时不显示那一行。</summary>
        private string _detail;

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
                _ticker?.Stop();
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
            _elapsed.Start();
            _ticker = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
            _ticker.Tick += (_, __) => ShowDetail();
            _ticker.Start();
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

            // 取消了，或者顺利改完、没有要交代的：直接关窗。
            if ((result == null && failure == null) || (result != null && result.Success && !result.NeedsAttention))
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
            _detail = progress.Detail;
            ShowDetail();

            // 定量的进度条只在分了好几批、而且已经做完一些时才有意义。只有一批时是 0/1 一步跳到 1/1，
            // 等 AI 的整段时间里进度条都是空的、一动不动，看着就像卡死了。
            if (progress.Total > 1 && progress.Done > 0)
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

        private void ShowDetail()
        {
            if (_detail == null)
            {
                DetailText.Visibility = Visibility.Collapsed;
                return;
            }

            DetailText.Text = $"{_detail} · 已用时 {(int)_elapsed.Elapsed.TotalSeconds} 秒";
            DetailText.Visibility = Visibility.Visible;
        }

        /// <summary>有结果了（或正在取消），不再显示实时情况。</summary>
        private void StopDetail()
        {
            _ticker?.Stop();
            _detail = null;
            ShowDetail();
        }

        private void ShowSuccess(string message)
        {
            StopDetail();
            _finished = true;
            StatusText.Text = message;
            ProgressMeter.IsIndeterminate = false;
            ProgressMeter.Maximum = 1;
            ProgressMeter.Value = 1;
            CancelButton.Visibility = Visibility.Collapsed;
        }

        private void ShowFailure(string message)
        {
            StopDetail();
            _finished = true;
            StatusText.Text = message;
            StatusText.Foreground = (System.Windows.Media.Brush)FindResource("Danger");
            ProgressMeter.IsIndeterminate = false;
            ProgressMeter.Value = 0;
            ProgressMeter.Visibility = Visibility.Collapsed;
            CancelButton.Visibility = Visibility.Collapsed;
        }

        private void OnCancel(object sender, RoutedEventArgs e)
        {
            // 先不关窗：要是已经在写回，取消不了，等它做完把结果显示出来，免得用户以为页面没动。
            _cancellation.Cancel();
            StopDetail();
            StatusText.Text = "正在取消…";
            CancelButton.IsEnabled = false;
        }
    }
}
