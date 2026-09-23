using System;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using OneNoteCodeHelper.Interop;
using OneNoteCodeHelper.Services;

namespace OneNoteCodeHelper.Views
{
    /// <summary>
    /// 「AI 优化」的进度小窗：处理中显示进度、可以取消；做完显示结果（AI 自己写的改动清单），
    /// 不会自己关，用户看完点标题栏的叉或按 Esc 关。
    ///
    /// 和插入窗口一样在 COM 代理进程的独立 STA 线程上 ShowDialog。真正的活（读页面、调接口、写回）
    /// 在线程池上跑，这个窗口只管显示。
    /// </summary>
    public partial class AiProgressWindow : Window
    {
        private readonly Func<IProgress<AiProgress>, CancellationToken, Task<AiReport>> _job;

        private readonly CancellationTokenSource _cancellation = new CancellationTokenSource();

        private Task<AiReport> _jobTask;

        /// <summary>等 AI 期间每秒刷新一次「已用时」。</summary>
        private DispatcherTimer _ticker;

        private readonly Stopwatch _elapsed = new Stopwatch();

        /// <summary>最近一次进度里的实时情况，为 null 时不显示那一行。</summary>
        private string _detail;

        /// <summary>任务已经有了结果（成功、失败或取消），关窗不再算取消。</summary>
        private bool _finished;

        private bool _closed;

        internal AiProgressWindow(string functionName, string modelId, string effort,
            Func<IProgress<AiProgress>, CancellationToken, Task<AiReport>> job, IntPtr ownerHandle)
        {
            InitializeComponent();

            _job = job;
            FunctionText.Text = functionName;
            ModelText.Text = $"{modelId} · 思考 {effort}";
            AppIcon.Source = LoadIcon("AiOptimize");

            // 不设的话标题栏上是宿主进程（dllhost）的图标。
            if (AppIcon.Source != null)
            {
                Icon = AppIcon.Source;
            }

            var interop = new WindowInteropHelper(this);
            if (ownerHandle != IntPtr.Zero)
            {
                interop.Owner = ownerHandle;
            }

            SourceInitialized += (_, __) =>
            {
                // 标题栏刷成窗口底色，和下面连成一整块（Windows 11 才有效果）。
                NativeMethods.TrySetCaptionColor(interop.Handle, ((SolidColorBrush)Background).Color);

                // 属主和居中的做法同插入窗口，见 NativeMethods.CenterOver。
                if (ownerHandle != IntPtr.Zero)
                {
                    NativeMethods.CenterOver(interop.Handle, ownerHandle);
                }
            };

            Loaded += (_, __) => Start();

            // Esc 和点叉一样：做完了就是关窗，没做完就是取消。
            PreviewKeyDown += (_, e) =>
            {
                if (e.Key == Key.Escape)
                {
                    e.Handled = true;
                    Close();
                }
            };

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
                StopRunning();
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
            SpinnerRotation.BeginAnimation(RotateTransform.AngleProperty,
                new DoubleAnimation(0, 360, TimeSpan.FromSeconds(0.9)) { RepeatBehavior = RepeatBehavior.Forever });
            _jobTask = Task.Run(() => _job(progress, _cancellation.Token));

            AiReport report = null;
            string failure = null;

            try
            {
                report = await _jobTask;
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

            _finished = true;
            StopRunning();

            if (report == null && failure == null)
            {
                // 用户点了取消，也真的取消掉了：这是用户自己要的，直接关。
                Close();
            }
            else if (report != null && report.Success)
            {
                ShowReport(report);
            }
            else
            {
                ShowFailure(failure ?? report.Error);
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

            // 比例只在分了好几批、而且已经做完一些时才有意义。只有一批时是 0/1 一步跳到 1/1，
            // 一直空着的进度条看着像卡死了，这时只靠转圈表示还在跑。
            if (progress.Total > 1 && progress.Done > 0)
            {
                ProgressMeter.Maximum = progress.Total;
                ProgressMeter.Value = progress.Done;
                ProgressMeter.Visibility = Visibility.Visible;
            }
            else
            {
                ProgressMeter.Visibility = Visibility.Collapsed;
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

        /// <summary>有结果了（或窗口关了）：停掉计时、转圈，藏起进度条和取消按钮。</summary>
        private void StopRunning()
        {
            _ticker?.Stop();
            _elapsed.Stop();
            _detail = null;
            SpinnerRotation.BeginAnimation(RotateTransform.AngleProperty, null);
            ProgressMeter.Visibility = Visibility.Collapsed;
            CancelButton.Visibility = Visibility.Collapsed;
        }

        private void ShowReport(AiReport report)
        {
            if (report.Changed)
            {
                ShowIcon(SuccessIcon);
                StatusText.Text = "已完成";
            }
            else if (report.Conflicted > 0)
            {
                ShowIcon(WarningIcon);
                StatusText.Text = "没有写回任何改动";
            }
            else
            {
                ShowIcon(InfoIcon);
                StatusText.Text = "没有需要修改的地方";
            }

            ShowSummary($"{report.Scope} · 用时 {ElapsedSeconds()} 秒");

            // 只删了空行、AI 没改字时不显示清单；改了字却没写说明时，清单位置只放一行提示。
            var hasChanges = report.Changes.Count > 0;
            ChangesList.ItemsSource = report.Changes;
            ChangesBox.Visibility = hasChanges ? Visibility.Visible : Visibility.Collapsed;
            NoChangesText.Visibility = hasChanges ? Visibility.Collapsed : Visibility.Visible;
            ChangesPanel.Visibility = hasChanges || report.Applied > 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        private void ShowFailure(string message)
        {
            ShowIcon(ErrorIcon);
            StatusText.Text = "没能完成";
            ShowSummary($"用时 {ElapsedSeconds()} 秒");
            ErrorText.Text = message;
            ErrorBox.Visibility = Visibility.Visible;
        }

        private void ShowSummary(string text)
        {
            DetailText.Text = text;
            DetailText.Visibility = Visibility.Visible;
        }

        private void ShowIcon(FrameworkElement icon)
        {
            foreach (var candidate in new FrameworkElement[] { SpinnerIcon, SuccessIcon, WarningIcon, InfoIcon, ErrorIcon })
            {
                candidate.Visibility = candidate == icon ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        private int ElapsedSeconds()
        {
            return (int)Math.Round(_elapsed.Elapsed.TotalSeconds);
        }

        private void OnCancel(object sender, RoutedEventArgs e)
        {
            // 先不关窗：要是已经在写回，取消不了，等它做完把结果显示出来，免得用户以为页面没动。
            _cancellation.Cancel();
            _detail = null;
            ShowDetail();
            StatusText.Text = "正在取消…";
            CancelButton.IsEnabled = false;
        }

        /// <summary>读嵌入资源里的图标（和功能区用的是同一张），找不到就不显示。</summary>
        private static ImageSource LoadIcon(string name)
        {
            try
            {
                var assembly = typeof(AiProgressWindow).Assembly;
                var resource = assembly.GetManifestResourceNames()
                    .FirstOrDefault(n => n.EndsWith("." + name + ".png", StringComparison.OrdinalIgnoreCase));
                if (resource == null)
                {
                    return null;
                }

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
                AddInLog.Warn("读取进度窗图标失败：" + name, ex);
                return null;
            }
        }
    }
}
