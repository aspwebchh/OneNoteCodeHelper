using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using Microsoft.Office.Interop.OneNote;
using OneNoteCodeHelper.Highlighting;
using OneNoteCodeHelper.Interop;
using OneNoteCodeHelper.Services;
using OneNoteCodeHelper.Views;

namespace OneNoteCodeHelper
{
    /// <summary>
    /// 外接程序入口。OneNote 经系统 DLL 代理进程加载这个类并调用 IDTExtensibility2。
    ///
    /// 这里所有 public 方法都可能被 OneNote 直接调用，任何一个漏出异常都可能让 OneNote
    /// 把本插件加进「已禁用项目」从此不再加载。所以每个入口都必须自己兜住异常并写日志。
    /// </summary>
    [ComVisible(true)]
    [Guid(ClassId)]
    [ProgId(ProgIdValue)]
    // 必须是 AutoDispatch，不能用 None。功能区是通过 IDispatch 按方法名回调我们的
    // （OnHighlightSelection 等），ClassInterfaceType.None 不生成自动类接口，
    // GetIDsOfNames 找不到这些方法，结果就是每个按钮点了都没反应。
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class AddIn : Extensibility.IDTExtensibility2, IRibbonExtensibility
    {
        internal const string ClassId = "441360A0-59D3-4969-9F91-7AF166C512BE";

        internal const string ProgIdValue = "OneNoteCodeHelper.AddIn";

        private const string AutoDetectLabel = "自动识别";

        private OneNoteApi _api;
        private PageEditor _editor;
        private AddInSettings _settings;
        private IRibbonUI _ribbon;

        // OnConnection 传进来的两个宿主对象。它们在本代理进程里是跨进程 RCW，
        // 断开时必须主动还回去，见 ReleaseHostReferences。
        private object _hostApplication;
        private object _hostAddIn;

        static AddIn()
        {
            // 静态构造里抛异常会让这个类型永远无法实例化（TypeInitializationException），
            // 宿主那边只会看到一个含糊的激活失败，所以这里必须整个兜住。
            try
            {
                // 本程序集由代理进程中的 mscoree 从 CodeBase 加载。依赖（如 PIA
                // 不在 GAC 时）仍需从插件目录解析。
                AppDomain.CurrentDomain.AssemblyResolve += ResolveFromAddInFolder;

                // 这条日志是排查「插件到底有没有被加载」的第一个落点：
                // 有它没有后续 = 类型加载成功但宿主没接着调 OnConnection；
                // 连它都没有 = 程序集/类型根本没被加载起来。
                AddInLog.Info("AddIn 类型已初始化，宿主进程：" + Process.GetCurrentProcess().ProcessName);
            }
            catch
            {
                // 连日志都写不了也只能继续，不能让宿主拿不到对象。
            }
        }

        public AddIn()
        {
            AddInLog.Info("AddIn 实例已创建，等待宿主调用 OnConnection。");
        }

        private static Assembly ResolveFromAddInFolder(object sender, ResolveEventArgs args)
        {
            try
            {
                var folder = Path.GetDirectoryName(new Uri(typeof(AddIn).Assembly.CodeBase).LocalPath);
                if (string.IsNullOrEmpty(folder))
                {
                    return null;
                }

                var path = Path.Combine(folder, new AssemblyName(args.Name).Name + ".dll");
                return File.Exists(path) ? Assembly.LoadFrom(path) : null;
            }
            catch (Exception ex)
            {
                AddInLog.Warn("解析依赖程序集失败：" + args.Name, ex);
                return null;
            }
        }

        #region IDTExtensibility2

        public void OnConnection(object application, Extensibility.ext_ConnectMode connectMode,
            object addInInst, ref Array custom)
        {
            try
            {
                AddInLog.Info($"OnConnection 开始，connectMode={connectMode}");

                _hostApplication = application;
                _hostAddIn = addInInst;
                _settings = SettingsStore.Load();

                // 宿主传来的是原始 COM IDispatch。用 PIA 为这个已有对象创建
                // 包装器；不能再 new Application()，否则会额外激活 OneNote，
                // 并在退出时留下跨进程 COM 引用。
                // CreateWrapperOfType 把包装器挂在 application 这个 RCW 上，
                // 释放 application 时会一并释放，不用单独管它的生命周期。
                var oneNote = application as IApplication;
                if (oneNote == null)
                {
                    if (application == null || !Marshal.IsComObject(application))
                    {
                        throw new InvalidCastException("宿主没有传入 OneNote COM 应用对象。");
                    }

                    oneNote = (IApplication)Marshal.CreateWrapperOfType(
                        application, typeof(ApplicationClass));
                }

                _api = new OneNoteApi(oneNote);
                _editor = new PageEditor(_api);
                AddInLog.Info("已连接到 OneNote。");
            }
            catch (Exception ex)
            {
                AddInLog.Error("OnConnection 失败。", ex);
            }
        }

        public void OnDisconnection(Extensibility.ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                AddInLog.Info($"OnDisconnection，removeMode={removeMode}");
                _editor = null;
                _api = null;
                ReleaseHostReferences();
                AddInLog.Info("已释放全部 OneNote COM 引用。");
            }
            catch (Exception ex)
            {
                AddInLog.Error("OnDisconnection 失败。", ex);
            }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }

        #endregion

        #region IRibbonExtensibility

        public string GetCustomUI(string ribbonId)
        {
            try
            {
                var assembly = typeof(AddIn).Assembly;
                var name = assembly.GetManifestResourceNames()
                    .FirstOrDefault(n => n.EndsWith("Ribbon.xml", StringComparison.OrdinalIgnoreCase));

                if (name == null)
                {
                    AddInLog.Error("嵌入资源里找不到 Ribbon.xml。");
                    return string.Empty;
                }

                using (var stream = assembly.GetManifestResourceStream(name))
                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                AddInLog.Error("读取功能区定义失败。", ex);
                return string.Empty;
            }
        }

        #endregion

        #region 功能区回调

        // 这些回调的第一个参数是 IRibbonControl。这里一律声明成 object：我们用不同的方法名
        // 区分控件，不需要读它的 Id，声明成 object 就少一次接口 QI，少一个出错的可能。

        public void OnRibbonLoad(object ribbonUi)
        {
            try
            {
                _ribbon = ribbonUi as IRibbonUI;
                AddInLog.Info("功能区已加载。");
            }
            catch (Exception ex)
            {
                AddInLog.Warn("保存功能区引用失败，主题切换后按钮状态可能不刷新。", ex);
            }
        }

        public void OnHighlightSelection(object control)
        {
            Guard("高亮选中", () =>
            {
                var result = _editor.HighlightSelection(_settings, _settings.LanguageId);
                Report(result);
            });
        }

        public void OnShowInsertWindow(object control)
        {
            Guard("插入代码", () =>
            {
                var editor = _editor;
                var settings = _settings;
                var ownerHandle = _api.GetMainWindowHandle();

                // OneNote 在 DLL 代理进程中以 MTA 调用功能区回调。WPF 窗口
                // 必须在 STA 线程创建；回调立即返回，避免阻塞 OneNote 的 COM 调用。
                var thread = new Thread(() =>
                {
                    try
                    {
                        var window = new InsertCodeWindow(editor, settings, ownerHandle);
                        window.ShowDialog();

                        if (window.SettingsChanged)
                        {
                            _settings = window.Settings;
                            SettingsStore.Save(_settings);
                            InvalidateRibbon();
                        }
                    }
                    catch (Exception ex)
                    {
                        AddInLog.Error("「插入代码」窗口失败。", ex);
                        ShowMessage($"「插入代码」失败：{ex.Message}\n\n详情见日志：\n{AddInLog.LogPath}",
                            MessageBoxImage.Error);
                    }
                });
                thread.SetApartmentState(ApartmentState.STA);
                thread.IsBackground = true;
                thread.Start();
            });
        }

        public int GetLanguageCount(object control)
        {
            // 第 0 项是「自动识别」，后面才是具体语言。
            return LanguageRegistry.All.Count + 1;
        }

        public string GetLanguageId(object control, int index)
        {
            return index == 0 ? LanguageRegistry.AutoDetectId : LanguageRegistry.All[index - 1].Id;
        }

        public string GetLanguageLabel(object control, int index)
        {
            return index == 0 ? AutoDetectLabel : LanguageRegistry.All[index - 1].DisplayName;
        }

        public int GetSelectedLanguageIndex(object control)
        {
            var settings = _settings ?? (_settings = SettingsStore.Load());

            for (var i = 0; i < LanguageRegistry.All.Count; i++)
            {
                if (string.Equals(LanguageRegistry.All[i].Id, settings.LanguageId,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return i + 1;
                }
            }

            return 0;
        }

        public void OnLanguageChanged(object control, string selectedId, int selectedIndex)
        {
            Guard("切换语言", () =>
            {
                _settings.LanguageId = selectedId;
                SettingsStore.Save(_settings);
                AddInLog.Info("语言切换为 " + selectedId);
            });
        }

        public bool GetDarkThemePressed(object control)
        {
            var settings = _settings ?? (_settings = SettingsStore.Load());
            return string.Equals(settings.ThemeId, CodeThemesDarkId, StringComparison.OrdinalIgnoreCase);
        }

        public void OnDarkThemeToggled(object control, bool pressed)
        {
            Guard("切换主题", () =>
            {
                _settings.ThemeId = pressed ? CodeThemesDarkId : CodeThemesLightId;
                SettingsStore.Save(_settings);
                AddInLog.Info("主题切换为 " + _settings.ThemeId);
            });
        }

        public void OnOpenLog(object control)
        {
            Guard("打开日志", () =>
            {
                if (!File.Exists(AddInLog.LogPath))
                {
                    AddInLog.Info("用户点了诊断日志。");
                }

                Process.Start(new ProcessStartInfo(AddInLog.LogPath) { UseShellExecute = true });
            });
        }

        #endregion

        private const string CodeThemesLightId = "light";

        private const string CodeThemesDarkId = "dark";

        /// <summary>
        /// 统一的回调外壳：保证前置条件齐备、异常不外泄、出错有提示也有日志。
        /// </summary>
        private void Guard(string action, Action body)
        {
            try
            {
                if (_editor == null || _api == null)
                {
                    ShowMessage("插件还没连上 OneNote，请重启 OneNote 后再试。", MessageBoxImage.Warning);
                    AddInLog.Error($"「{action}」时发现尚未连接 OneNote。");
                    return;
                }

                if (_settings == null)
                {
                    _settings = SettingsStore.Load();
                }

                body();
            }
            catch (Exception ex)
            {
                AddInLog.Error($"「{action}」失败。", ex);
                ShowMessage($"「{action}」失败：{ex.Message}\n\n详情见日志：\n{AddInLog.LogPath}",
                    MessageBoxImage.Error);
            }
        }

        private static void Report(EditResult result)
        {
            if (result.Success)
            {
                AddInLog.Info(result.Message);
                return;
            }

            ShowMessage(result.Message, MessageBoxImage.Warning);
        }

        private static void ShowMessage(string message, MessageBoxImage icon)
        {
            MessageBox.Show(message, "OneNote 代码高亮", MessageBoxButton.OK, icon);
        }

        /// <summary>
        /// 把本代理进程手里所有指向 OneNote 的 COM 引用还回去。
        ///
        /// 插件跑在 dllhost 里，拿到的 OneNote 对象全是跨进程代理。OneNote 关闭后 dllhost
        /// 会被直接结束，CLR 不会替还活着的 RCW 调 Release，OneNote 只能等 DCOM 的 ping
        /// 超时（约 6 分钟）才回收这些引用。这段时间 ONENOTE.EXE 没有窗口却一直留在后台，
        /// 期间再打开 OneNote 就会弹「正在清理上次打开之后的内容」。本机实测：不释放时
        /// OnDisconnection 之后正好 6 分钟 ONENOTE.EXE 才退出。
        /// </summary>
        private void ReleaseHostReferences()
        {
            FinalRelease(ref _ribbon);
            FinalRelease(ref _hostAddIn);
            FinalRelease(ref _hostApplication);

            // 功能区回调的 control 参数、Windows 集合之类的临时 RCW 没有字段指着，
            // 只能靠 GC。代理进程马上就要退出，不会再有机会跑终结器，所以在这里同步收掉。
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private static void FinalRelease<T>(ref T reference) where T : class
        {
            var target = reference;
            reference = null;

            if (target == null || !Marshal.IsComObject(target))
            {
                return;
            }

            try
            {
                Marshal.FinalReleaseComObject(target);
            }
            catch (Exception ex)
            {
                AddInLog.Warn("释放 OneNote COM 引用失败。", ex);
            }
        }

        private void InvalidateRibbon()
        {
            try
            {
                _ribbon?.Invalidate();
            }
            catch (Exception ex)
            {
                AddInLog.Warn("刷新功能区失败。", ex);
            }
        }
    }
}
