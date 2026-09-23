using System;
using System.Linq;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;

namespace OneNoteCodeHelper.Services
{
    /// <summary>
    /// OneNote COM 自动化接口的薄封装。
    ///
    /// 这里必须走早绑定。本机实测：OneNote 的类型库没有注册到 CLR 能找到的位置，导致
    /// Type.InvokeMember 抛 TYPE_E_LIBNOTREGISTERED、dynamic 抛 E_FAIL（同一 STA 线程上
    /// PowerShell 的原生晚绑定却是好的，说明 COM 对象本身没问题）。早绑定直接按接口 IID
    /// 做 QI，不碰类型库，因此可用。
    /// </summary>
    internal sealed class OneNoteApi
    {
        /// <summary>OneNote 页面 XML 的命名空间（当前桌面版返回的就是 2013 架构）。</summary>
        internal const string OneNs = "http://schemas.microsoft.com/office/onenote/2013/onenote";

        internal static readonly XNamespace One = XNamespace.Get(OneNs);

        private readonly IApplication _app;

        internal OneNoteApi(IApplication application)
        {
            _app = application ?? throw new ArgumentNullException(nameof(application));
        }

        /// <summary>脱离外接程序宿主时自行激活一个 OneNote 实例，供命令行自测使用。</summary>
        internal static OneNoteApi CreateStandalone() => new OneNoteApi(new Application());

        internal string GetHierarchy(string startNodeId, HierarchyScope scope)
        {
            _app.GetHierarchy(startNodeId, scope, out string xml);
            return xml;
        }

        internal string GetPageContent(string pageId, PageInfo info)
        {
            _app.GetPageContent(pageId, out string xml, info);
            return xml;
        }

        internal void UpdatePageContent(string pageChangesXml, DateTime expectedLastModified)
        {
            _app.UpdatePageContent(pageChangesXml, expectedLastModified);
        }

        internal void NavigateTo(string hierarchyObjectId, string objectId)
        {
            _app.NavigateTo(hierarchyObjectId, objectId, false);
        }

        /// <summary>
        /// 取当前正在查看的页面 ID。
        ///
        /// 注意：新版 OneNote（16.0.20326 实测）的 Windows.CurrentWindow.CurrentPageId 返回空字符串，
        /// 拿它去调 GetPageContent 会抛 0x80042005。所以只能从层级树里找 isCurrentlyViewed="true"。
        /// 先定位当前分区、再只查该分区的页面，避免在几千页的笔记本上整树扫描
        /// （实测整树 576KB、当前分区只有 18KB）。
        /// </summary>
        internal string GetCurrentPageId()
        {
            var sectionId = FindCurrentId(GetHierarchy(null, HierarchyScope.hsSections), "Section");
            if (!string.IsNullOrEmpty(sectionId))
            {
                var pageId = FindCurrentId(GetHierarchy(sectionId, HierarchyScope.hsPages), "Page");
                if (!string.IsNullOrEmpty(pageId))
                {
                    return pageId;
                }
            }

            // 兜底：整棵树扫一遍。
            return FindCurrentId(GetHierarchy(null, HierarchyScope.hsPages), "Page");
        }

        private static string FindCurrentId(string hierarchyXml, string elementName)
        {
            if (string.IsNullOrEmpty(hierarchyXml))
            {
                return null;
            }

            return XDocument.Parse(hierarchyXml)
                .Descendants(One + elementName)
                .Where(e => (string)e.Attribute("isCurrentlyViewed") == "true")
                .Select(e => (string)e.Attribute("ID"))
                .FirstOrDefault();
        }

        /// <summary>取 OneNote 主窗口句柄，用于把 WPF 对话框设为其子窗口。</summary>
        internal IntPtr GetMainWindowHandle()
        {
            try
            {
                var handle = _app.Windows.CurrentWindow?.WindowHandle ?? 0;
                return new IntPtr((long)handle);
            }
            catch (Exception ex)
            {
                AddInLog.Warn("取 OneNote 窗口句柄失败，对话框将不设置属主。", ex);
                return IntPtr.Zero;
            }
        }
    }
}
