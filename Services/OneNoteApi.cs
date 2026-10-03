using System;
using System.Linq;
using System.Xml.Linq;
using System.Threading;
using OneNoteCodeHelper.Services.Agent;
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
    internal sealed class OneNoteApi : Mcp.IOneNoteWorkspaceAccess
    {
        /// <summary>OneNote 页面 XML 的命名空间（当前桌面版返回的就是 2013 架构）。</summary>
        internal const string OneNs = "http://schemas.microsoft.com/office/onenote/2013/onenote";

        internal static readonly XNamespace One = XNamespace.Get(OneNs);

        /// <summary>生命周期归 AddIn 管：断开时由它连同宿主对象一起释放。</summary>
        private readonly IApplication _app;
        private readonly object _calls = new object();
        private volatile bool _stopping;

        private T Call<T>(Func<T> action)
        {
            lock (_calls)
            {
                if (_stopping) throw new AiException("OneNote 正在断开连接，操作已停止。");
                return action();
            }
        }

        internal void Stop() { _stopping = true; }

        internal void Disconnect(Action release)
        {
            _stopping = true;
            if (Monitor.TryEnter(_calls))
            {
                try { release(); } finally { Monitor.Exit(_calls); }
            }
            else ThreadPool.QueueUserWorkItem(_ => { lock (_calls) release(); });
        }

        string IOneNotePageAccess.GetPageContent(string pageId, PageInfo info) => GetPageContent(pageId, info);
        void IOneNotePageAccess.UpdatePageContent(string xml, DateTime expectedLastModified) => UpdatePageContent(xml, expectedLastModified);
        void IOneNotePageAccess.DeletePageContent(string pageId, string objectId, DateTime expectedLastModified) =>
            Call(() => { _app.DeletePageContent(pageId, objectId, expectedLastModified, false); return true; });

        internal OneNoteApi(IApplication application)
        {
            _app = application ?? throw new ArgumentNullException(nameof(application));
        }

        internal string GetHierarchy(string startNodeId, HierarchyScope scope)
        {
            return Call(() => { _app.GetHierarchy(startNodeId, scope, out string xml, XMLSchema.xs2013); return xml; });
        }

        string Mcp.IOneNoteWorkspaceAccess.GetHierarchy(string id, HierarchyScope scope) => GetHierarchy(id, scope);
        string Mcp.IOneNoteWorkspaceAccess.FindPages(string id, string query) => Call(() =>
        { _app.FindPages(id ?? "", query, out string xml, true, false, XMLSchema.xs2013); return xml; });
        string Mcp.IOneNoteWorkspaceAccess.CreatePage(string id) => Call(() =>
        { _app.CreateNewPage(id, out string page, NewPageStyle.npsBlankPageWithTitle); return page; });
        string Mcp.IOneNoteWorkspaceAccess.CreateSection(string id, string name) => Call(() =>
        { _app.OpenHierarchy(name + ".one", id, out string section, CreateFileType.cftSection); return section; });
        void Mcp.IOneNoteWorkspaceAccess.Recycle(string id, DateTime modified) => Call(() =>
        {
            var tree = Mcp.McpReadService.Xml(GetHierarchy("", HierarchyScope.hsPages));
            var node = tree.DescendantsAndSelf().FirstOrDefault(e => (string)e.Attribute("ID") == id && Mcp.McpReadService.IsLive(e));
            var notebook = node?.AncestorsAndSelf(One + "Notebook").FirstOrDefault();
            if (notebook == null || !notebook.Descendants(One + "SectionGroup").Any(e => (string)e.Attribute("isRecycleBin") == "true"))
                throw new AiException("没有可确认的笔记本回收站，删除已停止。");
            if (node.Name == One + "Page")
            {
                int Level(XElement page) => int.TryParse((string)page.Attribute("pageLevel"), out var level) ? level : 1;
                var next = node.ElementsAfterSelf(One + "Page").FirstOrDefault();
                if (next != null && Level(next) > Level(node)) throw new AiException("页面含子页面，回收站操作已停止。");
            }
            _app.DeleteHierarchy(id, modified, false); return true;
        });
        string Mcp.IOneNoteWorkspaceAccess.GetLink(string id, string objectId) => Call(() =>
        { _app.GetHyperlinkToObject(id, objectId ?? "", out string link); return link; });
        void Mcp.IOneNoteWorkspaceAccess.Navigate(string id, string objectId) => NavigateTo(id, objectId ?? "");
        void Mcp.IOneNoteWorkspaceAccess.ExportPdf(string id, string path) => Call(() =>
        { _app.Publish(id, path, PublishFormat.pfPDF, ""); return true; });

        internal string GetPageContent(string pageId, PageInfo info)
        {
            return Call(() => { _app.GetPageContent(pageId, out string xml, info, XMLSchema.xs2013); return xml; });
        }

        internal void UpdatePageContent(string pageChangesXml, DateTime expectedLastModified)
        {
            Call(() => { _app.UpdatePageContent(pageChangesXml, expectedLastModified, XMLSchema.xs2013, false); return true; });
        }

        internal void NavigateTo(string hierarchyObjectId, string objectId)
        {
            Call(() => { _app.NavigateTo(hierarchyObjectId, objectId, false); return true; });
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
                var handle = Call(() => _app.Windows.CurrentWindow?.WindowHandle ?? 0);
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
