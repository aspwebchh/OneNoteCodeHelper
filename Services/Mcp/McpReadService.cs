using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using OneNoteCodeHelper.Services.Agent;
using OneNoteCodeHelper.Mcp;

namespace OneNoteCodeHelper.Services.Mcp
{
    internal sealed class McpReadService
    {
        private static XNamespace One => OneNoteApi.One;
        private sealed class ReadResult
        {
            internal string Owner, Query, Kind;
            internal DateTime Expires;
            internal object[] Items;
            internal object[] Skipped = new object[0];
            internal McpPageModel Page;
            internal string Format;
        }
        private readonly object _gate = new object();
        private readonly Dictionary<string, ReadResult> _reads = new Dictionary<string, ReadResult>();
        private readonly IOneNotePageAccess _pages;
        private readonly IOneNoteWorkspaceAccess _workspace;
        private readonly Func<string> _current;
        private readonly Func<DateTime> _now;
        internal McpReadService(IOneNotePageAccess api, Func<string> current, Func<DateTime> now, Func<AgentOptions> options)
        { _pages = api; _workspace = api as IOneNoteWorkspaceAccess; _current = current; _now = now; }
        internal IOneNoteWorkspaceAccess Workspace => _workspace ?? throw new McpFault("workspace_unavailable", "当前页面访问器没有工作区能力。");
        internal static XElement Xml(string xml)
        {
            using (var reader = XmlReader.Create(new StringReader(xml), new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 16000000 }))
                return XElement.Load(reader, LoadOptions.PreserveWhitespace);
        }
        internal XElement ReadPage(string id, PageInfo info = PageInfo.piBasic)
        {
            var xml = _pages.GetPageContent(id, info);
            if (xml.Length > 16000000) throw new McpFault("page_limit", "页面 XML 超过独立读取的 16 MB 上限。");
            var page = AgentPageSnapshot.ParsePage(xml);
            if ((string)page.Attribute("ID") != id) throw new McpFault("page_mismatch", "读取页面身份不一致。");
            return page;
        }
        internal object Call(string owner, string name, IDictionary<string, object> args, CancellationToken cancellation, Action<int, string> progress = null)
        {
            string Arg(string field, string fallback = null) => args.TryGetValue(field, out var value) ? (string)value : fallback;
            if (name == "get_link" || name == "navigate_to")
            {
                var id = Arg("page_id"); var objectId = Arg("object_id");
                var tree = Xml(Workspace.GetHierarchy("", HierarchyScope.hsPages));
                if (!tree.DescendantsAndSelf().Any(e => (string)e.Attribute("ID") == id && IsLive(e))) throw new McpFault("node_unavailable", "真实层级 ID 不存在。");
                if (objectId != null && !ReadPage(id).Descendants().Any(e => (string)e.Attribute("objectID") == objectId))
                    throw new McpFault("object_unavailable", "对象不在指定页面中，不能使用草稿短 ID。");
                if (name == "get_link") return new { ok = true, page_id = id, link = Workspace.GetLink(id, objectId) };
                Workspace.Navigate(id, objectId); return new { ok = true };
            }
            if (name == "export_page") return Export(Arg("page_id"), Arg("format"), Arg("output_path"), cancellation);
            var query = McpJson.Serialize(args.Where(a => a.Key != "cursor" && a.Key != "limit").OrderBy(a => a.Key).ToDictionary(a => a.Key, a => a.Value));
            var limit = args.TryGetValue("limit", out var raw) ? Convert.ToInt32(raw) : 50;
            ReadResult result; string key; var offset = 0;
            if (args.TryGetValue("cursor", out var cursor))
            {
                var pieces = ((string)cursor).Split('.');
                if (pieces.Length != 2 || !int.TryParse(pieces[1], NumberStyles.None, CultureInfo.InvariantCulture, out offset) || offset < 0) throw new McpFault("cursor_invalid", "无效的读取游标。");
                key = pieces[0];
                lock (_gate)
                {
                    Prune();
                    if (!_reads.TryGetValue(key, out result)) throw new McpFault("cursor_expired", "只读快照已过期，请重新读取。");
                    if (result.Owner != owner) throw new McpFault("cursor_owner", "游标属于其他客户端。");
                    if (result.Kind != name || result.Query != query || offset > result.Items.Length) throw new McpFault("cursor_invalid", "游标与查询不一致。");
                }
            }
            else
            {
                result = new ReadResult { Owner = owner, Kind = name, Query = query, Expires = _now().AddMinutes(5) };
                var skipped = new List<object>(); var items = new List<object>();
                switch (name)
                {
                    case "read_page": case "read_selection":
                        var id = name == "read_page" ? Arg("page_id") : _current();
                        if (string.IsNullOrEmpty(id)) throw new McpFault("page_unavailable", "没有当前页面。");
                        var page = ReadPage(id, name == "read_selection" ? PageInfo.piSelection : PageInfo.piBasic);
                        var selection = name == "read_selection" ? AgentPageSnapshot.SelectedIds(page) : null;
                        if (selection != null && selection.Count == 0) throw new McpFault("selection_empty", "没有选中段落。");
                        selection?.UnionWith(PageEditor.FindSelectedBlankLines(page));
                        result.Page = new McpPageModel(page, selection); result.Format = Arg("format", "blocks"); items.AddRange(result.Page.Blocks); break;
                    case "list_nodes":
                        var parent = Arg("parent_id");
                        var tree = Xml(Workspace.GetHierarchy(parent ?? "", parent == null ? HierarchyScope.hsNotebooks : HierarchyScope.hsChildren));
                        var start = parent == null ? tree : tree.DescendantsAndSelf().FirstOrDefault(e => (string)e.Attribute("ID") == parent);
                        if (start == null) throw new McpFault("node_unavailable", "父节点不存在。");
                        var paths = Xml(Workspace.GetHierarchy("", HierarchyScope.hsSections));
                        if (!IsLive(start)) throw new McpFault("node_unavailable", "不能浏览回收站节点。");
                        items.AddRange(start.Elements().Where(e => IsNode(e) && IsLive(e)).Select(e => Node(e, paths))); break;
                    case "search_pages":
                        XElement found;
                        try { found = Xml(Workspace.FindPages(Arg("root_id", ""), Arg("query"))); }
                        catch (Exception) { throw new McpFault("search_unavailable", "OneNote 原生搜索不可用，请检查 Windows Search 和分区访问状态。"); }
                        var searchPaths = Xml(Workspace.GetHierarchy("", HierarchyScope.hsSections));
                        // 固定搜索候选及其版本。只读取返回批次；后来变更的页面不混入旧查询。
                        foreach (var e in found.DescendantsAndSelf(One + "Page").Where(IsLive))
                        {
                            cancellation.ThrowIfCancellationRequested(); var node = Node(e, searchPaths);
                            items.Add(node);
                        }
                        break;
                    case "find_tasks":
                        var hierarchy = Xml(Workspace.GetHierarchy(Arg("root_id", ""), HierarchyScope.hsPages)); var scanned = 0;
                        var taskPaths = Xml(Workspace.GetHierarchy("", HierarchyScope.hsSections));
                        foreach (var e in hierarchy.DescendantsAndSelf(One + "Page").Where(IsLive))
                        {
                            cancellation.ThrowIfCancellationRequested(); var pid = (string)e.Attribute("ID");
                            try
                            {
                                var taskPage = ReadPage(pid); var completed = args.TryGetValue("completed", out var value) && (bool)value;
                                foreach (var oe in taskPage.Descendants(One + "OE").Where(o => o.Elements(One + "Tag").Any(t => AgentMarks.KindOf(taskPage, t) == "todo" && AgentMarks.IsCompleted(t) == completed)))
                                    items.Add(new { page_id = pid, page_title = (string)e.Attribute("name"), path = Node(e, taskPaths)["path"], object_id = (string)oe.Attribute("objectID"),
                                        text = PageEditor.ExtractPlainText(oe), completed, link = SafeLink(pid, (string)oe.Attribute("objectID")) });
                            }
                            catch (Exception) { skipped.Add(new { page_id = pid, reason = "page_unreadable" }); }
                            progress?.Invoke(++scanned, "扫描待办页面");
                        }
                        break;
                    default: throw new McpFault("unknown_tool", "未知读取工具。");
                }
                cancellation.ThrowIfCancellationRequested(); result.Items = items.ToArray(); result.Skipped = skipped.ToArray(); key = Guid.NewGuid().ToString("N");
                lock (_gate) { Prune(); if (_reads.Count >= 64) throw new McpFault("read_limit", "只读快照数量达到上限，请稍后重试。"); _reads.Add(key, result); }
            }
            var batch = result.Items.Skip(offset).Take(limit).ToArray();
            if (name == "search_pages") lock (result)
                foreach (var item in batch.Cast<Dictionary<string, object>>())
                {
                    cancellation.ThrowIfCancellationRequested();
                    if (item.ContainsKey("summary")) continue;
                    item["summary"] = null; item["complete"] = false;
                    try
                    {
                        var candidate = ReadPage((string)item["id"]);
                        var expected = (string)item["last_modified"];
                        if (expected == null) item["reason"] = "version_unavailable";
                        else if (expected != (string)candidate.Attribute("lastModifiedTime")) item["reason"] = "page_changed";
                        else
                        {
                            var model = new McpPageModel(candidate);
                            item["summary"] = string.Join(" ", model.Blocks.Select(b => (string)b["text"])).SubstringSafe(240);
                            item["complete"] = model.Complete; item["issues"] = model.Issues.Distinct().ToArray();
                        }
                    }
                    catch (Exception) { item["reason"] = "page_unreadable"; }
                    item["link"] = SafeLink((string)item["id"], null); progress?.Invoke(offset + Array.IndexOf(batch, item) + 1, "读取搜索摘要");
                }
            // 任何页面段落过大时拒绝本批，不截断正文却声称完整读取。
            if (Encoding.UTF8.GetByteCount(McpJson.Serialize(batch)) > McpProtocol.MaxFrameBytes / 2) throw new McpFault("result_limit", "读取结果过大，请减小 limit。");
            return new { ok = true, read_id = key, page_id = result.Page?.PageId, page_title = result.Page?.Title, last_modified = result.Page?.Modified,
                format = result.Format, items = batch, markdown = result.Format == "markdown" ? string.Join("\n", batch.Cast<Dictionary<string, object>>().Select(b => (string)b["markdown"])) : null,
                complete = result.Page?.Complete ?? (result.Skipped.Length == 0 && (name != "search_pages" || batch.Cast<Dictionary<string, object>>().All(i => (bool)i["complete"]))), issues = result.Page?.Issues.Distinct().ToArray(), skipped = result.Skipped,
                total = result.Items.Length, next_cursor = offset + batch.Length < result.Items.Length ? key + "." + (offset + batch.Length) : null, expires_at = result.Expires.ToString("o") };
        }
        private string SafeLink(string id, string objectId) { try { return Workspace.GetLink(id, objectId); } catch (Exception) { return null; } }
        internal static bool IsNode(XElement e) => new[] { "Notebook", "SectionGroup", "Section", "Page" }.Contains(e.Name.LocalName);
        internal static bool IsLive(XElement e) => !e.AncestorsAndSelf().Any(a => (string)a.Attribute("isRecycleBin") == "true" || (string)a.Attribute("isInRecycleBin") == "true");
        internal static string PathOf(XElement e) => string.Join(" / ", e.AncestorsAndSelf().Reverse().Where(IsNode).Select(p => (string)p.Attribute("name") ?? ""));
        private static Dictionary<string, object> Node(XElement e, XElement paths)
        {
            var context = paths.DescendantsAndSelf().FirstOrDefault(p => (string)p.Attribute("ID") == (string)e.Attribute("ID")) ?? e;
            var prefix = context == e ? paths.DescendantsAndSelf().FirstOrDefault(p => (string)p.Attribute("ID") == (string)e.Parent?.Attribute("ID")) : null;
            return new Dictionary<string, object> { ["id"] = (string)e.Attribute("ID"), ["type"] = e.Name.LocalName.ToLowerInvariant(), ["name"] = (string)e.Attribute("name"),
                ["parent_id"] = (string)e.Parent?.Attribute("ID"), ["path"] = prefix == null ? PathOf(context) : PathOf(prefix) + " / " + (string)e.Attribute("name"), ["last_modified"] = (string)e.Attribute("lastModifiedTime") };
        }
        private object Export(string id, string format, string path, CancellationToken cancellation)
        {
            if (!Path.IsPathRooted(path) || path.StartsWith("\\") && !path.StartsWith("\\\\") || path.Length > 1 && path[1] == ':' && (path.Length < 3 || path[2] != '\\' && path[2] != '/')) throw new McpFault("invalid_path", "导出必须使用绝对路径。");
            path = Path.GetFullPath(path);
            if (File.Exists(path) || Directory.Exists(path)) throw new McpFault("file_exists", "导出目标已存在，不会覆盖。");
            if (!Directory.Exists(Path.GetDirectoryName(path))) throw new McpFault("invalid_path", "导出目录不存在。");
            var before = ReadPage(id); var model = new McpPageModel(before);
            var temp = Path.Combine(Path.GetDirectoryName(path), ".onenote-export-" + Guid.NewGuid().ToString("N") + (format == "pdf" ? ".pdf" : ".md"));
            try
            {
                cancellation.ThrowIfCancellationRequested();
                if (format == "pdf") Workspace.ExportPdf(id, temp); else File.WriteAllText(temp, model.Markdown, new UTF8Encoding(false));
                var after = ReadPage(id);
                if ((string)before.Attribute("lastModifiedTime") != (string)after.Attribute("lastModifiedTime") || McpPageModel.Signature(before, true) != McpPageModel.Signature(after, true)) throw new McpFault("export_conflict", "导出期间页面发生变化，未交付文件。");
                if (!File.Exists(temp) || format == "pdf" && new FileInfo(temp).Length == 0) throw new McpFault("export_failed", "导出未生成有效文件。");
                if (format == "pdf") using (var stream = File.OpenRead(temp))
                { var header = new byte[5]; if (stream.Read(header, 0, 5) != 5 || Encoding.ASCII.GetString(header) != "%PDF-") throw new McpFault("export_failed", "原生导出未生成 PDF 文件。"); }
                cancellation.ThrowIfCancellationRequested(); File.Move(temp, path);
                return new { ok = true, page_id = id, output_path = path, format, bytes = new FileInfo(path).Length, last_modified = model.Modified,
                    complete = format == "pdf" || model.Complete, issues = format == "pdf" ? new string[0] : model.Issues.Distinct().ToArray() };
            }
            finally { if (File.Exists(temp)) File.Delete(temp); }
        }
        private void Prune() { foreach (var key in _reads.Where(p => p.Value.Expires <= _now()).Select(p => p.Key).ToArray()) _reads.Remove(key); }
    }
    internal static class McpStringExtensions
    { internal static string SubstringSafe(this string s, int count) { if (s.Length <= count) return s; if (count > 0 && char.IsHighSurrogate(s[count - 1])) count--; return s.Substring(0, count); } }
}
