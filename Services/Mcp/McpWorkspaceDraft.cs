using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using OneNoteCodeHelper.Services.Agent;

namespace OneNoteCodeHelper.Services.Mcp
{
    /// <summary>一次工作区草稿只规划一项操作。任何 COM 写入开始后冻结，随后只核验、查状态或补偿撤销。</summary>
    internal sealed class McpWorkspaceDraft
    {
        private static XNamespace One => OneNoteApi.One;
        private static readonly object HierarchyGate = new object();
        private readonly McpReadService _read;
        private readonly AgentOptions _options;
        private readonly AddInSettings _settings;
        internal string Kind, PageId, SectionId, ParentId, Name, NewId, SourceSection;
        internal XElement Source, Desired, Written;
        internal string SourceSignature, AfterSignature, PlanJson;
        internal bool Started, RecycledSource, CreatedVerified, Undone;
        private string[] _warnings = new string[0];
        private bool _sourceRecycleAttempted, _sourceOutcomeKnown;
        private int _sourceLevel;
        private IOneNoteWorkspaceAccess Api => _read.Workspace;
        internal McpWorkspaceDraft(McpReadService read, AgentOptions options, AddInSettings settings) { _read = read; _options = options; _settings = settings; }

        internal object Stage(string kind, IDictionary<string, object> args)
        {
            var json = McpJson.Serialize(args.OrderBy(a => a.Key).ToDictionary(a => a.Key, a => a.Value));
            if (PlanJson != null)
            { if (Kind != kind || PlanJson != json) throw new McpFault("workspace_operation_limit", "每个工作区草稿只容纳一项操作，请新建草稿。"); return new { ok = true, draft_revision = 1, changed = false }; }
            if (!_options.EnableInsert || kind == "move_page" && !_options.EnableMoves) throw new McpFault("capability_disabled", "页面创建、复制或移动能力已关闭。");
            var section = args.TryGetValue("section_id", out var s) ? (string)s : null;
            XElement source = null, desired = null; string sourceSection = null; var sourceLevel = 1;
            string[] warnings = new string[0];
            if (kind == "create_section")
            {
                var parent = RequireNode((string)args["parent_id"], "Notebook", "SectionGroup"); var name = (string)args["name"];
                RequireRecycleBin(parent);
                if (name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || name != name.Trim() || name.EndsWith(".") || name.Length > 100 || name == "." || name == "..")
                    throw new McpFault("invalid_section_name", "分区名不能含路径字符、首尾空格或末尾句点，最多 100 字。");
                if (parent.Elements(One + "Section").Any(e => string.Equals((string)e.Attribute("name"), name, StringComparison.OrdinalIgnoreCase))) throw new McpFault("section_exists", "同名分区已存在。");
            }
            else
            {
                RequireRecycleBin(RequireNode(section, "Section"));
                if (kind == "create_page")
                {
                    var title = (string)args["title"];
                    if (title.IndexOfAny(new[] { '\r', '\n' }) >= 0 || title.Length > 500) throw new McpFault("invalid_title", "页面标题最多 500 字且不能换行。");
                    var content = McpContent.Parse(args.TryGetValue("content", out var c) ? (string)c : "", args.TryGetValue("format", out var f) ? (string)f : "plain", _options, _settings);
                    warnings = content.Warnings.Distinct().ToArray();
                    desired = new XElement(One + "Page", new XAttribute("ID", "draft"), new XAttribute("name", title),
                        content.Tags.Elements(), content.Styles.Elements(), new XElement(One + "Title", new XElement(One + "OE", new XElement(One + "T", new XCData(OneNoteHtmlEncoder.EncodePlainText(title))))));
                    if (args.TryGetValue("content", out var body) && !string.IsNullOrWhiteSpace((string)body)) { var (x, y) = PageEditor.NextFreePosition(desired); desired.Add(new XElement(One + "Outline", new XElement(One + "Position", new XAttribute("x", x), new XAttribute("y", y), new XAttribute("z", 0)), new XElement(One + "OEChildren", content.Nodes))); }
                }
                else
                {
                    var pageNode = RequireNode((string)args["page_id"], "Page"); sourceSection = (string)pageNode.Parent.Attribute("ID");
                    sourceLevel = Level(pageNode);
                    if (kind == "move_page")
                    {
                        RequireRecycleBin(pageNode);
                        if (section == sourceSection) throw new McpFault("same_section", "源页已经位于目标分区。");
                        if (HasChildren(pageNode)) throw new McpFault("page_has_children", "父页面含子页面，不能按单页移动。");
                    }
                    source = _read.ReadPage((string)args["page_id"], PageInfo.piBinaryData);
                    var basic = _read.ReadPage((string)args["page_id"]);
                    if (source.Descendants(One + "Image").Count() != basic.Descendants(One + "Image").Count()) throw new McpFault("image_unavailable", "图片二进制数据不完整。");
                    EnsureCopyable(source);
                    desired = new XElement(source);
                    foreach (var e in desired.DescendantsAndSelf()) AgentCode.StripIdentity(e);
                    desired.Descendants(One + "CallbackID").Remove();
                }
            }
            // 完成全部预检后才发布草稿；失败调用不能留下半个计划。
            Kind = kind; SectionId = section; Source = source; Desired = desired; SourceSection = sourceSection;
            PageId = args.TryGetValue("page_id", out var pid) ? (string)pid : null;
            ParentId = args.TryGetValue("parent_id", out var p) ? (string)p : null; Name = args.TryGetValue("name", out var n0) ? (string)n0 : null;
            SourceSignature = source == null ? null : McpPageModel.Signature(source, true); PlanJson = json; _warnings = warnings;
            _sourceLevel = sourceLevel;
            return new { ok = true, draft_revision = 1, changed = true, warnings };
        }
        internal object Pending(int revision) => new { ok = true, draft_revision = revision, unread_count = 0, operation = Kind,
            page_id = PageId, section_id = SectionId, parent_id = ParentId, name = Name,
            preview = Desired == null ? null : new McpPageModel(PreviewPage()).Markdown, warnings = _warnings, can_commit = Kind != null };
        private XElement PreviewPage() { var p = new XElement(Desired); p.SetAttributeValue("ID", "draft"); return p; }
        internal bool CanUndo => CreatedVerified && !Undone && (!_sourceRecycleAttempted || _sourceOutcomeKnown);

        internal object Commit(CancellationToken cancellation)
        {
            if (Kind == null) throw new McpFault("operation_required", "请先规划一项页面或分区操作。");
            lock (HierarchyGate)
            lock (PageEditCoordinator.ForPage(PageId ?? "workspace:" + (SectionId ?? ParentId)))
            {
                cancellation.ThrowIfCancellationRequested();
                if (Kind == "create_section")
                {
                    var parent = RequireNode(ParentId, "Notebook", "SectionGroup"); RequireRecycleBin(parent);
                    if (parent.Elements(One + "Section").Any(e => string.Equals((string)e.Attribute("name"), Name, StringComparison.OrdinalIgnoreCase))) throw new McpFault("section_exists", "同名分区已存在，未提交。");
                    Started = true;
                    try { NewId = Api.CreateSection(ParentId, Name); var node = RequireNode(NewId, "Section"); CreatedVerified = (string)node.Parent?.Attribute("ID") == ParentId && (string)node.Attribute("name") == Name && !node.Elements(One + "Page").Any(); return Result(CreatedVerified ? "Verified" : "CommitOutcomeUnknown"); }
                    catch (Exception) { return Result("CommitOutcomeUnknown"); }
                }
                RequireRecycleBin(RequireNode(SectionId, "Section"));
                if (Source != null && McpPageModel.Signature(_read.ReadPage(PageId, PageInfo.piBinaryData), true) != SourceSignature)
                    return Result("NoChange", "源页在规划期间发生变化，未写入。");
                if (Source != null && (string)RequireNode(PageId, "Page").Parent?.Attribute("ID") != SourceSection) return Result("NoChange", "源页位置已变化，未写入。");
                if (Kind == "move_page" && (HasChildren(RequireNode(PageId, "Page")) || Level(RequireNode(PageId, "Page")) != _sourceLevel)) return Result("NoChange", "源页层级或子页面发生变化，未写入。");
                Started = true;
                try
                {
                    NewId = Api.CreatePage(SectionId);
                    var fresh = _read.ReadPage(NewId); var desired = ForNewPage(NewId, fresh);
                    try { Api.UpdatePageContent(desired.ToString(SaveOptions.DisableFormatting), AgentPageSnapshot.Modified(fresh)); } catch (Exception) { /* 写入可能成功，只回读，绝不重放。 */ }
                    Written = _read.ReadPage(NewId, PageInfo.piBinaryData);
                    CreatedVerified = McpPageModel.Signature(desired, false) == McpPageModel.Signature(Written, false) && (string)RequireNode(NewId, "Page").Parent?.Attribute("ID") == SectionId;
                    if (!CreatedVerified) return Result("CommitOutcomeUnknown");
                    AfterSignature = McpPageModel.Signature(Written, true);
                    if (Kind == "move_page")
                    {
                        var sourceNode = RequireNode(PageId, "Page"); RequireRecycleBin(sourceNode);
                        var sourceNow = _read.ReadPage(PageId, PageInfo.piBinaryData);
                        sourceNode = RequireNode(PageId, "Page");
                        if (McpPageModel.Signature(sourceNow, true) != SourceSignature || (string)sourceNode.Parent?.Attribute("ID") != SourceSection || HasChildren(sourceNode) || Level(sourceNode) != _sourceLevel)
                            return Result("PartiallyApplied", "目标副本已核验，源页发生变化，源页保留。");
                        _sourceRecycleAttempted = true;
                        try { Api.Recycle(PageId, AgentPageSnapshot.Modified(sourceNow)); }
                        catch (Exception)
                        {
                            // COM 可能已执行删除再断开；只有确认源页状态后，才允许补偿撤销。
                            RecycledSource = !Exists(PageId); _sourceOutcomeKnown = true;
                            return Result(RecycledSource ? "Verified" : "PartiallyApplied", RecycledSource ? "源页已确认进入回收站。" : "目标副本已核验，源页保留。");
                        }
                        RecycledSource = !Exists(PageId);
                        _sourceOutcomeKnown = true;
                        if (!RecycledSource) return Result("PartiallyApplied", "目标副本已核验，源页未确认进入回收站。");
                    }
                    return Result("Verified");
                }
                catch (Exception ex) { return Result("CommitOutcomeUnknown", "工作区写回或核验失败（" + ex.GetType().Name + "），禁止重发。"); }
            }
        }
        internal object Undo(CancellationToken cancellation)
        {
            lock (HierarchyGate)
            lock (PageEditCoordinator.ForPage(NewId))
            {
                cancellation.ThrowIfCancellationRequested();
                if (!CanUndo) throw new McpFault("undo_unavailable", "没有可撤销的已核验工作区操作。");
                if (Kind == "create_section")
                {
                    var node = RequireNode(NewId, "Section"); RequireRecycleBin(node);
                    if (node.Elements(One + "Page").Any() || (string)node.Parent?.Attribute("ID") != ParentId || (string)node.Attribute("name") != Name) return Result("NoChange", "分区已有内容、名称或位置变化，保留分区。");
                    Api.Recycle(NewId, NodeModified(node)); Undone = !Exists(NewId); return Result(Undone ? "Verified" : "CommitOutcomeUnknown");
                }
                var current = _read.ReadPage(NewId, PageInfo.piBinaryData);
                var targetNode = RequireNode(NewId, "Page"); RequireRecycleBin(targetNode);
                if (HasChildren(targetNode)) return Result("NoChange", "目标页后来增加了子页面，撤销停止。");
                if (McpPageModel.Signature(current, true) != AfterSignature || (string)RequireNode(NewId, "Page").Parent?.Attribute("ID") != SectionId)
                    return Result("NoChange", "目标页已被编辑或移动，撤销跳过。");
                string restored = null;
                if (Kind == "move_page" && RecycledSource)
                {
                    if (Exists(PageId)) return Result("NoChange", "源页已恢复，撤销停止，请核对源页和目标页。");
                    RequireRecycleBin(RequireNode(SourceSection, "Section"));
                    restored = Api.CreatePage(SourceSection); var before = _read.ReadPage(restored); var desired = ForNewPage(restored, before);
                    Api.UpdatePageContent(desired.ToString(SaveOptions.DisableFormatting), AgentPageSnapshot.Modified(before));
                    if (McpPageModel.Signature(desired, false) != McpPageModel.Signature(_read.ReadPage(restored, PageInfo.piBinaryData), false))
                        return new { status = "CommitOutcomeUnknown", restored_page_id = restored, message = "源位置恢复副本未核验，目标页保留。" };
                    // 恢复期间目标页可能被用户修改：重新检查再删除。
                    current = _read.ReadPage(NewId, PageInfo.piBinaryData);
                    targetNode = RequireNode(NewId, "Page");
                    if (McpPageModel.Signature(current, true) != AfterSignature || (string)targetNode.Parent?.Attribute("ID") != SectionId || HasChildren(targetNode))
                        return new { status = "PartiallyApplied", restored_page_id = restored, message = "恢复副本已建立，目标页发生变化并保留。" };
                }
                Api.Recycle(NewId, AgentPageSnapshot.Modified(current)); Undone = !Exists(NewId);
                return new { status = Undone ? "Verified" : "CommitOutcomeUnknown", restored_page_id = restored, restored_link = restored == null ? null : Api.GetLink(restored, null), message = "撤销完成后页面 ID 可能变化。" };
            }
        }
        private object Result(string status, string message = null) => new { status, applied = CreatedVerified ? 1 : 0, operation = Kind,
            source_page_id = PageId, page_id = Kind == "create_section" ? null : NewId, section_id = Kind == "create_section" ? NewId : SectionId,
            source_recycled = RecycledSource, source_outcome_known = !_sourceRecycleAttempted || _sourceOutcomeKnown,
            creation_attempted = Started, link = NewId == null ? null : SafeLink(NewId), object_id_map = ObjectMap(), message };
        private string SafeLink(string id) { try { return Api.GetLink(id, null); } catch (Exception) { return null; } }
        private XElement ForNewPage(string id, XElement fresh)
        {
            var desired = new XElement(Desired); desired.SetAttributeValue("ID", id);
            // 原生标题栏固定在新建页已有容器中；复制及补偿恢复使用相同规则。
            desired.Element(One + "Title")?.SetAttributeValue("objectID", (string)fresh.Element(One + "Title")?.Attribute("objectID"));
            desired.Element(One + "Title")?.Element(One + "OE")?.SetAttributeValue("objectID", (string)fresh.Element(One + "Title")?.Element(One + "OE")?.Attribute("objectID"));
            return desired;
        }
        private object[] ObjectMap()
        {
            if (Source == null || Written == null || !CreatedVerified) return new object[0];
            var names = new[] { "Title", "Outline", "OE", "Table", "Row", "Cell", "Image" };
            return Source.Descendants().Where(e => names.Contains(e.Name.LocalName)).Zip(Written.Descendants().Where(e => names.Contains(e.Name.LocalName)),
                (a, b) => (object)new { source_id = (string)a.Attribute("objectID"), object_id = (string)b.Attribute("objectID") }).ToArray();
        }
        private XElement RequireNode(string id, params string[] kinds)
        {
            var tree = McpReadService.Xml(Api.GetHierarchy("", HierarchyScope.hsPages));
            var node = tree.DescendantsAndSelf().FirstOrDefault(e => (string)e.Attribute("ID") == id && kinds.Contains(e.Name.LocalName) && McpReadService.IsLive(e));
            if (node == null) throw new McpFault("node_unavailable", "目标层级节点不存在或类型不正确。");
            if (node.AncestorsAndSelf().Any(e => (string)e.Attribute("isReadOnly") == "true" || (string)e.Attribute("isLocked") == "true")) throw new McpFault("node_protected", "节点只读或已锁定。");
            return node;
        }
        private bool Exists(string id) => McpReadService.Xml(Api.GetHierarchy("", HierarchyScope.hsPages)).DescendantsAndSelf().Any(e => (string)e.Attribute("ID") == id && McpReadService.IsLive(e));
        private static void RequireRecycleBin(XElement node)
        {
            var notebook = node.AncestorsAndSelf(One + "Notebook").FirstOrDefault();
            if (notebook == null || !notebook.Descendants(One + "SectionGroup").Any(e => (string)e.Attribute("isRecycleBin") == "true"))
                throw new McpFault("recycle_bin_unavailable", "笔记本没有可确认的回收站，无法保证可撤销写入。请使用支持回收站的笔记本。");
        }
        private static DateTime NodeModified(XElement node) => DateTime.TryParse((string)node.Attribute("lastModifiedTime"), null, System.Globalization.DateTimeStyles.RoundtripKind, out var dt) ? dt : DateTime.MinValue;
        private static int Level(XElement page) => int.TryParse((string)page?.Attribute("pageLevel"), out var n) ? n : 1;
        private static bool HasChildren(XElement page) => page.ElementsAfterSelf(One + "Page").FirstOrDefault() is XElement next && Level(next) > Level(page);
        private static void EnsureCopyable(XElement page)
        {
            var allowed = new HashSet<string>(new[] { "Page", "Title", "Outline", "OEChildren", "OE", "T", "Table", "Columns", "Column", "Row", "Cell", "Image", "Data", "Position", "Size", "Meta", "QuickStyleDef", "TagDef", "Tag", "List", "Bullet", "Number", "OCRData", "OCRText", "OCRToken", "PageSettings" });
            if (page.DescendantsAndSelf().Any(e => e.Name.Namespace != One || !allowed.Contains(e.Name.LocalName))) throw new McpFault("unsupported_copy", "页面包含墨迹、附件、音视频或未支持对象，不能保真复制。");
            if (page.Descendants(One + "Image").Any(i => string.IsNullOrEmpty(i.Element(One + "Data")?.Value))) throw new McpFault("image_unavailable", "图片二进制数据不完整。");
            foreach (var image in page.Descendants(One + "Image"))
                try { if (Convert.FromBase64String(image.Element(One + "Data").Value).Length == 0) throw new FormatException(); }
                catch (FormatException) { throw new McpFault("image_unavailable", "图片二进制数据无效。"); }
            foreach (var oe in page.Descendants(One + "OE").Where(e => e.Elements(One + "T").Any()))
                if (!RichParagraph.Parse(oe).IsLossless) throw new McpFault("unsupported_copy", "页面包含不能无损解析的文字格式。");
        }
    }
}
