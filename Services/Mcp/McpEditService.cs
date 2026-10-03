using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web.Script.Serialization;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using OneNoteCodeHelper.Mcp;
using OneNoteCodeHelper.Services.Agent;

namespace OneNoteCodeHelper.Services.Mcp
{
    internal sealed class McpFault : Exception
    {
        internal string Code { get; }
        internal McpFault(string code, string message) : base(message) { Code = code; }
    }

    internal static class McpJson
    {
        internal static JavaScriptSerializer Serializer() => new JavaScriptSerializer { MaxJsonLength = McpProtocol.MaxFrameBytes * 2, RecursionLimit = 128 };
        internal static string Serialize(object value) => Serializer().Serialize(value);
        internal static T Deserialize<T>(string value) => Serializer().Deserialize<T>(value);
        internal static IDictionary<string, object> Arguments(string value) =>
            Serializer().DeserializeObject(value) as IDictionary<string, object> ?? throw new AiException("工具参数必须是对象。");
    }

    /// <summary>插件拥有全部编辑、核验和撤销状态。IPC 层不接触模型或 OneNote 页面内容。</summary>
    internal sealed class McpEditService : IDisposable
    {
        private sealed class Entry
        {
            internal readonly object Gate = new object();
            internal string Owner;
            internal AgentEditSession Session;
            internal McpWorkspaceDraft Workspace;
            internal AgentOptions WorkspaceOptions;
            internal AddInSettings Settings;
            internal readonly Dictionary<string, object> ContentResults = new Dictionary<string, object>();
            internal AgentOptions Options => Session?.Snapshot.Options ?? WorkspaceOptions;
            internal string PageId => Session?.Snapshot.PageId ?? Workspace?.PageId;
            internal bool Frozen => CompletedTicks != 0 || (Session?.Snapshot.Frozen ?? Workspace?.Started ?? false);
            internal CancellationTokenSource Cancellation;
            internal volatile string State = "Draft";
            internal volatile string ResultJson;
            internal volatile string UndoJson;
            internal volatile bool Busy;
            internal volatile bool CanUndo;
            internal volatile bool Retired;
            internal int Calls;
            internal int Revision;
            internal int SubmittedRevision = -1;
            internal long CompletedTicks;
            internal long CompletionOrder;
            internal DateTime Expires;
        }

        private readonly object _gate = new object();
        private readonly Dictionary<string, Entry> _entries = new Dictionary<string, Entry>();
        private readonly Dictionary<string, string> _clients = new Dictionary<string, string>();
        private readonly HashSet<string> _connected = new HashSet<string>();
        private readonly CancellationTokenSource _shutdown = new CancellationTokenSource();
        private int _beginning;
        private long _completionOrder;
        private readonly IOneNotePageAccess _api;
        private readonly AgentCommitter _committer;
        private readonly McpReadService _read;
        private readonly Func<string> _currentPage;
        private readonly Func<AgentOptions> _options;
        private readonly Func<AddInSettings> _settings;
        private readonly Func<DateTime> _now;
        private readonly Timer _cleanup;
        private volatile bool _stopped;
        internal string InstanceId { get; } = Guid.NewGuid().ToString("N");
        internal McpCatalog Catalog { get; }
        private static readonly Dictionary<string, AgentSchema> Management = new Dictionary<string, AgentSchema>
        {
            ["get_status"] = AgentSchema.Obj(new Dictionary<string, AgentSchema>()),
            ["get_current_page"] = AgentSchema.Obj(new Dictionary<string, AgentSchema>()),
            ["begin_edit"] = AgentSchema.Obj(new Dictionary<string, AgentSchema> { ["scope"] = AgentSchema.Str("page", "selection"), ["page_id"] = AgentSchema.Str() }, "scope"),
            ["get_edit_status"] = SnapshotSchema(), ["abort_edit"] = SnapshotSchema(), ["undo_edit"] = SnapshotSchema()
        };
        private static AgentSchema SnapshotSchema() => AgentSchema.Obj(new Dictionary<string, AgentSchema> { ["snapshot_id"] = AgentSchema.Str() }, "snapshot_id");

        internal static AgentOptions LoadConfiguredOptions()
        {
            try
            {
                // 外部会话只需要 Agent 设置；不解析模型、Key，也不记录配置解析异常正文。
                return File.Exists(AiConfigStore.ConfigPath)
                    ? AgentOptions.Parse(XDocument.Load(AiConfigStore.ConfigPath).Root?.Element("Agent")) : new AgentOptions();
            }
            catch (Exception)
            {
                AddInLog.Info("MCP settings state=default");
                return new AgentOptions();
            }
        }

        internal McpEditService(IOneNotePageAccess api, Func<string> currentPage, Func<AgentOptions> options,
            Func<AddInSettings> settings, Func<DateTime> now = null)
        {
            _api = api; _committer = new AgentCommitter(api); _currentPage = currentPage;
            _options = options; _settings = settings; _now = now ?? (() => DateTime.UtcNow);
            _read = new McpReadService(api, currentPage, _now, options);
            Catalog = BuildCatalog();
            _cleanup = new Timer(_ => Prune(), null, 60000, 60000);
        }

        private static McpCatalog BuildCatalog()
        {
            var agent = AgentTools.Catalog();
            var tools = agent.Definitions.Select(d =>
            {
                var f = (IDictionary<string, object>)McpJson.Arguments(McpJson.Serialize(d))["function"];
                var schema = (IDictionary<string, object>)f["parameters"];
                var name = (string)f["name"];
                if (name == "get_page_overview")
                {
                    ((IDictionary<string, object>)schema["properties"])["snapshot_id"] = AgentSchema.Str().Json();
                    schema["required"] = new[] { "snapshot_id" };
                }
                return new McpToolDefinition { Name = name, Description = (string)f["description"], InputSchemaJson = McpJson.Serialize(schema),
                    ReadOnly = name == "get_page_overview" || name == "read_blocks" || name == "read_image_text" || name == "get_pending_changes",
                    Destructive = name == "finish_edit" };
            }).ToList();
            var descriptions = new Dictionary<string, string>
            {
                ["get_status"] = "检查插件连接、实例标识、IPC 版本和编辑记录数量；不读取笔记正文。",
                ["get_current_page"] = "获取 OneNote 当前页面的标识、标题和选区段落数，不创建草稿。",
                ["begin_edit"] = "固定页面、选区和能力设置，创建编辑快照。scope=page 可用 page_id 指定页，省略保持当前页行为；selection 只允许当前页。返回 snapshot_id、可用工具及配额。",
                ["get_edit_status"] = "查询快照状态、提交结果和是否可以撤销。请求取消、断线或写回未确认后先查询；不要自动重发写入。",
                ["abort_edit"] = "取消未提交草稿。写入已经开始时继续核验实际结果，随后用 get_edit_status 查询。",
                ["undo_edit"] = "凭本客户端的快照撤销已核验的修改；仍使用页面冲突检查及整组撤销。重复调用返回原撤销结果，不能再次执行。"
            };
            foreach (var pair in Management) tools.Add(new McpToolDefinition { Name = pair.Key, Description = descriptions[pair.Key],
                InputSchemaJson = McpJson.Serialize(pair.Value.Json()), ReadOnly = pair.Key.StartsWith("get_", StringComparison.Ordinal), Destructive = pair.Key == "undo_edit" });
            tools.AddRange(McpExtensionCatalog.Definitions);
            return new McpCatalog { Tools = tools.ToArray(), Instructions =
                "只读查找使用 list_nodes/search_pages/read_page/find_tasks，无需编辑草稿。页面编辑先 begin_edit(scope,page_id 可选)；页面或分区操作先 begin_workspace_edit，每个草稿只规划一项操作。新增内容和工作区工具仅供外部 MCP 使用。所有笔记写入都先 get_pending_changes 再 finish_edit。后续工具携带同一个 snapshot_id。目录包含全部工具，实际可用工具以 begin_edit 返回的 available_tools 为准。" +
                AgentRunner.SystemPrompt(agent) +
                "上述内置格式助手的文字限制仅适用旧工具。外部用户明确要求新增正文或改写时，可用 append_content/insert_content/replace_text；replace_text 仍不能跨段落，代码框继续保护。只读结果的真实 object_id 不能直接用作草稿 p1 短 ID，须重新读取编辑快照进行匹配。" +
                "外部任务必须完整读取范围，不存在最后一轮提前提交例外。请求取消或断线后先 get_edit_status；CommitOutcomeUnknown 表示写回未确认，禁止自动重发写入。" +
                "已完成记录和撤销只在当前插件实例内保留 30 分钟、最多 16 条。插件重启后旧快照失效。外部编辑通过 undo_edit 撤销。" };
        }

        internal void BindClient(string client, string secret)
        {
            if (!Guid.TryParseExact(client, "N", out _) || !Guid.TryParseExact(secret, "N", out _)) throw new McpFault("invalid_client", "客户端标识无效。");
            lock (_gate)
            {
                EnsureRunning(); PruneLocked();
                if (_clients.TryGetValue(client, out var previous) && previous != secret) throw new McpFault("client_mismatch", "客户端恢复凭据不匹配。");
                // 每条连接有独立身份；完成记录以恢复凭据绑定，其他进程不能认领。
                if (_clients.Count >= 128 && !_clients.ContainsKey(client)) throw new McpFault("client_limit", "客户端数量达到上限。");
                _clients[client] = secret;
                _connected.Add(client);
            }
        }

        internal string Call(string client, string name, string arguments, CancellationToken cancellation, Action<int, string> progress = null)
        {
            EnsureRunning();
            lock (_gate) if (!_clients.ContainsKey(client)) throw new McpFault("invalid_client", "客户端未连接。");
            cancellation.ThrowIfCancellationRequested();
            var args = McpJson.Arguments(arguments);
            if (McpExtensionCatalog.Schemas.TryGetValue(name, out var extension))
            {
                extension.Validate(args);
                if (name == "begin_workspace_edit") return BeginWorkspace(client, cancellation);
                if (!McpExtensionCatalog.ContentTools.Contains(name) && !McpExtensionCatalog.WorkspaceTools.Contains(name))
                    return McpJson.Serialize(_read.Call(client, name, args, cancellation, progress));
            }
            if (Management.TryGetValue(name, out var schema))
            {
                schema.Validate(args);
                switch (name)
                {
                    case "get_status":
                        lock (_gate) { PruneLocked(); return McpJson.Serialize(new { connected = true, instance_id = InstanceId, ipc_version = McpProtocol.Version,
                            active_edits = _entries.Values.Count(e => e.CompletedTicks == 0), retained_results = _entries.Values.Count(e => e.CompletedTicks != 0) }); }
                    case "get_current_page":
                        var page = ReadCurrentPage();
                        return McpJson.Serialize(new { page_id = (string)page.Attribute("ID"), page_title = (string)page.Attribute("name"), selected_paragraphs = AgentPageSnapshot.SelectedIds(page).Count });
                    case "begin_edit": return Begin(client, (string)args["scope"], args.TryGetValue("page_id", out var pid) ? (string)pid : null, cancellation);
                }
            }
            if (!args.TryGetValue("snapshot_id", out var id) || !(id is string snapshotId)) throw new McpFault("snapshot_required", "必须提供 snapshot_id。");
            var entry = Find(client, snapshotId);
            if (name == "get_edit_status") return Status(snapshotId, entry);
            if (name == "abort_edit")
            {
                Cancel(entry);
                lock (_gate) if (!entry.Busy && entry.CompletedTicks == 0) CompleteLocked(entry, "CancelledBeforeCommit");
                return Status(snapshotId, entry);
            }
            lock (entry.Gate)
            {
                if (entry.Retired) throw new McpFault("snapshot_expired", "快照已过期或被淘汰。");
                if (name == "undo_edit") return Undo(entry, cancellation);
                if (name == "get_page_overview") args.Remove("snapshot_id");
                var call = new AgentToolCall { Name = name, Arguments = McpJson.Serialize(args) };
                if (entry.Workspace == null && extension == null) entry.Session.Tools.ValidateArguments(call);
                if (entry.Workspace != null && extension == null && name != "get_pending_changes" && name != "finish_edit") throw new McpFault("scope_mismatch", "工作区草稿只支持页面或分区操作。");
                if (entry.Workspace != null && name == "finish_edit")
                    AgentSchema.Obj(new Dictionary<string, AgentSchema> { ["snapshot_id"] = AgentSchema.Str(), ["draft_revision"] = AgentSchema.Num(0, 100000, true) }, "snapshot_id", "draft_revision").Validate(args);
                if (name == "finish_edit" && entry.ResultJson != null)
                {
                    if (Convert.ToInt32(args["draft_revision"]) != entry.SubmittedRevision) throw new McpFault("revision_conflict", "快照已提交，提交修订号不一致。");
                    return entry.ResultJson;
                }
                if (entry.Frozen) throw new McpFault("edit_frozen", "草稿已结束或冻结，请查询状态。");
                using (cancellation.Register(() => Cancel(entry)))
                {
                    entry.Cancellation.Token.ThrowIfCancellationRequested();
                    if (entry.Calls >= entry.Options.MaxToolCalls) { Cancel(entry); throw new McpFault("tool_limit", "工具调用达到上限，未提交草稿。"); }
                    Interlocked.Increment(ref entry.Calls);
                    lock (_gate) { entry.Busy = true; entry.State = name == "finish_edit" ? "Committing" : "Draft"; }
                    try
                    {
                        if (name == "finish_edit") entry.SubmittedRevision = Convert.ToInt32(args["draft_revision"]);
                        object outcome;
                        if (entry.Workspace != null)
                        {
                            if (name == "get_pending_changes") outcome = entry.Workspace.Pending(entry.Revision);
                            else if (name == "finish_edit")
                            {
                                if (Convert.ToInt32(args["draft_revision"]) != entry.Revision) throw new McpFault("revision_conflict", "提交修订号不是最新草稿。");
                                progress?.Invoke(0, "提交工作区操作"); outcome = entry.Workspace.Commit(entry.Cancellation.Token);
                                entry.ResultJson = McpJson.Serialize(outcome); entry.CanUndo = entry.Workspace.CanUndo;
                                Complete(entry, (string)McpJson.Arguments(entry.ResultJson)["status"]);
                            }
                            else if (McpExtensionCatalog.WorkspaceTools.Contains(name)) { outcome = entry.Workspace.Stage(name, args); entry.Revision = 1; }
                            else throw new McpFault("scope_mismatch", "内容工具需要页面编辑草稿。");
                        }
                        else if (McpExtensionCatalog.ContentTools.Contains(name))
                        {
                            var key = name + McpJson.Serialize(args.OrderBy(a => a.Key).ToDictionary(a => a.Key, a => a.Value));
                            if (!entry.ContentResults.TryGetValue(key, out outcome)) { outcome = McpContent.Apply(entry.Session.Snapshot, entry.Settings, name, args); entry.ContentResults.Add(key, outcome); }
                        }
                        else if (McpExtensionCatalog.WorkspaceTools.Contains(name)) throw new McpFault("scope_mismatch", "页面或分区操作需要工作区草稿。");
                        else { if (name == "finish_edit") progress?.Invoke(0, "提交并核验页面"); outcome = entry.Session.Execute(call); }
                        if (entry.Session != null && (name == "get_page_overview" || name == "read_blocks" || name == "read_image_text"))
                            outcome = WithObjectIds(outcome, entry.Session.Snapshot);
                        var result = McpJson.Serialize(outcome);
                        if (entry.Session != null) entry.Revision = entry.Session.Snapshot.Revision;
                        if (entry.Session?.Snapshot.Frozen == true)
                        {
                            entry.ResultJson = result; entry.CanUndo = entry.Session.Report?.CanUndo == true;
                            Complete(entry, entry.Session.Report?.Status ?? "CommitOutcomeUnknown");
                        }
                        else if (entry.CompletedTicks == 0) entry.State = "Draft";
                        progress?.Invoke(entry.Calls, "工具调用完成");
                        return result;
                    }
                    catch (OperationCanceledException)
                    {
                        Complete(entry, "CancelledBeforeCommit");
                        throw;
                    }
                    catch (Exception) when (entry.Frozen)
                    {
                        // 发生非预期异常时保留不确定结果，冻结快照，绝不重新执行。
                        entry.ResultJson = McpJson.Serialize(new { status = "CommitOutcomeUnknown", message = "写回未确认，请查看页面并查询状态，禁止自动重发。" });
                        Complete(entry, "CommitOutcomeUnknown"); return entry.ResultJson;
                    }
                    finally
                    {
                        lock (_gate)
                        {
                            entry.Busy = false;
                            if (entry.CompletedTicks == 0 && entry.Cancellation.IsCancellationRequested) CompleteLocked(entry, "CancelledBeforeCommit");
                            else if (entry.CompletedTicks == 0) entry.State = "Draft";
                            PruneLocked();
                            if (entry.Retired) entry.Cancellation.Dispose();
                        }
                        AddInLog.Info($"MCP tool={name} snapshot={snapshotId} count={entry.Calls} state={entry.State}");
                    }
                }
            }
        }

        private string Begin(string client, string scope, string pageId, CancellationToken cancellation)
        {
            // 预留名额后释放管理锁。读 COM 或配置时不能阻塞状态查询和宿主断开。
            lock (_gate)
            {
                EnsureRunning(); PruneLocked();
                if (_beginning + _entries.Values.Count(e => e.CompletedTicks == 0) >= 8) throw new McpFault("edit_limit", "未完成草稿达到上限，请先取消已有草稿。");
                _beginning++;
            }
            using (var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellation, _shutdown.Token))
            try
            {
                if (scope == "selection" && pageId != null && pageId != _currentPage()) throw new McpFault("scope_mismatch", "指定页选区只允许当前页面。");
                var page = pageId == null ? ReadCurrentPage() : _read.ReadPage(pageId, scope == "selection" ? PageInfo.piSelection : PageInfo.piBasic);
                var selection = scope == "selection" ? AgentPageSnapshot.SelectedIds(page) : null;
                if (selection != null && selection.Count == 0) throw new McpFault("selection_empty", "没有选中段落，请在 OneNote 中选择段落后重试。");
                // 和 Agent 窗口一致：选区里的空行也算进范围，删空行和代码框转换都要用到。
                selection?.UnionWith(PageEditor.FindSelectedBlankLines(page));
                var options = AgentOptions.Parse(new XElement("Agent", _options().ToElements()));
                try { AgentTools.UseInstalledFont(options); }
                catch (AiException ex) { throw new McpFault("font_unavailable", ex.Message); }
                // 自动缩进会改变段落和文本框的 XML 指纹；保留原有空白，不添加序列化排版。
                var snapshot = new AgentPageSnapshot(page.ToString(SaveOptions.DisableFormatting), selection, options);
                linked.Token.ThrowIfCancellationRequested();
                var source = new CancellationTokenSource(TimeSpan.FromSeconds(options.TimeoutSeconds));
                var entry = new Entry { Owner = client, Cancellation = source, Expires = _now().AddSeconds(options.TimeoutSeconds),
                    Settings = _settings().Clone() };
                entry.Session = new AgentEditSession(snapshot, _committer, source.Token, entry.Settings);
                lock (_gate)
                {
                    if (_stopped || !_connected.Contains(client) || linked.IsCancellationRequested)
                    { source.Dispose(); throw new OperationCanceledException(linked.Token); }
                    _entries.Add(snapshot.SnapshotId, entry);
                }
                return McpJson.Serialize(new { snapshot_id = snapshot.SnapshotId, instance_id = InstanceId, page_id = snapshot.PageId, page_title = snapshot.Title,
                    scope, draft_revision = 0, available_tools = Catalog.Tools.Where(t => entry.Session.Tools.Has(t.Name) || McpExtensionCatalog.ContentTools.Contains(t.Name) && (t.Name == "replace_text" || options.EnableInsert && (t.Name != "append_content" || scope == "page"))).Select(t => t.Name).ToArray(),
                    max_tool_calls = options.MaxToolCalls, timeout_seconds = options.TimeoutSeconds, expires_at = entry.Expires.ToString("o") });
            }
            finally { lock (_gate) _beginning--; }
        }

        private static object WithObjectIds(object outcome, AgentPageSnapshot snapshot)
        {
            var result = McpJson.Arguments(McpJson.Serialize(outcome));
            foreach (var field in new[] { "blocks", "tables", "images", "skipped" })
            {
                if (!result.TryGetValue(field, out var value) || !(value is System.Collections.IEnumerable items)) continue;
                foreach (var item in items)
                    if (item is IDictionary<string, object> map && map.TryGetValue("id", out var id))
                        map["object_id"] = snapshot.Blocks.FirstOrDefault(b => b.Id == (string)id)?.ObjectId ??
                            snapshot.Tables.FirstOrDefault(t => t.Id == (string)id)?.ObjectId ?? snapshot.Images.FirstOrDefault(i => i.Id == (string)id)?.ObjectId;
            }
            return result;
        }

        private string BeginWorkspace(string client, CancellationToken cancellation)
        {
            lock (_gate)
            {
                EnsureRunning(); PruneLocked();
                if (_beginning + _entries.Values.Count(e => e.CompletedTicks == 0) >= 8) throw new McpFault("edit_limit", "未完成草稿达到上限。");
                _beginning++;
            }
            try
            {
                var api = _read.Workspace;
                var options = AgentOptions.Parse(new XElement("Agent", _options().ToElements()));
                AgentTools.UseInstalledFont(options);
                var settings = _settings().Clone(); var id = Guid.NewGuid().ToString("N");
                var source = new CancellationTokenSource(TimeSpan.FromSeconds(options.TimeoutSeconds));
                var entry = new Entry { Owner = client, Cancellation = source, Expires = _now().AddSeconds(options.TimeoutSeconds), WorkspaceOptions = options,
                    Settings = settings, Workspace = new McpWorkspaceDraft(_read, options, settings) };
                lock (_gate)
                {
                    if (_stopped || !_connected.Contains(client) || cancellation.IsCancellationRequested) { source.Dispose(); throw new OperationCanceledException(); }
                    _entries.Add(id, entry);
                }
                return McpJson.Serialize(new { snapshot_id = id, instance_id = InstanceId, scope = "workspace", draft_revision = 0,
                    available_tools = McpExtensionCatalog.WorkspaceTools.Concat(new[] { "get_pending_changes", "finish_edit" }).ToArray(),
                    max_tool_calls = options.MaxToolCalls, timeout_seconds = options.TimeoutSeconds, expires_at = entry.Expires.ToString("o") });
            }
            finally { lock (_gate) _beginning--; }
        }

        private XElement ReadCurrentPage()
        {
            var id = _currentPage();
            if (string.IsNullOrEmpty(id)) throw new McpFault("page_unavailable", "OneNote 没有当前页面。");
            var page = AgentPageSnapshot.ParsePage(_api.GetPageContent(id, PageInfo.piSelection));
            if ((string)page.Attribute("ID") != id) throw new McpFault("page_mismatch", "当前页面读取不一致，未创建草稿。");
            return page;
        }

        private Entry Find(string client, string id)
        {
            lock (_gate)
            {
                PruneLocked();
                if (!_entries.TryGetValue(id, out var entry)) throw new McpFault("snapshot_expired", "快照已失效、过期或被淘汰；不能跨插件实例使用。");
                if (entry.Owner != client) throw new McpFault("snapshot_owner", "快照属于其他客户端。");
                return entry;
            }
        }

        private string Status(string id, Entry entry) => McpJson.Serialize(new { snapshot_id = id, instance_id = InstanceId, page_id = entry.PageId,
            state = entry.State, draft_revision = entry.Revision, tool_calls = entry.Calls, busy = entry.Busy,
            cancellation_requested = entry.Cancellation.IsCancellationRequested, can_undo = entry.CanUndo,
            result = entry.ResultJson == null ? null : McpJson.Serializer().DeserializeObject(entry.ResultJson),
            undo_result = entry.UndoJson == null ? null : McpJson.Serializer().DeserializeObject(entry.UndoJson) });

        private string Undo(Entry entry, CancellationToken cancellation)
        {
            if (entry.UndoJson != null) return entry.UndoJson;
            if (entry.ResultJson == null || !entry.CanUndo) throw new McpFault("undo_unavailable", "没有可撤销的已核验修改。");
            cancellation.ThrowIfCancellationRequested();
            lock (_gate) { entry.Busy = true; entry.State = "Undoing"; }
            // 撤销有独立超时，不沿用已经结束的草稿取消令牌。
            using (var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(entry.Options.TimeoutSeconds)))
            using (var linked = CancellationTokenSource.CreateLinkedTokenSource(timeout.Token, cancellation))
            {
                try
                {
                    var outcome = entry.Workspace == null ? _committer.Undo(entry.PageId, entry.Session.Report, entry.Options, linked.Token).ToToolResult() : entry.Workspace.Undo(linked.Token);
                    entry.UndoJson = McpJson.Serialize(outcome);
                    entry.State = "Undo" + (string)McpJson.Arguments(entry.UndoJson)["status"];
                }
                catch (OperationCanceledException)
                {
                    // Committer 只会在开始写入前抛出取消；以后仍完成回读核验。
                    entry.State = entry.Session?.Report.Status ?? "Verified";
                    throw;
                }
                catch (Exception)
                {
                    entry.UndoJson = McpJson.Serialize(new { status = "CommitOutcomeUnknown", message = "撤销写回未确认，禁止自动重发。" });
                    entry.State = "UndoOutcomeUnknown";
                }
                finally
                {
                    if (entry.UndoJson != null) { entry.Session?.Report.ClearUndo(); entry.CanUndo = false; }
                    lock (_gate) { entry.Busy = false; PruneLocked(); if (entry.Retired) entry.Cancellation.Dispose(); }
                }
            }
            return entry.UndoJson;
        }

        internal void DisconnectClient(string client)
        {
            lock (_gate)
            {
                _connected.Remove(client);
                foreach (var entry in _entries.Values.Where(e => e.Owner == client && e.CompletedTicks == 0))
                {
                    Cancel(entry);
                    if (!entry.Busy) CompleteLocked(entry, "CancelledBeforeCommit");
                }
                PruneLocked();
                if (!_entries.Values.Any(e => e.Owner == client)) _clients.Remove(client);
            }
        }

        private void Complete(Entry entry, string state) { lock (_gate) CompleteLocked(entry, state); }
        private void CompleteLocked(Entry entry, string state)
        {
            entry.State = state;
            if (entry.CompletedTicks == 0)
            {
                entry.CompletionOrder = ++_completionOrder;
                Interlocked.Exchange(ref entry.CompletedTicks, _now().Ticks);
            }
            entry.Cancellation.CancelAfter(Timeout.Infinite);
            if (entry.Session != null) entry.Session.Snapshot.Frozen = true;
        }
        private void Prune() { lock (_gate) PruneLocked(); }
        private void PruneLocked()
        {
            foreach (var entry in _entries.Values.Where(e => e.CompletedTicks == 0 && !e.Busy && (e.Cancellation.IsCancellationRequested || _now() >= e.Expires)))
            { Cancel(entry); CompleteLocked(entry, "CancelledBeforeCommit"); }
            var finished = _entries.Where(p => p.Value.CompletedTicks != 0).OrderBy(p => p.Value.CompletionOrder).ToList();
            var discard = finished.Where(p => _now().Ticks - p.Value.CompletedTicks >= TimeSpan.FromMinutes(30).Ticks)
                .Concat(finished.Take(Math.Max(0, finished.Count - 16))).Select(p => p.Key).Distinct().ToArray();
            foreach (var id in discard)
            {
                var entry = _entries[id]; entry.Retired = true;
                // 淘汰记录不阻断在途核验；由持有会话的调用在结束时释放取消源。
                if (!entry.Busy) entry.Cancellation.Dispose();
                _entries.Remove(id);
            }
            foreach (var client in _clients.Keys.Where(c => !_connected.Contains(c) && !_entries.Values.Any(e => e.Owner == c)).ToArray()) _clients.Remove(client);
        }
        private void EnsureRunning() { if (_stopped) throw new McpFault("plugin_stopping", "插件正在断开，停止接受请求。"); }
        private static void Cancel(Entry entry)
        {
            if (entry.Retired) return;
            try { entry.Cancellation.Cancel(); } catch (ObjectDisposedException) { }
        }
        public void Dispose()
        {
            lock (_gate)
            {
                if (_stopped) return;
                _stopped = true; _cleanup.Dispose();
                _shutdown.Cancel();
                foreach (var entry in _entries.Values) Cancel(entry);
            }
        }
    }
}
