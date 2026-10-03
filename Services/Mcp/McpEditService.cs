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
            ["begin_edit"] = AgentSchema.Obj(new Dictionary<string, AgentSchema> { ["scope"] = AgentSchema.Str("page", "selection") }, "scope"),
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
                ["begin_edit"] = "固定当前页面、选区和能力设置，创建编辑快照。scope=page 处理整页；selection 处理选中段落，没有选区时报错。返回 snapshot_id、可用工具及配额。",
                ["get_edit_status"] = "查询快照状态、提交结果和是否可以撤销。请求取消、断线或写回未确认后先查询；不要自动重发写入。",
                ["abort_edit"] = "取消未提交草稿。写入已经开始时继续核验实际结果，随后用 get_edit_status 查询。",
                ["undo_edit"] = "凭本客户端的快照撤销已核验的修改；仍使用页面冲突检查及整组撤销。重复调用返回原撤销结果，不能再次执行。"
            };
            foreach (var pair in Management) tools.Add(new McpToolDefinition { Name = pair.Key, Description = descriptions[pair.Key],
                InputSchemaJson = McpJson.Serialize(pair.Value.Json()), ReadOnly = pair.Key.StartsWith("get_", StringComparison.Ordinal), Destructive = pair.Key == "undo_edit" });
            return new McpCatalog { Tools = tools.ToArray(), Instructions =
                "先 get_current_page，再 begin_edit 固定范围；后续工具携带该 snapshot_id。目录包含全部工具，实际可用工具以 begin_edit 返回的 available_tools 为准。" +
                AgentRunner.SystemPrompt(agent) +
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

        internal string Call(string client, string name, string arguments, CancellationToken cancellation)
        {
            EnsureRunning();
            lock (_gate) if (!_clients.ContainsKey(client)) throw new McpFault("invalid_client", "客户端未连接。");
            cancellation.ThrowIfCancellationRequested();
            var args = McpJson.Arguments(arguments);
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
                    case "begin_edit": return Begin(client, (string)args["scope"], cancellation);
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
                entry.Session.Tools.ValidateArguments(call);
                if (name == "finish_edit" && entry.ResultJson != null)
                {
                    if (Convert.ToInt32(args["draft_revision"]) != entry.SubmittedRevision) throw new McpFault("revision_conflict", "快照已提交，提交修订号不一致。");
                    return entry.ResultJson;
                }
                if (entry.CompletedTicks != 0 || entry.Session.Snapshot.Frozen) throw new McpFault("edit_frozen", "草稿已结束或冻结，请查询状态。");
                using (cancellation.Register(() => Cancel(entry)))
                {
                    entry.Cancellation.Token.ThrowIfCancellationRequested();
                    if (entry.Calls >= entry.Session.Snapshot.Options.MaxToolCalls) { Cancel(entry); throw new McpFault("tool_limit", "工具调用达到上限，未提交草稿。"); }
                    Interlocked.Increment(ref entry.Calls);
                    lock (_gate) { entry.Busy = true; entry.State = name == "finish_edit" ? "Committing" : "Draft"; }
                    try
                    {
                        if (name == "finish_edit") entry.SubmittedRevision = Convert.ToInt32(args["draft_revision"]);
                        var outcome = entry.Session.Execute(call);
                        var result = McpJson.Serialize(outcome);
                        entry.Revision = entry.Session.Snapshot.Revision;
                        if (entry.Session.Snapshot.Frozen)
                        {
                            entry.ResultJson = result; entry.CanUndo = entry.Session.Report?.CanUndo == true;
                            Complete(entry, entry.Session.Report?.Status ?? "CommitOutcomeUnknown");
                        }
                        else entry.State = "Draft";
                        return result;
                    }
                    catch (OperationCanceledException)
                    {
                        Complete(entry, "CancelledBeforeCommit");
                        throw;
                    }
                    catch (Exception) when (entry.Session.Snapshot.Frozen)
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

        private string Begin(string client, string scope, CancellationToken cancellation)
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
                var page = ReadCurrentPage();
                var selection = scope == "selection" ? AgentPageSnapshot.SelectedIds(page) : null;
                if (selection != null && selection.Count == 0) throw new McpFault("selection_empty", "没有选中段落，请在 OneNote 中选择段落后重试。");
                var options = AgentOptions.Parse(new XElement("Agent", _options().ToElements()));
                var snapshot = new AgentPageSnapshot(page.ToString(), selection, options);
                linked.Token.ThrowIfCancellationRequested();
                var source = new CancellationTokenSource(TimeSpan.FromSeconds(options.TimeoutSeconds));
                var entry = new Entry { Owner = client, Cancellation = source, Expires = _now().AddSeconds(options.TimeoutSeconds),
                    Session = new AgentEditSession(snapshot, _committer, source.Token, _settings().Clone()) };
                lock (_gate)
                {
                    if (_stopped || !_connected.Contains(client) || linked.IsCancellationRequested)
                    { source.Dispose(); throw new OperationCanceledException(linked.Token); }
                    _entries.Add(snapshot.SnapshotId, entry);
                }
                return McpJson.Serialize(new { snapshot_id = snapshot.SnapshotId, instance_id = InstanceId, page_id = snapshot.PageId, page_title = snapshot.Title,
                    scope, draft_revision = 0, available_tools = Catalog.Tools.Where(t => entry.Session.Tools.Has(t.Name)).Select(t => t.Name).ToArray(),
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

        private string Status(string id, Entry entry) => McpJson.Serialize(new { snapshot_id = id, instance_id = InstanceId, page_id = entry.Session.Snapshot.PageId,
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
            using (var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(entry.Session.Snapshot.Options.TimeoutSeconds)))
            using (var linked = CancellationTokenSource.CreateLinkedTokenSource(timeout.Token, cancellation))
            {
                try
                {
                    var report = _committer.Undo(entry.Session.Snapshot.PageId, entry.Session.Report, entry.Session.Snapshot.Options, linked.Token);
                    entry.UndoJson = McpJson.Serialize(report.ToToolResult());
                    entry.State = "Undo" + report.Status;
                }
                catch (OperationCanceledException)
                {
                    // Committer 只会在开始写入前抛出取消；以后仍完成回读核验。
                    entry.State = entry.Session.Report.Status;
                    throw;
                }
                catch (Exception)
                {
                    entry.UndoJson = McpJson.Serialize(new { status = "CommitOutcomeUnknown", message = "撤销写回未确认，禁止自动重发。" });
                    entry.State = "UndoOutcomeUnknown";
                }
                finally
                {
                    if (entry.UndoJson != null) { entry.Session.Report.ClearUndo(); entry.CanUndo = false; }
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
            entry.Session.Snapshot.Frozen = true;
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
