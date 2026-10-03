using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using OneNoteCodeHelper.Mcp;
using OneNoteCodeHelper.Services;
using OneNoteCodeHelper.Services.Agent;
using OneNoteCodeHelper.Services.Mcp;

internal static partial class Program
{
    private const string McpClientA = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
    private const string McpSecretA = "11111111111111111111111111111111";
    private static McpEditService McpService(FakePage api, Func<AgentOptions> options = null, Func<DateTime> now = null)
    {
        var service = new McpEditService(api, () => "page", options ?? (() => new AgentOptions()), () => new AddInSettings(), now);
        service.BindClient(McpClientA, McpSecretA); return service;
    }
    private static IDictionary<string, object> Map(object value) => (IDictionary<string, object>)value;
    private static IDictionary<string, object> McpCall(McpEditService s, string tool, object args, string client = McpClientA) =>
        McpJson.Arguments(s.Call(client, tool, McpJson.Serialize(args), CancellationToken.None));
    private static string McpBegin(McpEditService s, string scope = "page") => (string)McpCall(s, "begin_edit", new { scope })["snapshot_id"];
    private static void McpRead(McpEditService s, string id, params string[] blocks) => McpCall(s, "read_blocks", new { snapshot_id = id, block_ids = blocks });
    private static void McpStyle(McpEditService s, string id) => McpCall(s, "set_paragraph_style", new { snapshot_id = id, block_ids = new[] { "p1" }, preset_id = "heading1" });
    private static int McpRevision(McpEditService s, string id) => (int)McpCall(s, "get_pending_changes", new { snapshot_id = id })["draft_revision"];
    private static IDictionary<string, object> McpFinish(McpEditService s, string id, int revision) => McpCall(s, "finish_edit", new { snapshot_id = id, draft_revision = revision });

    private static int TestMcp(string executable)
    {
        Test("MCP catalog maps every canonical Agent tool and overview snapshot schema", () =>
        {
            using (var s = McpService(new FakePage(Page(Paragraph("a", "中文😀")))))
            {
                var canonical = AgentTools.Catalog().Definitions.Select(d => Map(McpJson.Arguments(McpJson.Serialize(d))["function"])).ToArray();
                Equal(canonical.Length + 6 + McpExtensionCatalog.Schemas.Count, s.Catalog.Tools.Length);
                Equal(s.Catalog.Tools.Length, s.Catalog.Tools.Select(t => t.Name).Distinct().Count());
                foreach (var d in canonical)
                {
                    var tool = s.Catalog.Tools.Single(t => t.Name == (string)d["name"]);
                    Equal((string)d["description"], tool.Description);
                    var schema = McpJson.Arguments(tool.InputSchemaJson);
                    if (tool.Name == "get_page_overview") True(((IList)schema["required"]).Contains("snapshot_id"));
                    else Equal(McpJson.Serialize(d["parameters"]), tool.InputSchemaJson);
                }
                var id = McpBegin(s);
                Equal("page", McpCall(s, "get_page_overview", new { snapshot_id = id })["scope"]);
                Throws<McpFault>(() => McpCall(s, "get_page_overview", new { }));
                Throws(() => McpCall(s, "get_page_overview", new { snapshot_id = id, stray = true }));
            }
        });
        Test("MCP fixes selection and settings, no API key or window needed", () =>
        {
            var a = Paragraph("a", "选中文字"); a.SetAttributeValue("selected", "all");
            var api = new FakePage(Page(a, Paragraph("b", "范围外")));
            var options = new AgentOptions { EnableLists = false, EnableMoves = false };
            using (var s = McpService(api, () => options))
            {
                var id = McpBegin(s, "selection");
                options.EnableLists = options.EnableMoves = true;
                api.Page.Descendants().Attributes("selected").Remove();
                api.Page.SetAttributeValue("ID", "another-page");
                var overview = McpCall(s, "get_page_overview", new { snapshot_id = id });
                Equal(1, overview["total"]); Equal(false, overview["list_edit"]); Equal(false, overview["move"]);
                Equal("page", McpCall(s, "get_edit_status", new { snapshot_id = id })["page_id"]);
                Throws(() => McpCall(s, "set_list", new { snapshot_id = id, block_ids = new[] { "p1" }, list = "bullet" }));
                Throws(() => McpCall(s, "set_page_title", new { snapshot_id = id, title = "标题" }));
                api.Page.SetAttributeValue("ID", "page");
                Throws<McpFault>(() => McpBegin(s, "selection"));
                Equal(0, api.Writes);
            }
        });
        Test("MCP refuses unread and stale revisions, commits once, freezes, undoes once", () =>
        {
            var api = new FakePage(Page(Paragraph("a", "第一段"), Paragraph("b", "第二段")));
            using (var s = McpService(api))
            {
                var id = McpBegin(s); Equal(false, McpFinish(s, id, 0)["ok"]); Equal(0, api.Writes);
                McpRead(s, id, "p1", "p2"); McpStyle(s, id); var revision = McpRevision(s, id);
                Throws(() => McpFinish(s, id, revision - 1));
                Equal("Draft", McpCall(s, "get_edit_status", new { snapshot_id = id })["state"]);
                var result = McpFinish(s, id, revision); Equal("Verified", result["status"]); Equal(1, api.Writes);
                Equal(McpJson.Serialize(result), McpJson.Serialize(McpFinish(s, id, revision))); Equal(1, api.Writes);
                Throws<McpFault>(() => McpFinish(s, id, revision + 1)); Throws<McpFault>(() => McpStyle(s, id));
                Equal(true, McpCall(s, "get_edit_status", new { snapshot_id = id })["can_undo"]);
                var undo = McpCall(s, "undo_edit", new { snapshot_id = id }); Equal("Verified", undo["status"]); Equal(2, api.Writes);
                Equal(McpJson.Serialize(undo), McpJson.Serialize(McpCall(s, "undo_edit", new { snapshot_id = id }))); Equal(2, api.Writes);
                Equal(false, McpCall(s, "get_edit_status", new { snapshot_id = id })["can_undo"]);
            }
        });
        Test("MCP compact XML commits and undoes formatting without changing content or runs", () =>
        {
            var parent = Paragraph("parent", "父段落");
            parent.Add(new XElement(One + "OEChildren", Paragraph("child", "子段落")));
            var api = new FakePage(Page(Paragraph("a", "&nbsp;&nbsp;中文😀 <b>重点</b><br><a href='https://example.com/?a=1&amp;b=2'>链接</a>", "\t  "),
                Paragraph("blank", ""), parent, Paragraph("b", "最后一段"))) { ReadSaveOptions = SaveOptions.DisableFormatting };
            var before = string.Join("|", api.Page.Descendants(One + "OE").Select(e => new AgentRichText(e).Signature(api.Page, false)));
            var childFormat = AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "child"), api.Page);
            using (var s = McpService(api))
            {
                var id = McpBegin(s); McpRead(s, id, "p1", "p3", "p4", "p5"); McpStyle(s, id);
                McpCall(s, "set_paragraph_style", new { snapshot_id = id, block_ids = new[] { "p5" }, preset_id = "body" });
                var revision = McpRevision(s, id); Equal(0, api.Writes);
                var result = McpFinish(s, id, revision);
                Equal("Verified", result["status"]); Equal(2, result["applied"]); Equal(0, ((IList)result["skipped_conflict"]).Count); Equal(1, api.Writes);
                Equal(before, string.Join("|", api.Page.Descendants(One + "OE").Select(e => new AgentRichText(e).Signature(api.Page, false))));
                Equal(2, AgentCommitter.Find(api.Page, "a").Elements(One + "T").Count());
                var lastRun = new XElement(One + "OE", new XElement(AgentCommitter.Find(api.Page, "a").Elements(One + "T").Last()));
                Equal("\t  ", new AgentRichText(lastRun).Text);
                Equal(childFormat, AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "child"), api.Page));
                Equal("Verified", McpFinish(s, id, revision)["status"]); Equal(1, api.Writes);
                Equal("Verified", McpCall(s, "undo_edit", new { snapshot_id = id })["status"]); Equal(2, api.Writes);
                Equal(before, string.Join("|", api.Page.Descendants(One + "OE").Select(e => new AgentRichText(e).Signature(api.Page, false))));
            }
        });
        Test("MCP compact XML selection preserves unselected nested and outside paragraphs", () =>
        {
            var parent = Paragraph("a", "父段落"); parent.SetAttributeValue("selected", "all");
            parent.Add(new XElement(One + "OEChildren", Paragraph("child", "选区外的子段落")));
            var api = new FakePage(Page(parent, Paragraph("outside", "选区外正文"))) { ReadSaveOptions = SaveOptions.DisableFormatting };
            var childFormat = AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "child"), api.Page);
            var outsideFormat = AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "outside"), api.Page);
            var before = Texts(api.Page);
            using (var s = McpService(api))
            {
                var id = McpBegin(s, "selection");
                Equal(1, McpCall(s, "get_page_overview", new { snapshot_id = id })["total"]);
                McpRead(s, id, "p1"); McpStyle(s, id);
                api.Page.Descendants().Attributes("selected").Remove();
                Equal("Verified", McpFinish(s, id, McpRevision(s, id))["status"]); Equal(1, api.Writes);
                Equal(before, Texts(api.Page));
                Equal(childFormat, AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "child"), api.Page));
                Equal(outsideFormat, AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "outside"), api.Page));
            }
        });
        Test("MCP compact XML preserves real text, whitespace, link and format conflicts", () =>
        {
            foreach (var edit in new Action<XElement>[] {
                e => e.Element(One + "T").Value = "用户修改后的正文",
                e => e.Element(One + "T").Value = "  第一段  ",
                e => e.Element(One + "T").Value = "<a href='https://example.org/changed'>第一段</a>",
                e => e.SetAttributeValue("style", "font-size:22pt") })
            {
                var api = new FakePage(Page(Paragraph("a", "第一段"), Paragraph("b", "第二段"))) { ReadSaveOptions = SaveOptions.DisableFormatting };
                using (var s = McpService(api))
                {
                    var id = McpBegin(s); McpRead(s, id, "p1", "p2"); McpStyle(s, id);
                    edit(AgentCommitter.Find(api.Page, "a")); var changed = api.Page.ToString(SaveOptions.DisableFormatting);
                    var result = McpFinish(s, id, McpRevision(s, id));
                    Equal("NoChange", result["status"]); Equal(0, result["applied"]); Equal(0, api.Writes);
                    Equal("p1", string.Join(",", ((IList)result["skipped_conflict"]).Cast<string>()));
                    Equal(changed, api.Page.ToString(SaveOptions.DisableFormatting));
                }
            }
        });
        Test("MCP compact XML merges and undoes whole groups, skips concurrent group changes", () =>
        {
            foreach (var conflict in new[] { false, true })
            {
                var api = new FakePage(TwoBoxes()) { ReadSaveOptions = SaveOptions.DisableFormatting };
                var before = Texts(api.Page);
                using (var s = McpService(api))
                {
                    var id = McpBegin(s); McpRead(s, id, "p1", "p2", "p3", "p4", "p5");
                    McpCall(s, "merge_outlines", new { snapshot_id = id, source_id = "B", target_id = "p2", position = "after" });
                    if (conflict) AgentCommitter.Find(api.Page, "a1").Element(One + "T").Value = "用户修改目标框";
                    var result = McpFinish(s, id, McpRevision(s, id));
                    if (conflict)
                    {
                        Equal("NoChange", result["status"]); Equal(0, api.Writes);
                        Equal(2, api.Page.Elements(One + "Outline").Count()); True(Texts(api.Page).Contains("用户修改目标框"));
                    }
                    else
                    {
                        Equal("Verified", result["status"]); Equal(1, api.Writes);
                        Equal(1, api.Page.Elements(One + "Outline").Count()); Equal(before, Texts(api.Page));
                        Equal("Verified", McpCall(s, "undo_edit", new { snapshot_id = id })["status"]); Equal(2, api.Writes);
                        Equal(2, api.Page.Elements(One + "Outline").Count()); Equal(before, Texts(api.Page));
                    }
                }
            }
        });
        Test("MCP concurrent clients own separate drafts and cannot claim snapshots", () =>
        {
            using (var s = McpService(new FakePage(Page(Paragraph("a", "正文")))))
            {
                var b = Guid.NewGuid().ToString("N"); s.BindClient(b, Guid.NewGuid().ToString("N"));
                var id = McpBegin(s);
                var second = (string)McpCall(s, "begin_edit", new { scope = "page" }, b)["snapshot_id"];
                True(id != second);
                Throws<McpFault>(() => McpCall(s, "get_edit_status", new { snapshot_id = id }, b));
                Throws<McpFault>(() => s.BindClient(McpClientA, Guid.NewGuid().ToString("N")));
                Parallel.Invoke(() => McpRead(s, id, "p1"), () => McpCall(s, "read_blocks", new { snapshot_id = second, block_ids = new[] { "p1" } }, b));
                McpStyle(s, id); Equal(0, McpCall(s, "get_pending_changes", new { snapshot_id = second }, b)["draft_revision"]);
            }
        });
        Test("MCP preserves concurrent user edits and refuses group undo after conflict", () =>
        {
            var api = new FakePage(TwoBoxes());
            using (var s = McpService(api))
            {
                var original = Texts(api.Page); var id = McpBegin(s); McpRead(s, id, "p1", "p2", "p3", "p4", "p5");
                McpCall(s, "merge_outlines", new { snapshot_id = id, source_id = "B", target_id = "p2", position = "after" });
                Equal("Verified", McpFinish(s, id, McpRevision(s, id))["status"]);
                Equal("Verified", McpCall(s, "undo_edit", new { snapshot_id = id })["status"]); Equal(original, Texts(api.Page));
                id = McpBegin(s); McpRead(s, id, "p1", "p2", "p3", "p4", "p5");
                var outlines = (IList)McpCall(s, "get_page_overview", new { snapshot_id = id })["outlines"];
                var restoredSource = (string)Map(outlines[1])["container_id"];
                McpCall(s, "merge_outlines", new { snapshot_id = id, source_id = restoredSource, target_id = "p2", position = "after" });
                Equal("Verified", McpFinish(s, id, McpRevision(s, id))["status"]);
                AgentCommitter.Find(api.Page, "a1").Element(One + "T").Value = "用户保留的整组改动";
                var writes = api.Writes;
                True(((IList)McpCall(s, "undo_edit", new { snapshot_id = id })["skipped_conflict"]).Count > 0);
                Equal(writes, api.Writes); Equal(1, api.Page.Elements(One + "Outline").Count());
                True(Texts(api.Page).Contains("用户保留的整组改动"));
                id = McpBegin(s); McpRead(s, id, "p1", "p2", "p3", "p4", "p5"); McpStyle(s, id);
                AgentCommitter.Find(api.Page, "a1").Element(One + "T").Value = "用户同时改动";
                Equal("NoChange", McpFinish(s, id, McpRevision(s, id))["status"]);
                True(Texts(api.Page).Contains("用户同时改动"));
            }
        });
        Test("MCP uncertain write is retained, not retried, cancel after save still verifies", () =>
        {
            var api = new FakePage(Page(Paragraph("a", "中文")));
            using (var s = McpService(api))
            {
                var id = McpBegin(s); McpRead(s, id, "p1"); McpStyle(s, id); var revision = McpRevision(s, id);
                api.FailReadAfterSave = true;
                Equal("CommitOutcomeUnknown", McpFinish(s, id, revision)["status"]);
                Equal("CommitOutcomeUnknown", McpFinish(s, id, revision)["status"]); Equal(1, api.Writes);
            }
            api = new FakePage(Page(Paragraph("a", "中文")));
            using (var s = McpService(api))
            {
                var id = McpBegin(s); McpRead(s, id, "p1"); McpStyle(s, id);
                api.AfterSave = () => McpCall(s, "abort_edit", new { snapshot_id = id });
                Equal("Verified", McpFinish(s, id, McpRevision(s, id))["status"]); Equal(1, api.Writes);
            }
        });
        Test("MCP disconnect cancels drafts, completed result survives resume, records expire", () =>
        {
            var now = DateTime.UtcNow; var api = new FakePage(Page(Paragraph("a", "正文")));
            using (var s = McpService(api, now: () => now))
            {
                var id = McpBegin(s); McpRead(s, id, "p1"); McpStyle(s, id); var revision = McpRevision(s, id); McpFinish(s, id, revision);
                var draft = McpBegin(s); s.DisconnectClient(McpClientA); s.BindClient(McpClientA, McpSecretA);
                Equal("CancelledBeforeCommit", McpCall(s, "get_edit_status", new { snapshot_id = draft })["state"]);
                Equal("Verified", McpFinish(s, id, revision)["status"]); Equal(1, api.Writes);
                now = now.AddMinutes(31); Throws<McpFault>(() => McpCall(s, "get_edit_status", new { snapshot_id = id }));
            }
        });
        Test("MCP active limit, tool quota, draft deadline and oldest-result eviction", () =>
        {
            var now = DateTime.UtcNow;
            using (var s = McpService(new FakePage(Page(Paragraph("a", "正文"))), () => new AgentOptions { MaxToolCalls = 6 }, () => now))
            {
                var ids = Enumerable.Range(0, 8).Select(_ => McpBegin(s)).ToArray(); Throws<McpFault>(() => McpBegin(s));
                for (var i = 0; i < 6; i++) McpCall(s, "get_page_overview", new { snapshot_id = ids[0] });
                Throws<McpFault>(() => McpCall(s, "get_page_overview", new { snapshot_id = ids[0] }));
                now = now.AddMinutes(11); Equal("CancelledBeforeCommit", McpCall(s, "get_edit_status", new { snapshot_id = ids[1] })["state"]);
                for (var i = 0; i < 17; i++) { var id = McpBegin(s); McpCall(s, "abort_edit", new { snapshot_id = id }); now = now.AddSeconds(1); }
                Throws<McpFault>(() => McpCall(s, "get_edit_status", new { snapshot_id = ids[0] }));
                Equal(16, McpCall(s, "get_status", new { })["retained_results"]);
            }
        });
        Test("IPC UTF-8 framing, invalid length and truncated data", () =>
        {
            using (var stream = new MemoryStream())
            {
                McpProtocol.WriteAsync(stream, "中文😀", CancellationToken.None).GetAwaiter().GetResult(); stream.Position = 0;
                Equal("中文😀", McpProtocol.ReadAsync(stream, CancellationToken.None).GetAwaiter().GetResult());
            }
            using (var stream = new MemoryStream(new byte[] { 0, 0, 0, 0 })) Throws<InvalidDataException>(() => McpProtocol.ReadAsync(stream, CancellationToken.None).GetAwaiter().GetResult());
            using (var stream = new MemoryStream(new byte[] { 3, 0, 0, 0, 1 })) Throws<EndOfStreamException>(() => McpProtocol.ReadAsync(stream, CancellationToken.None).GetAwaiter().GetResult());
        });
        Test("MCP selection includes blank lines between selected paragraphs like the Agent window", () =>
        {
            var a = Paragraph("a", "第一段"); a.SetAttributeValue("selected", "all");
            var c = Paragraph("c", "第三段"); c.SetAttributeValue("selected", "all");
            using (var s = McpService(new FakePage(Page(a, Paragraph("b", ""), c, Paragraph("d", "范围外")))))
            {
                var overview = McpCall(s, "get_page_overview", new { snapshot_id = McpBegin(s, "selection") });
                Equal(3, overview["total"]);
                True(((IList)overview["blocks"]).Cast<object>().Any(b => (string)Map(b)["reason"] == "empty"));
            }
        });
        Test("pipe host keeps serving past 16 connections, waits at the limit and recovers", () =>
        {
            using (var s = McpService(new FakePage(Page(Paragraph("a", "正文")))))
            using (var host = new McpPipeHost(s, "OneNoteCodeHelper.Test." + Guid.NewGuid().ToString("N")))
            {
                host.Start(); True(host.Ready.Wait(5000));
                var name = PipeName(host);
                var clients = new List<NamedPipeClientStream>();
                try
                {
                    for (var i = 0; i < McpPipeHost.MaxConnections; i++) { var pipe = RawPipe(name); clients.Add(pipe); Equal(true, RawHello(pipe).Success); }
                    Equal(true, McpCall(s, "get_status", new { })["connected"]);
                    // 满额时不建新实例：新客户端等空位，服务本身照常。实例全忙时 net48 客户端报 IOException（ERROR_SEM_TIMEOUT），.NET 10 报 TimeoutException。
                    try { RawPipe(name, 500).Dispose(); throw new Exception("Expected the full pipe host to refuse a new connection"); }
                    catch (TimeoutException) { }
                    catch (IOException) { }
                    clients[0].Dispose(); clients.RemoveAt(0);
                    using (var waiting = RawPipe(name)) Equal(true, RawHello(waiting).Success);
                }
                finally { foreach (var pipe in clients) pipe.Dispose(); }
                using (var fresh = RawPipe(name)) Equal(true, RawHello(fresh).Success);
                Equal(true, McpCall(s, "get_status", new { })["connected"]);
            }
        });
        Test("pipe host drops a connection that never says hello", () =>
        {
            using (var s = McpService(new FakePage(Page(Paragraph("a", "正文")))))
            using (var host = new McpPipeHost(s, "OneNoteCodeHelper.Test." + Guid.NewGuid().ToString("N")))
            {
                host.Start(); True(host.Ready.Wait(5000));
                using (var silent = RawPipe(PipeName(host)))
                {
                    var watch = Stopwatch.StartNew();
                    var read = silent.ReadAsync(new byte[1], 0, 1);
                    True(read.Wait(McpProtocol.ConnectTimeoutMs + 3000)); Equal(0, read.Result);
                    True(watch.ElapsedMilliseconds >= McpProtocol.ConnectTimeoutMs - 500);
                }
                using (var pipe = RawPipe(PipeName(host))) Equal(true, RawHello(pipe).Success);
            }
        });
        if (string.IsNullOrEmpty(executable) || !File.Exists(executable)) throw new Exception("MCP EXE 不存在，请先构建解决方案或传入 --mcp-only <exe>。");
        Test("real MCP subprocess initialize, discover all tools, Chinese, errors, edit, repeat, undo and EOF", () =>
        {
            var api = new FakePage(Page(Paragraph("a", "中文😀第一段"), Paragraph("b", "第二段"))) { ReadSaveOptions = SaveOptions.DisableFormatting };
            using (var s = McpService(api))
            using (var host = new McpPipeHost(s, "OneNoteCodeHelper.Test." + Guid.NewGuid().ToString("N")))
            {
                host.Start(); True(host.Ready.Wait(5000));
                using (var child = new McpChild(executable, PipeName(host)))
                {
                    child.Initialize(); var list = child.Rpc("tools/list", new { }); Equal(s.Catalog.Tools.Length, ((IList)list["tools"]).Count);
                    var id = (string)child.Call("begin_edit", new { scope = "page" })["snapshot_id"];
                    var refused = child.ToolResult("finish_edit", new { snapshot_id = id, draft_revision = 0 }); Equal(true, refused["isError"]);
                    var read = child.Call("read_blocks", new { snapshot_id = id, block_ids = new[] { "p1", "p2" } }); True(McpJson.Serialize(read).Contains("中文"));
                    child.Call("set_paragraph_style", new { snapshot_id = id, block_ids = new[] { "p1" }, preset_id = "heading1" });
                    child.Call("set_page_title", new { snapshot_id = id, title = "中文标题😀" });
                    var revision = (int)child.Call("get_pending_changes", new { snapshot_id = id })["draft_revision"];
                    Equal("Verified", child.Call("finish_edit", new { snapshot_id = id, draft_revision = revision })["status"]);
                    Equal("中文标题😀", AgentPageTitle.Text(api.Page.Element(One + "Title")));
                    Equal("Verified", child.Call("finish_edit", new { snapshot_id = id, draft_revision = revision })["status"]); Equal(1, api.Writes);
                    Equal(true, child.ToolResult("get_edit_status", new { snapshot_id = "missing" })["isError"]);
                    Equal("Verified", child.Call("undo_edit", new { snapshot_id = id })["status"]);
                    child.Call("undo_edit", new { snapshot_id = id }); Equal(2, api.Writes);
                    var draft = (string)child.Call("begin_edit", new { scope = "page" })["snapshot_id"];
                    child.CloseInput(); Equal(0, child.ExitCode());
                    True(!child.Stderr.Contains("中文😀第一段"));
                    // 真实子进程退出后取消未提交草稿（服务层保留其状态）。
                    True(SpinWait.SpinUntil(() => (int)McpCall(s, "get_status", new { })["active_edits"] == 0, 5000));
                }
            }
        });
        Test("real MCP cancellation after write verifies and allows status query", () =>
        {
            var api = new FakePage(Page(Paragraph("a", "中文")));
            using (var s = McpService(api))
            using (var saved = new ManualResetEventSlim())
            using (var release = new ManualResetEventSlim())
            using (var host = new McpPipeHost(s, "OneNoteCodeHelper.Test." + Guid.NewGuid().ToString("N")))
            {
                host.Start(); True(host.Ready.Wait(5000));
                using (var child = new McpChild(executable, PipeName(host)))
                {
                    child.Initialize(); var id = (string)child.Call("begin_edit", new { scope = "page" })["snapshot_id"];
                    child.Call("read_blocks", new { snapshot_id = id, block_ids = new[] { "p1" } });
                    child.Call("set_paragraph_style", new { snapshot_id = id, block_ids = new[] { "p1" }, preset_id = "heading1" });
                    var revision = (int)child.Call("get_pending_changes", new { snapshot_id = id })["draft_revision"];
                    api.AfterSave = () => { saved.Set(); True(release.Wait(5000)); };
                    var request = child.Send("tools/call", new { name = "finish_edit", arguments = new { snapshot_id = id, draft_revision = revision } });
                    try
                    {
                        True(saved.Wait(5000)); child.Notify("notifications/cancelled", new { requestId = request, reason = "test cancellation" });
                        True(SpinWait.SpinUntil(() => (bool)child.Call("get_edit_status", new { snapshot_id = id })["cancellation_requested"], 3000));
                    }
                    finally { release.Set(); }
                    True(SpinWait.SpinUntil(() => (string)child.Call("get_edit_status", new { snapshot_id = id })["state"] == "Verified", 5000));
                    Equal("Verified", child.Call("finish_edit", new { snapshot_id = id, draft_revision = revision })["status"]); Equal(1, api.Writes);
                }
            }
        });
        Test("real MCP clients isolate sessions and serialize concurrent page commits", () =>
        {
            var api = new FakePage(Page(Paragraph("a", "甲"), Paragraph("b", "乙")));
            using (var s = McpService(api))
            using (var host = new McpPipeHost(s, "OneNoteCodeHelper.Test." + Guid.NewGuid().ToString("N")))
            {
                host.Start(); True(host.Ready.Wait(5000));
                using (var a = new McpChild(executable, PipeName(host)))
                using (var b = new McpChild(executable, PipeName(host)))
                {
                    a.Initialize(); b.Initialize();
                    var first = (string)a.Call("begin_edit", new { scope = "page" })["snapshot_id"];
                    var second = (string)b.Call("begin_edit", new { scope = "page" })["snapshot_id"];
                    Equal(true, b.ToolResult("get_edit_status", new { snapshot_id = first })["isError"]);
                    a.Call("read_blocks", new { snapshot_id = first, block_ids = new[] { "p1", "p2" } });
                    b.Call("read_blocks", new { snapshot_id = second, block_ids = new[] { "p1", "p2" } });
                    a.Call("set_paragraph_style", new { snapshot_id = first, block_ids = new[] { "p1" }, preset_id = "heading1" });
                    b.Call("set_paragraph_style", new { snapshot_id = second, block_ids = new[] { "p2" }, preset_id = "heading2" });
                    Parallel.Invoke(
                        () => Equal("Verified", a.Call("finish_edit", new { snapshot_id = first, draft_revision = 1 })["status"]),
                        () => Equal("Verified", b.Call("finish_edit", new { snapshot_id = second, draft_revision = 1 })["status"]));
                    Equal(2, api.Writes); Equal("甲|乙", Texts(api.Page));
                }
            }
        });
        Test("real MCP transport reconnect preserves commit result and cancels old draft", () =>
        {
            var api = new FakePage(Page(Paragraph("a", "连接恢复")));
            using (var s = McpService(api))
            using (var host = new McpPipeHost(s, "OneNoteCodeHelper.Test." + Guid.NewGuid().ToString("N")))
            {
                host.Start(); True(host.Ready.Wait(5000));
                using (var child = new McpChild(executable, PipeName(host)))
                {
                    child.Initialize(); var id = (string)child.Call("begin_edit", new { scope = "page" })["snapshot_id"];
                    child.Call("read_blocks", new { snapshot_id = id, block_ids = new[] { "p1" } });
                    child.Call("set_paragraph_style", new { snapshot_id = id, block_ids = new[] { "p1" }, preset_id = "heading1" });
                    child.Call("finish_edit", new { snapshot_id = id, draft_revision = 1 });
                    var draft = (string)child.Call("begin_edit", new { scope = "page" })["snapshot_id"];
                    // 注入一次传输故障，服务和实例保持不变，监听管道保留。
                    var pipes = (ConcurrentDictionary<NamedPipeServerStream, byte>)typeof(McpPipeHost)
                        .GetField("_pipes", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(host);
                    foreach (var pipe in pipes.Keys) if (pipe.IsConnected) pipe.Dispose();
                    True(SpinWait.SpinUntil(() => !(bool)child.ToolResult("get_status", new { })["isError"], 5000));
                    Equal("CancelledBeforeCommit", child.Call("get_edit_status", new { snapshot_id = draft })["state"]);
                    Equal("Verified", child.Call("finish_edit", new { snapshot_id = id, draft_revision = 1 })["status"]); Equal(1, api.Writes);
                }
            }
        });
        Test("real MCP cancellation before write never commits", () =>
        {
            var api = new FakePage(Page(Paragraph("a", "写入前取消")));
            using (var s = McpService(api))
            using (var host = new McpPipeHost(s, "OneNoteCodeHelper.Test." + Guid.NewGuid().ToString("N")))
            {
                host.Start(); True(host.Ready.Wait(5000));
                using (var child = new McpChild(executable, PipeName(host)))
                {
                    child.Initialize(); var id = (string)child.Call("begin_edit", new { scope = "page" })["snapshot_id"];
                    child.Call("read_blocks", new { snapshot_id = id, block_ids = new[] { "p1" } });
                    child.Call("set_paragraph_style", new { snapshot_id = id, block_ids = new[] { "p1" }, preset_id = "heading1" });
                    lock (PageEditCoordinator.ForPage("page"))
                    {
                        var request = child.Send("tools/call", new { name = "finish_edit", arguments = new { snapshot_id = id, draft_revision = 1 } });
                        True(SpinWait.SpinUntil(() => (bool)child.Call("get_edit_status", new { snapshot_id = id })["busy"], 3000));
                        child.Notify("notifications/cancelled", new { requestId = request, reason = "cancel before write" });
                        True(SpinWait.SpinUntil(() => (bool)child.Call("get_edit_status", new { snapshot_id = id })["cancellation_requested"], 3000));
                    }
                    True(SpinWait.SpinUntil(() => (string)child.Call("get_edit_status", new { snapshot_id = id })["state"] == "CancelledBeforeCommit", 3000));
                    Equal(0, api.Writes);
                }
            }
        });
        Test("real MCP reconnect rejects old-instance snapshots, permits fresh drafts", () =>
        {
            var name = "OneNoteCodeHelper.Test." + Guid.NewGuid().ToString("N");
            var api = new FakePage(Page(Paragraph("a", "重连"))); var first = McpService(api); var host = new McpPipeHost(first, name);
            host.Start(); True(host.Ready.Wait(5000));
            using (var child = new McpChild(executable, name))
            {
                try
                {
                    child.Initialize(); var id = (string)child.Call("begin_edit", new { scope = "page" })["snapshot_id"];
                    host.Dispose();
                    using (var next = McpService(api))
                    using (var replacement = new McpPipeHost(next, name))
                    {
                        replacement.Start(); True(replacement.Ready.Wait(5000));
                        // 先等原连接读循环观察 EOF；一次失败只返回错误，没有重发写入。
                        True(SpinWait.SpinUntil(() => !(bool)child.ToolResult("get_status", new { })["isError"], 5000));
                        var old = child.ToolResult("get_edit_status", new { snapshot_id = id }); Equal(true, old["isError"]);
                        Equal("snapshot_expired", Map(old["structuredContent"])["error_code"]);
                        var fresh = (string)child.Call("begin_edit", new { scope = "page" })["snapshot_id"]; True(id != fresh);
                    }
                }
                finally { host.Dispose(); }
            }
        });
        Test("doctor, unavailable plugin exit, protocol mismatch, same-user ACL", () =>
        {
            var name = "OneNoteCodeHelper.Test." + Guid.NewGuid().ToString("N");
            using (var s = McpService(new FakePage(Page(Paragraph("a", "诊断")))))
            using (var host = new McpPipeHost(s, name))
            {
                host.Start(); True(host.Ready.Wait(5000));
                using (var child = new McpChild(executable, name, true)) { Equal(0, child.ExitCode()); True(child.Stdout().Contains("ipc_version")); }
                using (var pipe = new NamedPipeClientStream(".", name, PipeDirection.InOut, PipeOptions.Asynchronous))
                {
                    pipe.Connect(3000);
                    var acl = pipe.GetAccessControl().GetAccessRules(true, false, typeof(SecurityIdentifier));
                    True(acl.Cast<PipeAccessRule>().Any(r => r.IdentityReference.Equals(WindowsIdentity.GetCurrent().User)));
                    True(acl.Cast<PipeAccessRule>().Any(r => r.IdentityReference.Equals(new SecurityIdentifier(WellKnownSidType.NetworkSid, null)) && r.AccessControlType == System.Security.AccessControl.AccessControlType.Deny));
                    McpProtocol.WriteAsync(pipe, McpJson.Serialize(new McpRequest { ProtocolVersion = 999, Method = "hello", RequestId = "version" }), CancellationToken.None).GetAwaiter().GetResult();
                    var reply = McpJson.Deserialize<McpResponse>(McpProtocol.ReadAsync(pipe, CancellationToken.None).GetAwaiter().GetResult());
                    Equal("protocol_mismatch", reply.ErrorCode);
                }
            }
            using (var child = new McpChild(executable, name, true)) { Equal(1, child.ExitCode()); Equal("", child.Stdout()); True(child.Stderr.Contains("plugin_unavailable")); }
            using (var child = new McpChild(executable, name)) { Equal(1, child.ExitCode()); Equal("", child.Stdout()); }
        });
        TestMcpExtensions(executable);
        Console.WriteLine($"MCP/Agent tests: {_passed} passed, {_failed} failed");
        return _failed == 0 ? 0 : 1;
    }

    private static string PipeName(McpPipeHost host) => (string)typeof(McpPipeHost).GetField("_name", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(host);

    /// <summary>不经 MCP 程序、直接连插件管道的客户端。</summary>
    private static NamedPipeClientStream RawPipe(string name, int timeout = 3000)
    {
        var pipe = new NamedPipeClientStream(".", name, PipeDirection.InOut, PipeOptions.Asynchronous);
        try { pipe.Connect(timeout); return pipe; }
        catch { pipe.Dispose(); throw; }
    }

    private static McpResponse RawHello(NamedPipeClientStream pipe)
    {
        var hello = new McpRequest { Method = "hello", RequestId = Guid.NewGuid().ToString("N"), ClientId = Guid.NewGuid().ToString("N"), ResumeSecret = Guid.NewGuid().ToString("N") };
        McpProtocol.WriteAsync(pipe, McpJson.Serialize(hello), CancellationToken.None).GetAwaiter().GetResult();
        var reply = McpProtocol.ReadAsync(pipe, CancellationToken.None).GetAwaiter().GetResult();
        return reply == null ? null : McpJson.Deserialize<McpResponse>(reply);
    }

    private sealed class McpChild : IDisposable
    {
        private readonly Process _process;
        private readonly Task<string> _stderr;
        private readonly StreamWriter _input;
        private readonly Dictionary<int, IDictionary<string, object>> _replies = new Dictionary<int, IDictionary<string, object>>();
        internal readonly List<IDictionary<string, object>> Notifications = new List<IDictionary<string, object>>();
        private int _id;
        internal string Stderr => _stderr.GetAwaiter().GetResult();
        internal McpChild(string exe, string pipe, bool doctor = false)
        {
            _process = Process.Start(new ProcessStartInfo(exe, (doctor ? "--doctor " : "") + "--pipe-name " + pipe)
            { UseShellExecute = false, CreateNoWindow = true, RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true,
                StandardOutputEncoding = new UTF8Encoding(false), StandardErrorEncoding = new UTF8Encoding(false) });
            _stderr = _process.StandardError.ReadToEndAsync();
            _input = new StreamWriter(_process.StandardInput.BaseStream, new UTF8Encoding(false)) { AutoFlush = true };
        }
        internal void Initialize()
        {
            var result = Rpc("initialize", new { protocolVersion = "2025-11-25", capabilities = new { }, clientInfo = new { name = "OneNote offline tests", version = "1.0" } });
            Equal("2025-11-25", result["protocolVersion"]); Notify("notifications/initialized", new { });
        }
        internal int Send(string method, object args)
        {
            var id = ++_id;
            _input.WriteLine(McpJson.Serialize(new { jsonrpc = "2.0", id, method, @params = args })); return id;
        }
        internal void Notify(string method, object args) { _input.WriteLine(McpJson.Serialize(new { jsonrpc = "2.0", method, @params = args })); }
        internal IDictionary<string, object> Rpc(string method, object args)
        {
            var id = Send(method, args);
            while (!_replies.ContainsKey(id))
            {
                var line = _process.StandardOutput.ReadLineAsync();
                if (!line.Wait(8000)) throw new Exception("MCP stdout response timed out");
                if (line.Result == null) throw new Exception("MCP stdout EOF: " + Stderr);
                var message = McpJson.Arguments(line.Result); // 任意非协议 stdout 都会失败。
                Equal("2.0", message["jsonrpc"]);
                if (message.TryGetValue("id", out var responseId)) _replies[Convert.ToInt32(responseId)] = message;
                else Notifications.Add(message);
            }
            var reply = _replies[id]; _replies.Remove(id);
            if (reply.ContainsKey("error")) throw new Exception("MCP error: " + McpJson.Serialize(reply["error"]));
            return Map(reply["result"]);
        }
        internal IDictionary<string, object> ToolResult(string name, object arguments) => Rpc("tools/call", new { name, arguments });
        internal IDictionary<string, object> Call(string name, object arguments)
        {
            var result = ToolResult(name, arguments);
            if ((bool)result["isError"]) throw new Exception("MCP tool error: " + McpJson.Serialize(result));
            return Map(result["structuredContent"]);
        }
        internal void CloseInput() => _input.Close();
        internal string Stdout() => _process.StandardOutput.ReadToEnd();
        internal int ExitCode() { if (!_process.WaitForExit(6000)) throw new Exception("MCP did not exit"); return _process.ExitCode; }
        public void Dispose()
        {
            _input.Close();
            if (!_process.WaitForExit(5000)) { _process.Kill(); _process.WaitForExit(); }
            _process.Dispose();
        }
    }
}
