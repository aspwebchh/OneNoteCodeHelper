using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using OneNoteCodeHelper.Highlighting;
using OneNoteCodeHelper.Highlighting.Themes;
using OneNoteCodeHelper.Services;
using OneNoteCodeHelper.Services.Agent;

internal static class Program
{
    private static readonly XNamespace One = OneNoteApi.One;
    private static int _passed, _failed;
    private static readonly Dictionary<string, object> Empty = new Dictionary<string, object>();

    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length > 0 && args[0] == "--probe") return Probe.Run(args.Skip(1).ToArray());
        if (args.Length == 2 && args[0] == "--render-ui") return WindowPreview.Render(args[1]);
        if (args.Length == 2 && args[0] == "--inspect-fixture")
        {
            var page = AgentPageSnapshot.ParsePage(System.IO.File.ReadAllText(args[1]));
            foreach (var oe in page.Descendants(One + "OE").Where(e => e.Elements(One + "T").Any()))
            {
                try { Console.WriteLine(new AgentRichText(oe).Text); }
                catch (Exception ex) { Console.WriteLine(ex); }
            }
            return 0;
        }
        Test("HTML entities, whitespace, links and multiple runs survive formatting", () =>
        {
            var oe = Paragraph("a", "&nbsp;A&amp;B <b>重点</b><br><a href='https://example.com'>链接</a>", "中文😀");
            var original = new AgentRichText(oe).Signature(Page(oe), false);
            ParagraphStyles.Apply(oe, "heading1", Empty, new AgentOptions());
            Equal(original, new AgentRichText(oe).Signature(Page(oe), false));
            Equal(2, oe.Elements(One + "T").Count());
            True(oe.ToString().Contains("&nbsp;"));
        });
        Test("native OneNote unquoted HTML attributes and local font names", () =>
        {
            var oe = Paragraph("a", "<span\nstyle='font-family:微软雅黑' lang=zh-CN>&nbsp;先</span><a href=\"https://example.com/?a=1&amp;b=2\"><span lang=x-none>链接</span></a><span style='font-family:\"Segoe UI Emoji\"'>&#128512;</span>");
            var rich = new AgentRichText(oe); Equal("\u00a0先链接😀", rich.Text);
            rich.Format(1, 1, new Dictionary<string, string> { ["font-weight"] = "bold" });
            Equal("\u00a0先链接😀", new AgentRichText(oe).Text);
        });
        Test("removing bold affects only exact selected characters", () =>
        {
            var oe = Paragraph("a", "<b>甲乙丙</b>");
            new AgentRichText(oe).Format(1, 1, new Dictionary<string, string> { ["font-weight"] = "normal" });
            var html = oe.Element(One + "T").Value;
            True(html.Contains("normal")); True(html.Contains("<b>甲</b>")); True(html.Contains("<b>丙</b>"));
            Equal("甲乙丙", new AgentRichText(oe).Text);
        });
        Test("surrogate and combining character boundaries rejected", () =>
        {
            var rich = new AgentRichText(Paragraph("a", "😀e\u0301"));
            Throws(() => rich.Format(1, 1, new Dictionary<string, string> { ["color"] = "#222222" }));
            Throws(() => rich.Format(2, 1, new Dictionary<string, string> { ["color"] = "#222222" }));
        });
        Test("unknown/unclosed HTML protected", () =>
        {
            var snapshot = Snapshot(Page(Paragraph("a", "<img src='x'>"), Paragraph("b", "<b>text"), Paragraph("c", "ok")));
            Equal(1, snapshot.Blocks.Count(b => b.Editable));
        });
        Test("emoji joiner and skin tone cannot be split", () =>
        {
            var rich = new AgentRichText(Paragraph("a", "👩‍💻👍🏽"));
            Throws(() => rich.Format(0, 2, new Dictionary<string, string> { ["font-weight"] = "bold" }));
            Throws(() => rich.Format(5, 2, new Dictionary<string, string> { ["font-weight"] = "bold" }));
        });
        Test("changing underline preserves strike-through", () =>
        {
            var oe = Paragraph("a", "<s><u>保留删除线</u></s>");
            new AgentRichText(oe).Format(0, 5, new Dictionary<string, string> { ["text-decoration"] = "none" });
            True(oe.Element(One + "T").Value.Contains("line-through"));
        });
        Test("selection scope and inherited code font", () =>
        {
            var code = Paragraph("code", "code"); code.SetAttributeValue("style", "font-family:Consolas");
            var snapshot = new AgentPageSnapshot(Page(code, Paragraph("a", "text"), Paragraph("b", "secret")).ToString(), new HashSet<string> { "a", "code" }, new AgentOptions());
            Equal(2, snapshot.Blocks.Count); True(!snapshot.Blocks[0].Editable); True(!snapshot.Blocks.Any(b => b.Text == "secret"));
        });
        Test("mixed Outline is protected until enabled", () =>
        {
            var page = Page(Paragraph("a", "text"), new XElement(One + "OE", new XAttribute("objectID", "image"), new XElement(One + "Image", new XElement(One + "CallbackID", "binary"))));
            True(!new AgentPageSnapshot(page.ToString(), null, new AgentOptions { EnableMixedOutlines = false }).Blocks[0].Editable);
            True(new AgentPageSnapshot(page.ToString(), null, new AgentOptions { EnableMixedOutlines = true }).Blocks[0].Editable);
        });
        Test("tool requires complete read and rejects unknown fields", () =>
        {
            var s = Snapshot(); var t = Tools(s);
            Throws(() => Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" }, preset_id = "body" }));
            Throws(() => Invoke(t, "get_page_overview", new { pageId = "other" }));
            Equal(0, s.Revision);
        });
        Test("atomic batch leaves no partial draft on bad target", () =>
        {
            var s = Snapshot(); var t = Tools(s); Read(t, s);
            Throws(() => Invoke(t, "set_text_style", new { snapshot_id = s.SnapshotId, targets = new object[] {
                new { block_id = "p1", quote = "第一段", occurrence = 1, style = new { bold = true } },
                new { block_id = "p2", quote = "missing", occurrence = 1, style = new { bold = true } } } }));
            Equal(0, s.Revision); True(s.Blocks.All(b => !b.Changed));
        });
        Test("preset updates are idempotent", () =>
        {
            var s = Snapshot(); var t = Tools(s); Read(t, s); Style(t, s);
            var version = s.Revision; Style(t, s); Equal(version, s.Revision);
        });
        Test("page title preset allowed only on Title", () =>
        {
            var p = Page(Paragraph("a", "body")); p.AddFirst(new XElement(One + "Title", Paragraph("title", "标题")));
            var s = Snapshot(p); var t = Tools(s); Read(t, s);
            Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" }, preset_id = "page_title" });
            Throws(() => Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p2" }, preset_id = "page_title" }));
        });
        Test("capabilities reject unverified paragraph spacing", () =>
        {
            var s = Snapshot(); s.Options.EnableParagraphSpacing = false; var t = Tools(s); Read(t, s);
            Throws(() => Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" }, preset_id = "body", overrides = new { space_after_pt = 10 } }));
        });
        Test("commit writes once, verifies and records undo", () =>
        {
            var s = Prepared(); var api = new FakePage(s.Page); var report = new AgentCommitter(api).Commit(s, CancellationToken.None);
            Equal("Verified", report.Status); Equal(1, report.Applied); Equal(1, api.Writes); Equal(1, report.Undo.Count);
            Equal("第一段", new AgentRichText(AgentCommitter.Find(api.Page, "a")).Text);
        });
        Test("concurrent format-only edit skipped", () => Conflict(p => AgentCommitter.Find(p, "a").SetAttributeValue("style", "font-size:19pt")));
        Test("concurrent text edit skipped", () => Conflict(p => AgentCommitter.Find(p, "a").Element(One + "T").Value = "用户修改"));
        Test("concurrent deletion skipped", () => Conflict(p => AgentCommitter.Find(p, "a").Remove()));
        Test("concurrent ancestor format skipped", () => Conflict(p => p.Element(One + "Outline").SetAttributeValue("style", "font-size:19pt")));
        Test("concurrent reorder skipped", () => Conflict(p => { var oe = AgentCommitter.Find(p, "a"); oe.Remove(); p.Descendants(One + "OEChildren").First().Add(oe); }));
        Test("shared quick style changes are conflicts", () =>
        {
            var p = Page(Paragraph("a", "第一段")); p.AddFirst(new XElement(One + "QuickStyleDef", new XAttribute("index", "0"), new XAttribute("font", "Arial"), new XAttribute("fontSize", "11")));
            AgentCommitter.Find(p, "a").SetAttributeValue("quickStyleIndex", "0");
            var s = Prepared(p); var api = new FakePage(p); api.Page.Element(One + "QuickStyleDef").SetAttributeValue("fontSize", "20");
            Equal(1, new AgentCommitter(api).Commit(s, CancellationToken.None).Conflicts); Equal(0, api.Writes);
        });
        Test("timestamp conflict re-reads and preserves concurrent other paragraph", () =>
        {
            var s = Prepared(); var api = new FakePage(s.Page) { ConflictsRemaining = 1 };
            api.OnConflict = () => AgentCommitter.Find(api.Page, "b").Element(One + "T").Value = "new content";
            var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
            Equal(1, r.Applied); Equal(2, api.Attempts); Equal("new content", new AgentRichText(AgentCommitter.Find(api.Page, "b")).Text);
        });
        Test("parent and child staged together do not conflict with own changes", () =>
        {
            var parent = Paragraph("a", "父段落"); parent.Add(new XElement(One + "OEChildren", Paragraph("b", "子段落")));
            var s = Snapshot(Page(parent)); var t = Tools(s); Read(t, s);
            Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1", "p2" }, preset_id = "body" });
            var api = new FakePage(s.Page); var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
            Equal(2, r.Applied); Equal(0, r.Conflicts);
        });
        Test("new quick style index collision rebased without altering user definition", () =>
        {
            var s = Prepared(); var api = new FakePage(s.Page);
            api.Page.AddFirst(new XElement(One + "QuickStyleDef", new XAttribute("index", "0"), new XAttribute("name", "custom"), new XAttribute("font", "Arial"), new XAttribute("fontSize", "29")));
            var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal("custom", (string)api.Page.Elements(One + "QuickStyleDef").First(d => (string)d.Attribute("index") == "0").Attribute("name"));
        });
        Test("persistent timestamp conflict never forces write", () =>
        {
            var s = Prepared(); var api = new FakePage(s.Page) { ConflictsRemaining = 5 };
            Throws(() => new AgentCommitter(api).Commit(s, CancellationToken.None)); Equal(3, api.Attempts); Equal(0, api.Writes);
        });
        Test("missing timestamp cannot write", () =>
        {
            var s = Prepared(); var api = new FakePage(s.Page); api.Page.Attribute("lastModifiedTime").Remove();
            Throws(() => new AgentCommitter(api).Commit(s, CancellationToken.None)); Equal(0, api.Attempts);
        });
        Test("COM exception after successful save is verified, not retried", () =>
        {
            var s = Prepared(); var api = new FakePage(s.Page) { ThrowAfterSave = true };
            var r = new AgentCommitter(api).Commit(s, CancellationToken.None); Equal("Verified", r.Status); Equal(1, api.Attempts);
        });
        Test("readback failure reported as unknown", () =>
        {
            var s = Prepared(); var api = new FakePage(s.Page) { FailReadAfterSave = true };
            var r = new AgentCommitter(api).Commit(s, CancellationToken.None); Equal("CommitOutcomeUnknown", r.Status); Equal(0, r.Undo.Count);
        });
        Test("protected code formatting altered during write is not reported verified", () =>
        {
            var code = Paragraph("b", "code"); code.SetAttributeValue("style", "font-family:Consolas;font-size:11pt");
            var s = Prepared(Page(Paragraph("a", "第一段"), code)); var api = new FakePage(s.Page);
            api.AfterSave = () => AgentCommitter.Find(api.Page, "b").SetAttributeValue("style", "font-family:Arial");
            Equal("CommitOutcomeUnknown", new AgentCommitter(api).Commit(s, CancellationToken.None).Status);
        });
        Test("OneNote table wrapper replacement and automatic width are normalized", () =>
        {
            var s = Prepared(TablePage(false)); var api = new FakePage(s.Page);
            api.AfterSave = () =>
            {
                api.Page.Descendants(One + "Table").Single().Parent.SetAttributeValue("objectID", "new-wrapper");
                api.Page.Descendants(One + "Column").Single().SetAttributeValue("width", "85.25");
            };
            Equal("Verified", new AgentCommitter(api).Commit(s, CancellationToken.None).Status);
        });
        Test("locked column width changes are not reported verified", () =>
        {
            var s = Prepared(TablePage(true)); var api = new FakePage(s.Page);
            api.AfterSave = () => api.Page.Descendants(One + "Column").Single().SetAttributeValue("width", "250");
            Equal("CommitOutcomeUnknown", new AgentCommitter(api).Commit(s, CancellationToken.None).Status);
        });
        Test("cell shading changes are not reported verified", () =>
        {
            var s = Prepared(TablePage(false)); var api = new FakePage(s.Page);
            api.AfterSave = () => api.Page.Descendants(One + "Cell").Single().SetAttributeValue("shadingColor", "#FF0000");
            Equal("CommitOutcomeUnknown", new AgentCommitter(api).Commit(s, CancellationToken.None).Status);
        });
        Test("table row replacement remains a structural mismatch", () =>
        {
            var s = Prepared(TablePage(false)); var api = new FakePage(s.Page);
            api.AfterSave = () => api.Page.Descendants(One + "Row").Single().SetAttributeValue("objectID", "different-row");
            Equal("CommitOutcomeUnknown", new AgentCommitter(api).Commit(s, CancellationToken.None).Status);
        });
        Test("cancel before commit has no writes", () =>
        {
            var s = Prepared(); var api = new FakePage(s.Page); var c = new CancellationTokenSource(); c.Cancel();
            Throws<OperationCanceledException>(() => new AgentCommitter(api).Commit(s, c.Token)); Equal(0, api.Writes);
        });
        Test("cancel during COM still verifies actual result", () =>
        {
            var s = Prepared(); var api = new FakePage(s.Page); var c = new CancellationTokenSource(); api.AfterSave = () => c.Cancel();
            Equal("Verified", new AgentCommitter(api).Commit(s, c.Token).Status);
        });
        Test("undo restores original formatting", () =>
        {
            var s = Prepared(); var api = new FakePage(s.Page); var committer = new AgentCommitter(api); var r = committer.Commit(s, CancellationToken.None);
            var undo = committer.Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal(1, undo.Applied); Equal((string)AgentCommitter.Find(s.Page, "a").Attribute("style"), (string)AgentCommitter.Find(api.Page, "a").Attribute("style"));
        });
        Test("undo skips later user formatting", () =>
        {
            var s = Prepared(); var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            AgentCommitter.Find(api.Page, "a").SetAttributeValue("style", "font-size:22pt");
            Equal(1, c.Undo(s.PageId, r, s.Options, CancellationToken.None).Conflicts); Equal(1, api.Writes);
        });
        Test("undo handles OneNote coalescing multiple T runs", () =>
        {
            var s = Prepared(Page(Paragraph("a", "第一", "段"))); var api = new FakePage(s.Page);
            api.AfterSave = () =>
            {
                var oe = AgentCommitter.Find(api.Page, "a"); var text = string.Concat(oe.Elements(One + "T").Select(t => t.Value));
                oe.Elements(One + "T").Remove(); oe.AddFirst(new XElement(One + "T", new XCData(text)));
            };
            var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None); Equal(1, r.Applied);
            api.AfterSave = null; Equal(1, c.Undo(s.PageId, r, s.Options, CancellationToken.None).Applied);
        });
        Test("text replacement keeps formatting and links of untouched characters", () =>
        {
            var oe = Paragraph("a", "我们<b>按装</b>了&nbsp;<a href='https://example.com'>链结</a>", "尾巴");
            new AgentRichText(oe).Replace(2, 2, "安装");
            new AgentRichText(oe).Replace(6, 2, "链接");
            var rich = new AgentRichText(oe);
            Equal("我们安装了 链接尾巴", rich.Text);
            var html = oe.Element(One + "T").Value;
            True(html.Contains("<b>安装</b>")); True(html.Contains("href=\"https://example.com\">链接</a>")); True(html.Contains("&nbsp;"));
            Equal(2, oe.Elements(One + "T").Count()); Equal("尾巴", oe.Elements(One + "T").Last().Value);
            // 纯删除（「的的」改「的」）和跨格式边界的修正：没变的字保留各自的格式。
            var mixed = Paragraph("b", "<b>重要</b>的的事");
            new AgentRichText(mixed).Replace(1, 3, "要的");
            Equal("重要的事", new AgentRichText(mixed).Text); True(mixed.Element(One + "T").Value.StartsWith("<b>重要</b>的", StringComparison.Ordinal));
            Throws(() => new AgentRichText(Paragraph("c", "第一<br>第二")).Replace(1, 2, "一 第"));
        });
        Test("fix_text rejects bad targets without side effects", () =>
        {
            var mono = Paragraph("m", "int x = 1;"); mono.SetAttributeValue("style", "font-family:Consolas");
            var s = Snapshot(Page(Paragraph("a", "我们按装了软件"), mono)); var t = Tools(s);
            Rejects("请先完整读取", () => Fix(t, s, "p1", "按装", "安装"));
            Read(t, s); Invoke(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p2" } });
            Rejects("找不到", () => Fix(t, s, "p1", "安装", "按装"));
            Rejects("相同", () => Fix(t, s, "p1", "按装", "按装"));
            Rejects("换行", () => Fix(t, s, "p1", "按装", "安\n装"));
            Rejects("代码段落", () => Fix(t, s, "p2", "int", "var"));
            Rejects("字符串无效", () => Fix(t, s, "p1", "按装", new string('安', AgentTools.MaxFixChars + 1)));
            Rejects("找不到", () => Invoke(t, "fix_text", new { snapshot_id = s.SnapshotId, fixes = new object[] {
                new { block_id = "p1", quote = "按装", occurrence = 1, replacement = "安装" },
                new { block_id = "p1", quote = "软体", occurrence = 1, replacement = "软件" } } }));
            Equal(0, s.Revision); True(s.Blocks.All(b => !b.Changed && b.TextFixes.Count == 0));
        });
        Test("fix_text commits verified text, combines with styles and undo restores it", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "我们按装了软件"), Paragraph("b", "第二段"))); var t = Tools(s); Read(t, s);
            var result = AgentChatClient.Serializer().Serialize(Fix(t, s, "p1", "按装", "安装"));
            True(result.Contains("我们安装了软件")); Equal(1, s.Revision);
            // 修正后读到的、局部格式定位用的都是改过的文字。
            True(AgentChatClient.Serializer().Serialize(Invoke(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" } })).Contains("我们安装了软件"));
            Invoke(t, "set_text_style", new { snapshot_id = s.SnapshotId, targets = new[] { new { block_id = "p1", quote = "安装", occurrence = 1, style = new { bold = true } } } });
            Style(t, s);
            True(AgentChatClient.Serializer().Serialize(Invoke(t, "get_pending_changes", new { snapshot_id = s.SnapshotId })).Contains("「按装」→「安装」"));
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(1, r.Applied); Equal(1, api.Writes); True(r.Message.Contains("修正文字 1 处"));
            Equal("「按装」→「安装」", r.TextFixes.Single());
            var written = AgentCommitter.Find(api.Page, "a");
            Equal("我们安装了软件", new AgentRichText(written).Text); True(written.Attribute("style").Value.Contains("16pt"));
            Equal("第二段", new AgentRichText(AgentCommitter.Find(api.Page, "b")).Text);
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal("Verified", undo.Status); True(undo.Message.Contains("还原文字 1 处"));
            Equal("我们按装了软件", new AgentRichText(AgentCommitter.Find(api.Page, "a")).Text);
            Equal((string)AgentCommitter.Find(s.Page, "a").Attribute("style"), (string)AgentCommitter.Find(api.Page, "a").Attribute("style"));
        });
        Test("fixed paragraph edited during processing is skipped", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "我们按装了软件"))); var t = Tools(s); Read(t, s); Fix(t, s, "p1", "按装", "安装");
            var api = new FakePage(s.Page); AgentCommitter.Find(api.Page, "a").Element(One + "T").Value = "我们按装了软件，用户又补了一句";
            var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
            Equal(1, r.Conflicts); Equal(0, api.Writes); Equal(0, r.TextFixes.Count);
        });
        Test("text changed outside fix_text is blocked at commit", () =>
        {
            var s = Prepared(); s.Blocks[0].Draft.Element(One + "T").Value = "偷改的文字";
            var api = new FakePage(s.Page);
            Rejects("改变了正文", () => new AgentCommitter(api).Commit(s, CancellationToken.None)); Equal(0, api.Writes);
        });
        Test("code conversion discards pending text fixes", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "示例："), Paragraph("c", "x = 1 # 按装"))); var t = Tools(s); Read(t, s);
            Fix(t, s, "p2", "按装", "安装"); True(s.Blocks[1].TextFixes.Count == 1);
            Code(t, s, "python", "p2");
            Equal(0, s.Blocks[1].TextFixes.Count); Equal("x = 1 # 按装", s.Blocks[1].CurrentText);
        });
        Test("stream interleaved tool argument fragments", StreamTest);
        Test("DONE without finish reason cannot execute", () =>
        {
            var r = new AgentReply(); AgentChatClient.AbsorbEvent("[DONE]", r); Throws(r.Validate);
        });
        Test("length finish cannot execute tools", () => Throws(new AgentReply { FinishReason = "length" }.Validate));
        Test("stream error or malformed JSON rejected", () =>
        {
            Throws(() => AgentChatClient.AbsorbEvent("{broken", new AgentReply()));
            Throws(() => AgentChatClient.AbsorbEvent("{\"error\":{\"message\":\"private\"}}", new AgentReply()));
        });
        Test("runner completes native multi-turn tools", () =>
        {
            var s = Snapshot(); var api = new FakePage(s.Page); var model = new ScriptedClient(s);
            var r = new AgentRunner(model, new AgentCommitter(api)).RunAsync(s, "美化", null, CancellationToken.None).GetAwaiter().GetResult();
            Equal("Verified", r.Status); Equal(1, api.Writes); True(model.SawToolResult); True(model.SawReasoning);
        });
        Test("runner reports turns and one summarized step per tool", () =>
        {
            var s = Snapshot(); var api = new FakePage(s.Page); var log = new ProgressLog();
            var r = new AgentRunner(new ScriptedClient(s), new AgentCommitter(api)).RunAsync(s, "美化", log, CancellationToken.None).GetAwaiter().GetResult();
            Equal("Verified", r.Status);
            Equal(5, log.Items.Max(p => p.Turn));
            Equal(5, log.Items.Count(p => p.Step?.State == AgentStepState.Running));
            var finished = log.Items.Where(p => p.Step != null && p.Step.State != AgentStepState.Running).Select(p => p.Step).ToList();
            Equal(5, finished.Count); True(finished.All(step => step.State == AgentStepState.Done));
            Equal("读取页面概况 · 共 " + s.Blocks.Count + " 段", finished[0].Text);
            Equal("读取段落 · 1 段", finished[1].Text);
            Equal("设置段落样式 · 一级标题 · 1 段", finished[2].Text);
            Equal("检查格式草稿", finished[3].Text);
            Equal("写回并验证", finished[4].Text);
        });
        Test("step descriptions summarize arguments, results and failures", () =>
        {
            Equal(("读取段落 · 2 段", AgentStepState.Done), AgentTools.DescribeStep("read_blocks", "{\"block_ids\":[\"a\",\"b\"]}", "{}"));
            Equal(("设置重点文字样式 · 2 处", AgentStepState.Done), AgentTools.DescribeStep("set_text_style", "{\"targets\":[{},{}]}", "{\"ok\":true}"));
            Equal(("修正错别字 · 3 处", AgentStepState.Done), AgentTools.DescribeStep("fix_text", "{\"fixes\":[{},{},{}]}", "{\"ok\":true}"));
            Equal(("高亮代码 · " + LanguageRegistry.Find("python").DisplayName + " · 3 段", AgentStepState.Done),
                AgentTools.DescribeStep("highlight_code", "{\"block_ids\":[\"a\",\"b\",\"c\"],\"language\":\"auto\"}", "{\"ok\":true,\"language\":\"python\"}"));
            Equal(("高亮代码 · 自动识别 · 1 段：无法自动识别代码语言。", AgentStepState.Failed),
                AgentTools.DescribeStep("highlight_code", "{\"block_ids\":[\"a\"],\"language\":\"auto\"}", "{\"ok\":false,\"error\":\"无法自动识别代码语言。\"}"));
            Equal(("读取段落", AgentStepState.Done), AgentTools.DescribeStep("read_blocks", "{broken", null));
            Equal(("校验工具请求：未知工具。", AgentStepState.Failed), AgentTools.DescribeStep("no_such_tool", "{}", "{\"ok\":false,\"error\":\"未知工具。\"}"));
        });
        Test("runner text-only claim is not reported as success", () =>
        {
            var s = Snapshot(); var api = new FakePage(s.Page);
            var r = new AgentRunner(new TextOnlyClient(), new AgentCommitter(api)).RunAsync(s, "美化", null, CancellationToken.None).GetAwaiter().GetResult();
            Equal("NoChange", r.Status); Equal(0, api.Writes);
        });
        Test("runner turn budget discards uncommitted draft", () =>
        {
            var s = Snapshot(); s.Options.MaxTurns = 3; var api = new FakePage(s.Page);
            Throws(() => new AgentRunner(new ScriptedClient(s), new AgentCommitter(api)).RunAsync(s, "美化", null, CancellationToken.None).GetAwaiter().GetResult());
            Equal(0, api.Writes);
        });
        Test("real HTTP client parses SSE and omits old JSON output mode", () =>
        {
            var handler = new StubHttp("data: {\"choices\":[{\"index\":0,\"delta\":{\"tool_calls\":[{\"index\":0,\"id\":\"x\",\"function\":{\"name\":\"get_page_overview\",\"arguments\":\"{}\"}}]},\"finish_reason\":\"tool_calls\"}]}\n\ndata: [DONE]\n\n", "text/event-stream");
            var c = AiConfigStore.Parse(XElement.Parse("<AiConfig><ApiUrl>https://test.invalid</ApiUrl><ApiKey>test</ApiKey></AiConfig>"));
            using (var http = new HttpClient(handler))
            {
                var reply = new AgentChatClient(c, "test", "none", http).CompleteAsync(new List<object> { new { role = "user", content = "test" } }, Tools(Snapshot()).Definitions, null, CancellationToken.None).GetAwaiter().GetResult();
                Equal("get_page_overview", reply.Calls[0].Name); True(!handler.Body.Contains("response_format"));
                True(handler.Body.Contains("tool_choice"));
            }
        });
        Test("HTTP stream progress shows reasoning excerpt and the tool being prepared", () =>
        {
            var handler = new StubHttp("data: {\"choices\":[{\"index\":0,\"delta\":{\"reasoning_content\":\"先读取段落\\n\\n再定标题\"}}]}\n\n" +
                "data: {\"choices\":[{\"index\":0,\"delta\":{\"tool_calls\":[{\"index\":0,\"id\":\"x\",\"function\":{\"name\":\"read_blocks\",\"arguments\":\"{}\"}}]},\"finish_reason\":\"tool_calls\"}]}\n\ndata: [DONE]\n\n", "text/event-stream");
            var c = AiConfigStore.Parse(XElement.Parse("<AiConfig><ApiUrl>https://test.invalid</ApiUrl><ApiKey>test</ApiKey></AiConfig>"));
            var log = new ProgressLog();
            using (var http = new HttpClient(handler))
                new AgentChatClient(c, "test", "medium", http).CompleteAsync(new List<object>(), new object[0], log, CancellationToken.None).GetAwaiter().GetResult();
            // 第一段立即报告；后面的被限频吞掉，收完再报一次最终状态。
            Equal("模型正在思考…", log.Items[0].Status); Equal("先读取段落\n再定标题", log.Items[0].Thinking);
            Equal("模型正在准备：读取段落", log.Items.Last().Status);
        });
        Test("SSE wrapper overhead above former 600K limit preserves complete tool call", () =>
        {
            var events = new System.Text.StringBuilder();
            for (var i = 0; i < 12000; i++)
                events.Append("data: {\"choices\":[{\"index\":0,\"delta\":{\"reasoning_content\":\"x\"}}]}\n\n");
            events.Append("data: {\"choices\":[{\"index\":0,\"delta\":{\"tool_calls\":[{\"index\":0,\"id\":\"last\",\"function\":{\"name\":\"get_page_overview\",\"arguments\":\"{}\"}}]},\"finish_reason\":\"tool_calls\"}]}\n\ndata: [DONE]\n\n");
            True(events.Length > 600000);
            var handler = new StubHttp(events.ToString(), "text/event-stream");
            var c = AiConfigStore.Parse(XElement.Parse("<AiConfig><ApiUrl>https://test.invalid</ApiUrl><ApiKey>test</ApiKey></AiConfig>"));
            using (var http = new HttpClient(handler))
            {
                var reply = new AgentChatClient(c, "test", "medium", http).CompleteAsync(new List<object>(), new object[0], null, CancellationToken.None).GetAwaiter().GetResult();
                Equal(12000, reply.Reasoning.Length); Equal("get_page_overview", reply.Calls[0].Name);
                Equal(reply.Reasoning, (string)((IDictionary<string, object>)reply.ToMessage(true))["reasoning_content"]);
            }
        });
        Test("oversized useful SSE output still stops before tools execute", () =>
        {
            var chunk = new string('x', 300000);
            var sse = "data: {\"choices\":[{\"index\":0,\"delta\":{\"reasoning_content\":\"" + chunk + "\"}}]}\n\n" +
                "data: {\"choices\":[{\"index\":0,\"delta\":{\"reasoning_content\":\"" + chunk + "\"}}]}\n\n" +
                "data: {\"choices\":[{\"index\":0,\"delta\":{},\"finish_reason\":\"tool_calls\"}]}\n\ndata: [DONE]\n\n";
            var handler = new StubHttp(sse, "text/event-stream");
            var c = AiConfigStore.Parse(XElement.Parse("<AiConfig><ApiUrl>https://test.invalid</ApiUrl><ApiKey>test</ApiKey></AiConfig>"));
            using (var http = new HttpClient(handler))
                Throws(() => new AgentChatClient(c, "test", "medium", http).CompleteAsync(new List<object>(), new object[0], null, CancellationToken.None).GetAwaiter().GetResult());
        });
        Test("HTTP 400 does not echo upstream private data", () =>
        {
            var handler = new StubHttp("PRIVATE_PAGE_CONTENT", "text/plain") { Status = HttpStatusCode.BadRequest };
            var c = AiConfigStore.Parse(XElement.Parse("<AiConfig><ApiUrl>https://test.invalid</ApiUrl><ApiKey>test</ApiKey></AiConfig>"));
            using (var http = new HttpClient(handler))
            {
                try { new AgentChatClient(c, "test", "none", http).CompleteAsync(new List<object>(), new object[0], null, CancellationToken.None).GetAwaiter().GetResult(); throw new Exception("Expected HTTP failure"); }
                catch (AiException ex) { True(!ex.Message.Contains("PRIVATE_PAGE_CONTENT")); True(ex.Message.Contains("400")); }
            }
        });
        Test("configuration absent Agent node preserves defaults", () =>
        {
            var c = AiConfigStore.Parse(XElement.Parse("<AiConfig><Agent><MaxTurns>999</MaxTurns><ReplayReasoning>false</ReplayReasoning></Agent></AiConfig>"));
            Equal(30, c.Agent.MaxTurns); True(!c.Agent.ReplayReasoning); True(c.Agent.EnableMixedOutlines);
        });
        Test("code classification: plain text, unhighlighted code, code box and inline code", () =>
        {
            var mono = Paragraph("m", "int x = 1;"); mono.SetAttributeValue("style", "font-family:Consolas");
            var spans = Paragraph("s", "<span style='font-family:Consolas'>var</span><span style=\"font-family:'Courier New'\"> y;</span>");
            var inline = Paragraph("i", "调用 <span style='font-family:Consolas'>Run()</span> 方法");
            var box = new XElement(One + "OE", new XAttribute("objectID", "w"), CodeBlockBuilder.BuildTable("a = 1", LanguageRegistry.Find("python"), CodeThemes.Light, new AddInSettings()));
            var n = 0; foreach (var e in box.Descendants().Where(e => e.Name.LocalName != "Columns" && e.Name.LocalName != "Column" && e.Name.LocalName != "OEChildren" && e.Name.LocalName != "T")) e.SetAttributeValue("objectID", "box" + n++);
            var s = Snapshot(Page(Paragraph("a", "正文"), mono, spans, inline, box));
            Equal(null, s.Blocks[0].ProtectedReason); Equal("unhighlighted_code", s.Blocks[1].ProtectedReason); Equal("unhighlighted_code", s.Blocks[2].ProtectedReason);
            Equal("protected_code", s.Blocks[3].ProtectedReason); Equal("highlighted_code", s.Blocks[4].ProtectedReason);
            var t = Tools(s);
            Invoke(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p2" } });
            Throws(() => Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p2" }, preset_id = "body" }));
            Throws(() => Invoke(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p5" } }));
        });
        Test("code highlight disabled keeps whole monospace paragraphs protected", () =>
        {
            var mono = Paragraph("m", "int x = 1;"); mono.SetAttributeValue("style", "font-family:Consolas");
            var s = new AgentPageSnapshot(Page(mono).ToString(), null, new AgentOptions { EnableCodeHighlight = false });
            Equal("protected_code", s.Blocks[0].ProtectedReason);
            True(!AgentChatClient.Serializer().Serialize(Tools(s).Definitions).Contains("highlight_code"));
        });
        Test("highlight_code converts contiguous paragraphs, verifies and records undo", () =>
        {
            var s = CodePrepared(); var api = new FakePage(s.Page);
            var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(1, r.CodeBlocks); Equal(1, api.Writes); Equal(1, r.CodeUndo.Count);
            True(AgentCommitter.Find(api.Page, "c1") == null && AgentCommitter.Find(api.Page, "c3") == null);
            var box = AgentCommitter.Find(api.Page, "a").ElementsAfterSelf().First();
            var expected = CodeBlockBuilder.BuildTable("def f(x):\n\n    return x + 1", LanguageRegistry.Find("python"), CodeThemes.Light, new AddInSettings());
            Equal(string.Join("|", expected.Descendants(One + "OE").Select(AgentCode.PlainText)), string.Join("|", box.Descendants(One + "OE").Select(AgentCode.PlainText)));
            Equal("结尾", new AgentRichText(AgentCommitter.Find(api.Page, "b")).Text);
        });
        Test("highlight_code rejects invalid ranges without side effects", () =>
        {
            var parent = Paragraph("n1", "if x:"); parent.Add(new XElement(One + "OEChildren", Paragraph("n2", "print(x)")));
            var listed = Paragraph("l", "SELECT 1;"); listed.AddFirst(new XElement(One + "List", new XElement(One + "Bullet", new XAttribute("bullet", "2"))));
            var p = CodePage(); p.Descendants(One + "OEChildren").First().Add(parent, Paragraph("n3", "y = 2"), listed);
            p.AddFirst(new XElement(One + "Title", Paragraph("title", "x = 1")));
            var s = Snapshot(p); var t = Tools(s);
            Rejects("请先完整读取", () => Code(t, s, "python", "p3", "p4", "p5"));
            Read(t, s);
            Rejects("必须连续", () => Code(t, s, "python", "p3", "p5"));
            Rejects("页面标题", () => Code(t, s, "python", "p1"));
            Rejects("无法自动识别", () => Code(t, s, "auto", "p2"));
            Rejects("下级段落", () => Code(t, s, "python", "p7"));
            Rejects("缩进更深", () => Code(t, s, "python", "p8", "p9"));
            Rejects("项目符号", () => Code(t, s, "sql", "p10"));
            Equal(0, s.Revision); True(s.Blocks.All(b => b.Conversion == null));
        });
        Test("conversion discards pending format and blocks later styling", () =>
        {
            var s = Snapshot(CodePage()); var t = Tools(s); Read(t, s);
            Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p2" }, preset_id = "body" });
            True(s.Blocks[1].Changed);
            var result = AgentChatClient.Serializer().Serialize(Code(t, s, "python", "p2", "p3", "p4"));
            True(result.Contains("discarded_format")); True(result.Contains("\"p2\""));
            True(!s.Blocks[1].Changed); Equal(2, s.Revision);
            Throws(() => Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p2" }, preset_id = "body" }));
            Code(t, s, "python", "p4", "p3", "p2"); Equal(2, s.Revision);
            Throws(() => Code(t, s, "python", "p2", "p3"));
        });
        Test("concurrent edit of source code paragraph skips the conversion", () =>
        {
            var s = CodePrepared(); var api = new FakePage(s.Page);
            AgentCommitter.Find(api.Page, "c3").Element(One + "T").Value = "return 2";
            var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
            Equal("NoChange", r.Status); Equal(3, r.Conflicts); Equal(0, api.Writes);
        });
        Test("paragraph format and code conversion commit together", () =>
        {
            var s = CodePrepared(); var t = Tools(s);
            Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" }, preset_id = "heading1" });
            var api = new FakePage(s.Page); var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(1, r.Applied); Equal(1, r.CodeBlocks); Equal(1, api.Writes); Equal(2, r.Undo.Count + r.CodeUndo.Count);
        });
        Test("undo turns the code box back into the original paragraphs", () =>
        {
            var s = CodePrepared(); var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal("Verified", undo.Status); Equal(1, undo.CodeBlocks); Equal(2, api.Writes);
            True(!api.Page.Descendants(One + "Table").Any());
            Equal("示例：|def f(x):||    return x + 1|结尾", string.Join("|", api.Page.Descendants(One + "OE").Select(AgentCode.PlainText)));
            Equal("font-size:9pt", (string)api.Page.Descendants(One + "OE").ElementAt(1).Attribute("style"));
        });
        Test("undo skips a code box edited afterwards", () =>
        {
            var s = CodePrepared(); var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            api.Page.Descendants(One + "Table").Single().Descendants(One + "T").First().Value = "def g(x):";
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal(1, undo.Conflicts); Equal(0, undo.CodeBlocks); Equal(1, api.Writes);
        });
        Test("indented code keeps nesting through conversion and undo", () =>
        {
            var parent = Paragraph("c1", "def f():"); parent.Add(new XElement(One + "OEChildren", Paragraph("c2", "return 1")));
            var s = Snapshot(Page(Paragraph("a", "示例："), parent)); var t = Tools(s); Read(t, s);
            Code(t, s, "python", "p2", "p3");
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status);
            Equal("def f():|    return 1", string.Join("|", api.Page.Descendants(One + "Table").Single().Descendants(One + "OE").Select(AgentCode.PlainText)));
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status);
            var restored = api.Page.Descendants(One + "OE").ElementAt(1);
            Equal("def f():", AgentCode.PlainText(restored)); Equal("return 1", AgentCode.PlainText(restored.Element(One + "OEChildren").Element(One + "OE")));
        });
        Test("runner converts code through the highlight_code tool", () =>
        {
            var s = Snapshot(CodePage()); var api = new FakePage(s.Page); var model = new CodeScriptClient(s);
            var r = new AgentRunner(model, new AgentCommitter(api), new AddInSettings()).RunAsync(s, "排版", null, CancellationToken.None).GetAwaiter().GetResult();
            Equal("Verified", r.Status); Equal(1, r.CodeBlocks); Equal(1, api.Writes); True(model.SawCodeTool);
        });
        Console.WriteLine($"Agent: {_passed} passed, {_failed} failed");
        return _failed == 0 ? 0 : 1;
    }

    internal static XElement Paragraph(string id, params string[] runs) => new XElement(One + "OE", new XAttribute("objectID", id), runs.Select(t => new XElement(One + "T", new XCData(t))));
    internal static XElement Page(params XElement[] paragraphs) => new XElement(One + "Page", new XAttribute("ID", "page"), new XAttribute("lastModifiedTime", "2026-09-26T00:00:00Z"),
        new XElement(One + "Outline", new XAttribute("objectID", "outline"), new XAttribute("style", "font-family:Arial;font-size:11pt;color:#222222"), new XElement(One + "OEChildren", paragraphs)));
    private static AgentPageSnapshot Snapshot(XElement page = null) => new AgentPageSnapshot((page ?? Page(Paragraph("a", "第一段"), Paragraph("b", "第二段"))).ToString(), null, new AgentOptions());
    private static XElement TablePage(bool locked) => Page(new XElement(One + "OE", new XAttribute("objectID", "wrapper"),
        new XElement(One + "Table", new XAttribute("objectID", "table"),
            new XElement(One + "Columns", new XElement(One + "Column", new XAttribute("index", "0"), new XAttribute("width", "200"), new XAttribute("isLocked", locked))),
            new XElement(One + "Row", new XAttribute("objectID", "row"), new XElement(One + "Cell", new XAttribute("objectID", "cell"), new XElement(One + "OEChildren", Paragraph("a", "第一段")))))));
    private static AgentTools Tools(AgentPageSnapshot s) => new AgentTools(s, new AgentCommitter(new FakePage(s.Page)), CancellationToken.None);
    private static object Invoke(AgentTools tools, string name, object args) => tools.Execute(new AgentToolCall { Id = "test", Name = name, Arguments = AgentChatClient.Serializer().Serialize(args) });
    private static void Read(AgentTools tools, AgentPageSnapshot s) => Invoke(tools, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = s.Blocks.Where(b => b.Editable).Select(b => b.Id).ToArray() });
    private static void Style(AgentTools tools, AgentPageSnapshot s) => Invoke(tools, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" }, preset_id = "heading1" });
    private static AgentPageSnapshot Prepared(XElement page = null) { var s = Snapshot(page); var t = Tools(s); Read(t, s); Style(t, s); return s; }
    private static object Fix(AgentTools tools, AgentPageSnapshot s, string id, string quote, string replacement) =>
        Invoke(tools, "fix_text", new { snapshot_id = s.SnapshotId, fixes = new[] { new { block_id = id, quote, occurrence = 1, replacement } } });
    /// <summary>正文、三行代码（中间一个空行）、正文。代码段落 p2–p4。</summary>
    private static XElement CodePage()
    {
        var indented = Paragraph("c3", "&nbsp;&nbsp;&nbsp;&nbsp;return x + 1"); indented.SetAttributeValue("style", "font-size:9pt");
        var first = Paragraph("c1", "def f(x):"); first.SetAttributeValue("style", "font-size:9pt");
        return Page(Paragraph("a", "示例："), first, Paragraph("c2", ""), indented, Paragraph("b", "结尾"));
    }
    private static object Code(AgentTools tools, AgentPageSnapshot s, string language, params string[] ids) =>
        Invoke(tools, "highlight_code", new { snapshot_id = s.SnapshotId, block_ids = ids, language });
    private static AgentPageSnapshot CodePrepared() { var s = Snapshot(CodePage()); var t = Tools(s); Read(t, s); Code(t, s, "python", "p2", "p3", "p4"); return s; }
    private static void Conflict(Action<XElement> mutate)
    {
        var s = Prepared(); var api = new FakePage(s.Page); mutate(api.Page);
        var r = new AgentCommitter(api).Commit(s, CancellationToken.None); Equal(1, r.Conflicts); Equal(0, api.Writes);
    }
    private static void StreamTest()
    {
        var r = new AgentReply();
        AgentChatClient.AbsorbEvent("{\"choices\":[{\"index\":0,\"delta\":{\"reasoning_content\":\"think\",\"tool_calls\":[{\"index\":0,\"id\":\"c0\",\"type\":\"function\",\"function\":{\"name\":\"read_blocks\",\"arguments\":\"{\\\"a\\\":\"}},{\"index\":1,\"id\":\"c1\",\"function\":{\"name\":\"get_page_overview\",\"arguments\":\"{\"}}]}}]}", r);
        AgentChatClient.AbsorbEvent("{\"choices\":[{\"index\":0,\"delta\":{\"tool_calls\":[{\"index\":1,\"function\":{\"arguments\":\"}\"}},{\"index\":0,\"function\":{\"arguments\":\"1}\"}}]},\"finish_reason\":\"tool_calls\"}]}", r);
        r.Validate(); Equal("{\"a\":1}", r.Calls[0].Arguments); Equal("{}", r.Calls[1].Arguments); Equal("think", r.Reasoning);
    }
    private static void Test(string name, Action test)
    {
        try { test(); _passed++; Console.WriteLine("PASS " + name); }
        catch (Exception ex) { _failed++; Console.WriteLine("FAIL " + name + ": " + ex); }
    }
    internal static void True(bool condition) { if (!condition) throw new Exception("Assertion failed"); }
    internal static void Equal<T>(T expected, T actual) { if (!EqualityComparer<T>.Default.Equals(expected, actual)) throw new Exception("Expected " + expected + "; actual " + actual); }
    private static void Throws(Action body) => Throws<AiException>(body);
    private static void Rejects(string reason, Action body)
    {
        try { body(); } catch (AiException ex) { if (ex.Message.Contains(reason)) return; throw new Exception("Expected '" + reason + "'; actual " + ex.Message); }
        throw new Exception("Expected AiException");
    }
    private static void Throws<T>(Action body) where T : Exception
    { try { body(); } catch (T) { return; } throw new Exception("Expected " + typeof(T).Name); }

    internal sealed class FakePage : IOneNotePageAccess
    {
        internal XElement Page;
        internal int Attempts, Writes, ConflictsRemaining;
        internal bool ThrowAfterSave, FailReadAfterSave;
        internal Action OnConflict, AfterSave;
        private int _ids;
        internal FakePage(XElement page) { Page = new XElement(page); }
        public string GetPageContent(string id, PageInfo info)
        { if (FailReadAfterSave && Writes > 0) throw new Exception("read failed"); return Page.ToString(); }
        public void UpdatePageContent(string xml, DateTime expected)
        {
            Attempts++; True(expected != DateTime.MinValue);
            if (ConflictsRemaining-- > 0) { OnConflict?.Invoke(); throw new COMException("conflict", unchecked((int)0x80042010)); }
            foreach (var c in XElement.Parse(xml).Elements())
            {
                // OneNote 给新建的段落、表格分配 ID。
                foreach (var e in c.DescendantsAndSelf().Where(e => new[] { "OE", "Table", "Row", "Cell" }.Contains(e.Name.LocalName) && e.Attribute("objectID") == null))
                    e.SetAttributeValue("objectID", "new-" + ++_ids);
                var identity = c.Name == One + "QuickStyleDef" ? "index" : "objectID";
                var current = Page.Elements(c.Name).FirstOrDefault(e => (string)e.Attribute(identity) == (string)c.Attribute(identity));
                if (current == null) Page.Add(new XElement(c)); else current.ReplaceWith(new XElement(c));
            }
            Writes++; Page.SetAttributeValue("lastModifiedTime", DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture));
            AfterSave?.Invoke();
            if (ThrowAfterSave) throw new COMException("uncertain");
        }
    }
    private sealed class StubHttp : HttpMessageHandler
    {
        private readonly string _response, _media;
        internal string Body;
        internal HttpStatusCode Status = HttpStatusCode.OK;
        internal StubHttp(string response, string media) { _response = response; _media = media; }
        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellation)
        {
            Body = await request.Content.ReadAsStringAsync();
            return new HttpResponseMessage(Status) { Content = new StringContent(_response, System.Text.Encoding.UTF8, _media) };
        }
    }
    /// <summary>同步记下每条进度；Progress&lt;T&gt; 会投递到同步上下文，测试里看不到顺序。</summary>
    private sealed class ProgressLog : IProgress<AgentProgress>
    {
        internal readonly List<AgentProgress> Items = new List<AgentProgress>();
        public void Report(AgentProgress value) { lock (Items) Items.Add(value); }
    }
    private sealed class TextOnlyClient : IAgentChatClient
    {
        public Task<AgentReply> CompleteAsync(List<object> messages, object[] tools, IProgress<AgentProgress> progress, CancellationToken cancellation) =>
            Task.FromResult(new AgentReply { Content = "已完成", FinishReason = "stop" });
    }
    private sealed class CodeScriptClient : IAgentChatClient
    {
        private readonly AgentPageSnapshot _s; private int _turn;
        internal bool SawCodeTool;
        internal CodeScriptClient(AgentPageSnapshot s) { _s = s; }
        public Task<AgentReply> CompleteAsync(List<object> messages, object[] tools, IProgress<AgentProgress> progress, CancellationToken cancellation)
        {
            SawCodeTool |= AgentChatClient.Serializer().Serialize(tools).Contains("highlight_code");
            string name; object args;
            switch (_turn++)
            {
                case 0: name = "read_blocks"; args = new { snapshot_id = _s.SnapshotId, block_ids = new[] { "p1", "p2", "p4", "p5" } }; break;
                case 1: name = "highlight_code"; args = new { snapshot_id = _s.SnapshotId, block_ids = new[] { "p2", "p3", "p4" }, language = "auto" }; break;
                default: name = "finish_edit"; args = new { snapshot_id = _s.SnapshotId, draft_revision = _s.Revision }; break;
            }
            var reply = new AgentReply { FinishReason = "tool_calls" };
            reply.Calls.Add(0, new AgentToolCall { Id = "k" + _turn, Name = name, Arguments = AgentChatClient.Serializer().Serialize(args) });
            return Task.FromResult(reply);
        }
    }
    private sealed class ScriptedClient : IAgentChatClient
    {
        private readonly AgentPageSnapshot _s; private int _turn;
        internal bool SawToolResult, SawReasoning;
        internal ScriptedClient(AgentPageSnapshot s) { _s = s; }
        public Task<AgentReply> CompleteAsync(List<object> messages, object[] tools, IProgress<AgentProgress> progress, CancellationToken cancellation)
        {
            var json = AgentChatClient.Serializer().Serialize(messages);
            SawToolResult |= json.Contains("tool_call_id"); SawReasoning |= json.Contains("reasoning_content");
            string name; object args;
            switch (_turn++)
            {
                case 0: name = "get_page_overview"; args = new { }; break;
                case 1: name = "read_blocks"; args = new { snapshot_id = _s.SnapshotId, block_ids = new[] { "p1" } }; break;
                case 2: name = "set_paragraph_style"; args = new { snapshot_id = _s.SnapshotId, block_ids = new[] { "p1" }, preset_id = "heading1" }; break;
                case 3: name = "get_pending_changes"; args = new { snapshot_id = _s.SnapshotId }; break;
                default: name = "finish_edit"; args = new { snapshot_id = _s.SnapshotId, draft_revision = _s.Revision }; break;
            }
            var reply = new AgentReply { FinishReason = "tool_calls", Reasoning = "opaque provider history" };
            reply.Calls.Add(0, new AgentToolCall { Id = "c" + _turn, Name = name, Arguments = AgentChatClient.Serializer().Serialize(args) });
            return Task.FromResult(reply);
        }
    }
}
