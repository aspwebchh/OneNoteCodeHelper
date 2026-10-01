using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
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
        if (args.Length == 2 && args[0] == "--probe-code-spacing") return CodeSpacingProbe.Run(args[1]);
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
        Test("undo maps quick styles and tags by content after OneNote renumbers them", () =>
        {
            var todo = Paragraph("a", "第一段"); todo.AddFirst(Tag("0"));
            var p = Page(todo, Paragraph("b", "第二段")); p.AddFirst(TagDef("0", 3, "待办事项"));
            p.AddFirst(new XElement(One + "QuickStyleDef", new XAttribute("index", "0"), new XAttribute("name", "p"), new XAttribute("font", "Calibri"), new XAttribute("fontSize", "11")));
            foreach (var oe in p.Descendants(One + "OE")) oe.SetAttributeValue("quickStyleIndex", "0");
            var s = Snapshot(p); var t = Tools(s); Read(t, s); Style(t, s);
            Invoke(t, "set_tag", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" }, tag = "important" });
            Invoke(t, "set_tag", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" }, tag = "none" });
            Invoke(t, "set_tag", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p2" }, tag = "important" });
            var api = new FakePage(s.Page);
            // OneNote 回存后按自己的顺序给样式和标记定义重新编号：新加的 h1、重要排到 0，原来的正文、待办变成 1。
            api.AfterSave = () =>
            {
                foreach (var name in new[] { "QuickStyleDef", "TagDef" })
                    foreach (var d in api.Page.Elements(One + name)) d.SetAttributeValue("index", (string)d.Attribute("index") == "0" ? "1" : "0");
                foreach (var a in api.Page.Descendants().Attributes("quickStyleIndex")) a.Value = a.Value == "0" ? "1" : "0";
                foreach (var tag in api.Page.Descendants(One + "Tag")) tag.SetAttributeValue("index", (string)tag.Attribute("index") == "0" ? "1" : "0");
            };
            var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None); Equal("Verified", r.Status);
            api.AfterSave = null;
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status);
            string Named(XElement oe) => (string)api.Page.Elements(One + "QuickStyleDef").Single(d => (string)d.Attribute("index") == (string)oe.Attribute("quickStyleIndex")).Attribute("name");
            var a0 = AgentCommitter.Find(api.Page, "a"); Equal("p", Named(a0)); Equal("p", Named(AgentCommitter.Find(api.Page, "b")));
            Equal("3", (string)AgentMarks.Definition(api.Page, a0.Elements(One + "Tag").Single()).Attribute("symbol"));
            True(!AgentCommitter.Find(api.Page, "b").Elements(One + "Tag").Any());
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
        Test("markdown analysis removes only marker characters", () =>
        {
            Equal("标题|目标|code 项目|第一|已完成|引用 重点|粗斜与删除|C# 和 #标签|snake_case_name 和 2 * 3 * 4 和 a*b*c 和 *.md|**不是强调**|",
                Strip("# 标题", "## 目标 ##", "- `code` 项目", "1. 第一", "  - [x] 已完成", "> 引用 **重点**", "***粗斜***与~~删除~~",
                    "C# 和 #标签", "snake_case_name 和 2 * 3 * 4 和 a*b*c 和 *.md", "`**不是强调**`", "---"));
            // markdown 围栏里照常处理，里面带语言的围栏是代码；不写语言的围栏也是代码。
            Equal("|标题||// *x* # y||正文|||- 保留|", Strip("```markdown", "# 标题", "```java", "// *x* # y", "```", "**正文**", "```", "```", "- 保留", "```"));
            var kinds = new HashSet<string>(AgentMarkdown.Kinds);
            var todo = AgentMarkdown.Analyze("  - [ ] 待办", new[] { AgentMarkdown.LineRole.Text }, kinds, false);
            Equal("bullet", todo.List); Equal(false, todo.Todo); Equal(2, todo.Indent); True(!todo.Separator);
            Equal(3, AgentMarkdown.Analyze("### 三级", new[] { AgentMarkdown.LineRole.Text }, kinds, false).Heading);
            Equal("number", AgentMarkdown.Analyze("12) 第十二", new[] { AgentMarkdown.LineRole.Text }, kinds, false).List);
            True(AgentMarkdown.Analyze("> 引用", new[] { AgentMarkdown.LineRole.Text }, kinds, false).Quote);
            True(AgentMarkdown.Analyze("* * *", new[] { AgentMarkdown.LineRole.Text }, kinds, false).Separator);
            // 只处理部分种类：编号留着，行内代码照样不当强调。
            var numbered = AgentMarkdown.Analyze("1. `a*b*` **粗**", new[] { AgentMarkdown.LineRole.Text }, new HashSet<string> { "emphasis" }, true);
            Equal("1. `a*b*` 粗", AgentMarkdown.Remove("1. `a*b*` **粗**", numbered.Marks)); Equal(null, numbered.List);
            Equal(12, numbered.Formats.Single().Start); Equal("bold", numbered.Formats.Single().Css["font-weight"]);
        });
        Test("strip_markdown removes markers and separator lines, commits verified text and undo restores it", () =>
        {
            var s = Snapshot(Page(Paragraph("f1", "```markdown"), Paragraph("h", "# 标题"), Paragraph("l", "- `code` 项目"), Paragraph("b", "<b>**重点**</b>文字"),
                Paragraph("r", "---"), Paragraph("f2", "```"), Paragraph("x", "普通段落")));
            var t = Tools(s);
            Rejects("请先完整读取", () => Invoke(t, "strip_markdown", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p2" } }));
            Read(t, s);
            var result = Json(Invoke(t, "strip_markdown", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1", "p2", "p3", "p4", "p5", "p6", "p7" } }));
            True(result.Contains("\"removed_lines\":[\"p1\",\"p5\",\"p6\"]")); True(result.Contains("\"noop\":[\"p7\"]"));
            True(result.Contains("{\"id\":\"p2\",\"heading\":1")); True(result.Contains("{\"id\":\"p3\",\"heading\":null,\"list\":\"bullet\""));
            Equal("标题", s.Blocks[1].CurrentText); Equal("重点文字", s.Blocks[3].CurrentText);
            // 删掉的段落不再出现在概况里，读取时跳过，格式和文字工具都拒绝。
            True(!Json(Invoke(t, "get_page_overview", new { })).Contains("\"id\":\"p1\""));
            True(Json(Invoke(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1", "p2" } })).Contains("{\"id\":\"p1\",\"reason\":\"removed\"}"));
            Rejects("已删除", () => Fix(t, s, "p1", "`", "'"));
            Rejects("已删除", () => Style(t, s));
            Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p2" }, preset_id = "heading1" });
            True(Json(Invoke(t, "get_pending_changes", new { snapshot_id = s.SnapshotId })).Contains("\"markdown_removed\":[\"p1\",\"p5\",\"p6\"]"));
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(9, r.MarkdownMarks); True(r.Message.Contains("去除 Markdown 符号 9 处")); Equal(0, r.TextFixes.Count);
            Equal("标题|code 项目|重点文字|普通段落", Texts(api.Page));
            // 删符号不动其余文字的格式：「重点」仍是粗体。
            True(AgentCommitter.Find(api.Page, "b").Element(One + "T").Value.Contains("<b>重点</b>"));
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal("Verified", undo.Status);
            Equal("```markdown|# 标题|- `code` 项目|**重点**文字|---|```|普通段落", Texts(api.Page));
        });
        Test("strip_markdown formats emphasis, keeps fenced code and empties separators it cannot remove", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "这是**重点**和*斜体*"))); var t = Tools(s); Read(t, s);
            Invoke(t, "strip_markdown", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" }, emphasis = "format" });
            var api = new FakePage(s.Page); var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(4, r.MarkdownMarks);
            var html = AgentCommitter.Find(api.Page, "a").Element(One + "T").Value;
            Equal("这是重点和斜体", new AgentRichText(AgentCommitter.Find(api.Page, "a")).Text);
            True(html.Contains("font-weight:bold") && html.Contains("font-style:italic"));

            var page = Page(Paragraph("f1", "```"), Paragraph("c", "- item: *x*"), Paragraph("f2", "```"));
            var off = new AgentPageSnapshot(page.ToString(), null, new AgentOptions { EnableBlankLineRemoval = false }); var o = Tools(off); Read(o, off);
            var result = Json(Invoke(o, "strip_markdown", new { snapshot_id = off.SnapshotId, block_ids = new[] { "p1", "p2", "p3" } }));
            True(result.Contains("\"emptied\":[\"p1\",\"p3\"]")); True(result.Contains("\"code_lines\":[\"p2\"]")); True(result.Contains("\"removed_lines\":[]"));
            var offApi = new FakePage(off.Page); Equal("Verified", new AgentCommitter(offApi).Commit(off, CancellationToken.None).Status);
            Equal("|- item: *x*|", Texts(offApi.Page));
            // 文本框里只剩分隔线时留一段，清空文字。
            var alone = Snapshot(Page(Paragraph("r", "***"))); var a = Tools(alone); Read(a, alone);
            True(Json(Invoke(a, "strip_markdown", new { snapshot_id = alone.SnapshotId, block_ids = new[] { "p1" } })).Contains("\"emptied\":[\"p1\"]"));
            Equal("", alone.Blocks[0].CurrentText);
        });
        Test("strip_markdown is a switchable tool with prompt and step text", () =>
        {
            var page = Page(Paragraph("a", "# 标题")).ToString();
            True(!Tools(new AgentPageSnapshot(page, null, new AgentOptions { EnableMarkdownCleanup = false })).Has("strip_markdown"));
            var tools = Tools(new AgentPageSnapshot(page, null, new AgentOptions()));
            True(tools.Has("strip_markdown")); True(AgentRunner.SystemPrompt(tools).Contains(AgentRunner.MarkdownPrompt)); True(AgentRunner.SystemPrompt(tools).Contains("highlight_code 转换"));
            Equal(("去除 Markdown 符号 · 3 段 · 删除围栏和分隔线 2 行", AgentStepState.Done),
                AgentTools.DescribeStep("strip_markdown", "{\"block_ids\":[\"a\"]}", "{\"ok\":true,\"changed\":[{},{},{}],\"removed_lines\":[\"x\",\"y\"]}"));
        });
        Test("markdown fenced code survives repeated and split cleanup with either removal setting", () =>
        {
            foreach (var remove in new[] { true, false }) foreach (var split in new[] { true, false })
            {
                var page = Page(Paragraph("f1", "```java"), Paragraph("c", "var s = \"**keep**\";"), Paragraph("f2", "```"), Paragraph("h", "# 正文"));
                var original = Texts(page);
                var s = new AgentPageSnapshot(page.ToString(), null, new AgentOptions { EnableBlankLineRemoval = remove }); var t = Tools(s); Read(t, s);
                if (split) Cleanup(t, s, "p1", "p3");
                Cleanup(t, s, "p1", "p2", "p3", "p4");
                var revision = s.Revision;
                True(Json(Cleanup(t, s, "p1", "p2", "p3", "p4")).Contains("\"code_lines\":[\"p2\"]"));
                Equal(revision, s.Revision); Equal("var s = \"**keep**\";", s.Blocks[1].CurrentText);
                var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
                Equal("Verified", r.Status); Equal(3, r.MarkdownMarks);
                Equal("var s = \"**keep**\";", new AgentRichText(AgentCommitter.Find(api.Page, "c")).Text);
                var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
                Equal("Verified", undo.Status); Equal(3, undo.MarkdownMarks); Equal(original, Texts(api.Page));
            }
        });
        Test("markdown selection reads fence context locally and preserves unselected paragraphs", () =>
        {
            var page = Page(Paragraph("f1", "```java"), Paragraph("c", "var s = \"**keep**\";"), Paragraph("f2", "```"),
                Paragraph("h", "# 标题"), Paragraph("secret", "**选区外**"));
            var s = new AgentPageSnapshot(page.ToString(), new HashSet<string> { "c", "h" }, new AgentOptions()); var t = Tools(s); Read(t, s);
            var overview = Json(Invoke(t, "get_page_overview", new { })); True(!overview.Contains("选区外") && !overview.Contains("```java"));
            True(Json(Cleanup(t, s, "p1", "p2")).Contains("\"code_lines\":[\"p1\"]"));
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(1, r.MarkdownMarks);
            Equal("```java|var s = \"**keep**\";|```|标题|**选区外**", Texts(api.Page));
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status); Equal(Texts(page), Texts(api.Page));
        });
        Test("markdown code protection follows a paragraph moved to another text box", () =>
        {
            var page = Boxes(Box("A", 100, Paragraph("f1", "```java"), Paragraph("c", "var s = \"**keep**\";"), Paragraph("f2", "```")),
                Box("B", 300, Paragraph("b", "正文")));
            var s = Snapshot(page); var t = Tools(s); Read(t, s); Move(t, s, new[] { "p2" }, "p4", "after");
            True(Json(Cleanup(t, s, "p2")).Contains("\"code_lines\":[\"p2\"]")); Equal(0, s.Blocks[1].MarkdownMarks);
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal("正文|var s = \"**keep**\";", BoxTexts(api.Page, "B"));
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status); Equal(Texts(page), Texts(api.Page));
        });
        Test("markdown fence contexts stay separate across table cells and surrounding text", () =>
        {
            var page = GridPage(); var children = page.Element(One + "Outline").Element(One + "OEChildren");
            children.AddFirst(Paragraph("f1", "```java")); children.Add(Paragraph("c", "var s = \"**keep**\";"), Paragraph("f2", "```"));
            var cell = page.Descendants(One + "Cell").First().Element(One + "OEChildren");
            cell.ReplaceNodes(Paragraph("cf1", "```java"), Paragraph("cc", "var s = \"**cell**\";"), Paragraph("cf2", "```"));
            AgentCommitter.Find(page, "h2").Element(One + "T").Value = "**说明**";
            var s = Snapshot(page); var t = Tools(s); Read(t, s); Cleanup(t, s, s.Blocks.Select(b => b.Id).ToArray());
            Equal("var s = \"**keep**\";", s.Blocks.Single(b => b.ObjectId == "c").CurrentText);
            Equal("var s = \"**cell**\";", s.Blocks.Single(b => b.ObjectId == "cc").CurrentText);
            Equal("说明", s.Blocks.Single(b => b.ObjectId == "h2").CurrentText);
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(6, r.MarkdownMarks);
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status); Equal(Texts(page), Texts(api.Page));
        });
        Test("markdown skips only the text flow with unreadable fence context", () =>
        {
            var page = Boxes(Box("A", 100, Paragraph("bad", "<img src='x'>"), Paragraph("a", "**保留**")), Box("B", 300, Paragraph("b", "# 标题")));
            var s = new AgentPageSnapshot(page.ToString(), new HashSet<string> { "a", "b" }, new AgentOptions()); var t = Tools(s); Read(t, s);
            True(Json(Cleanup(t, s, "p1", "p2")).Contains("markdown_context_unavailable"));
            Equal("**保留**", s.Blocks[0].CurrentText); Equal("标题", s.Blocks[1].CurrentText);
            var api = new FakePage(s.Page); var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(1, r.MarkdownMarks); Equal("<img src='x'>", AgentCommitter.Find(api.Page, "bad").Element(One + "T").Value);
        });
        Test("markdown recognizes OneNote hard spaces without changing content whitespace", () =>
        {
            var s = Snapshot(Page(Paragraph("h", "&nbsp;##&nbsp;标题&nbsp;正文&nbsp;##"), Paragraph("l", "&nbsp;&nbsp;-&nbsp;[&nbsp;]&nbsp;待办"),
                Paragraph("q", "&nbsp;&gt;&nbsp;引用"), Paragraph("r", "&nbsp;*&nbsp;*&nbsp;*"), Paragraph("f1", "&nbsp;```java"),
                Paragraph("c", "var s = \"**keep**\";"), Paragraph("f2", "&nbsp;```")));
            var t = Tools(s); Read(t, s); var result = Json(Cleanup(t, s, s.Blocks.Select(b => b.Id).ToArray()));
            Equal("标题\u00a0正文", s.Blocks[0].CurrentText); Equal("待办", s.Blocks[1].CurrentText); Equal("引用", s.Blocks[2].CurrentText);
            True(result.Contains("\"todo\":false") && result.Contains("\"indent\":2") && result.Contains("\"heading\":2"));
            Equal("var s = \"**keep**\";", s.Blocks[5].CurrentText);
            var api = new FakePage(s.Page); var r = new AgentCommitter(api).Commit(s, CancellationToken.None); Equal("Verified", r.Status);
            Equal("标题\u00a0正文", new AgentRichText(AgentCommitter.Find(api.Page, "h")).Text);
        });
        Test("markdown pairs nested emphasis and preserves unmatched markers and inline code", () =>
        {
            Equal("粗体里有 斜体|斜体里有 粗体|粗斜|删除里有 粗体|粗体 删除|粗体 斜体|粗体 斜体 正文|**未闭合|*未闭合**|snake_case 和 a*b*c 和 2 * 3|粗体 *保留*|😀 𠮷字",
                Strip("**粗体里有 *斜体***", "*斜体里有 **粗体***", "***粗斜***", "~~删除里有 **粗体**~~", "**粗体 ~~删除~~**", "__粗体 _斜体___",
                    "**粗体 *斜体* 正文**", "**未闭合", "*未闭合**", "snake_case 和 a*b*c 和 2 * 3", "**粗体 `*保留*`**", "**😀 *𠮷字***"));
            foreach (var pair in new[] { ("**粗体里有 *斜体***", "<b>粗体里有 <i>斜体</i></b>"), ("*斜体里有 **粗体***", "<i>斜体里有 <b>粗体</b></i>"),
                ("~~删除里有 **粗体**~~", "<s>删除里有 <b>粗体</b></s>"), ("***粗斜***", "<b><i>粗斜</i></b>") })
            {
                var s = Snapshot(Page(Paragraph("a", pair.Item1))); var t = Tools(s); Read(t, s);
                Invoke(t, "strip_markdown", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" }, emphasis = "format" });
                var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None); Equal("Verified", r.Status);
                var expected = Paragraph("a", pair.Item2); var expectedPage = Page(expected);
                Equal(new AgentRichText(expected).Signature(expectedPage, true), new AgentRichText(AgentCommitter.Find(api.Page, "a")).Signature(api.Page, true));
                Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status); Equal(pair.Item1, Texts(api.Page));
            }
        });
        Test("markdown outer emphasis survives unmatched inner markers in removal and format modes", () =>
        {
            var cases = new[] {
                ("**粗体中有 _未闭合下划线**", "粗体中有 _未闭合下划线", "<b>粗体中有 _未闭合下划线</b>"),
                ("**粗体中有 *未闭合星号**", "粗体中有 *未闭合星号", "<b>粗体中有 *未闭合星号</b>"),
                ("**正文 _悬空 *斜体***", "正文 _悬空 斜体", "<b>正文 _悬空 <i>斜体</i></b>"),
                ("*斜体 **字面星号*", "斜体 **字面星号", "<i>斜体 **字面星号</i>"),
                ("~~删除中有 _悬空~~", "删除中有 _悬空", "<s>删除中有 _悬空</s>"),
                ("**甲 _悬空** __乙 *悬空__", "甲 _悬空 乙 *悬空", "<b>甲 _悬空</b> <b>乙 *悬空</b>"),
                ("**😀&nbsp;<a href='https://example.com'>_𠮷字</a>**", "😀\u00a0_𠮷字", "<b>😀&nbsp;<a href='https://example.com'>_𠮷字</a></b>") };
            foreach (var format in new[] { false, true }) foreach (var item in cases)
            {
                var s = Snapshot(Page(Paragraph("a", item.Item1))); var t = Tools(s); Read(t, s);
                Invoke(t, "strip_markdown", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" }, emphasis = format ? "format" : "remove" });
                Equal(item.Item2, s.Blocks[0].CurrentText);
                var revision = s.Revision; Cleanup(t, s, "p1"); Equal(revision, s.Revision);
                var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
                Equal("Verified", r.Status); Equal(0, r.Unverified);
                var html = format ? item.Item3 : item.Item3.Replace("<b>", "").Replace("</b>", "").Replace("<i>", "").Replace("</i>", "").Replace("<s>", "").Replace("</s>", "");
                var expected = Paragraph("a", html);
                Equal(new AgentRichText(expected).Signature(Page(expected), true), new AgentRichText(AgentCommitter.Find(api.Page, "a")).Signature(api.Page, true));
                Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status); Equal(Texts(s.Page), Texts(api.Page));
            }
        });
        Test("markdown fallback keeps incomplete closers literal and does not pair across a closed outer span", () =>
        {
            Equal("*未闭合**|**未闭合*|__未闭合_|粗体 _悬空 后续_|粗体 ~~悬空 后续~~",
                Strip("*未闭合**", "**未闭合*", "__未闭合_", "**粗体 _悬空** 后续_", "**粗体 ~~悬空** 后续~~"));
        });
        Test("markdown fallback handles many incomplete delimiters without repeated backward scans", () =>
        {
            var text = string.Concat(Enumerable.Repeat("***开 *开 内容** ", 20000));
            var watch = System.Diagnostics.Stopwatch.StartNew();
            var result = AgentMarkdown.Analyze(text, new[] { AgentMarkdown.LineRole.Text }, new HashSet<string>(AgentMarkdown.Kinds), false);
            watch.Stop(); Equal(0, result.Marks.Count); True(watch.Elapsed < TimeSpan.FromSeconds(2));
        });
        Test("cleaned markdown tables preserve text edits, format, links, counts and undo in both commit paths", () =>
        {
            foreach (var structure in new[] { true, false }) foreach (var rows in new[] { true, false })
            {
                var paragraphs = new List<XElement> { Paragraph("h", "| **名称** | 内容 |") };
                if (!rows) paragraphs.Add(Paragraph("sep", "| --- | --- |"));
                paragraphs.Add(Paragraph("d", "| <a href='https://example.com'>按装</a> | **重点** |"));
                if (structure) { paragraphs.Insert(0, Paragraph("f1", "```markdown")); paragraphs.Add(Paragraph("f2", "```")); }
                var page = Page(paragraphs.ToArray()); var original = Texts(page); var s = Snapshot(page); var t = Tools(s); Read(t, s);
                Fix(t, s, s.Blocks.Single(b => b.ObjectId == "d").Id, "按装", "安装");
                Invoke(t, "strip_markdown", new { snapshot_id = s.SnapshotId, block_ids = s.Blocks.Select(b => b.Id).ToArray(), emphasis = "format" });
                var ids = s.Blocks.Where(b => b.ObjectId == "h" || b.ObjectId == "d" || b.ObjectId == "sep").Select(b => b.Id).ToArray();
                var result = rows
                    ? Invoke(t, "text_to_table", new { snapshot_id = s.SnapshotId, block_ids = ids, rows = new[] { new[] { "名称", "内容" }, new[] { "安装", "重点" } }, header_row = false })
                    : Invoke(t, "text_to_table", new { snapshot_id = s.SnapshotId, block_ids = ids, delimiter = "pipe", header_row = false });
                True(Json(result).Contains("\"discarded_format\":[]")); Equal("名称|内容|安装|重点", Texts(s.CodeConversions.Single().Table));
                var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
                Equal("Verified", r.Status); Equal(1, r.TextTables); Equal(structure ? 6 : 4, r.MarkdownMarks); Equal(1, r.TextFixes.Count);
                Equal("名称|内容|安装|重点", Texts(api.Page));
                var html = api.Page.ToString(); True(html.Contains("font-weight:bold") && html.Contains("https://example.com"));
                var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
                Equal("Verified", undo.Status); Equal(r.MarkdownMarks, undo.MarkdownMarks); Equal(1, undo.TextFixes.Count); Equal(original, Texts(api.Page));
            }
        });
        Test("cleaned markdown table conflicts never overwrite user text or count skipped edits", () =>
        {
            foreach (var structure in new[] { true, false })
            {
                var paragraphs = new List<XElement> { Paragraph("h", "| **名称** | 内容 |"), Paragraph("d", "| 甲 | **重点** |") };
                if (structure) { paragraphs.Insert(0, Paragraph("f1", "```markdown")); paragraphs.Add(Paragraph("f2", "```")); }
                var s = Snapshot(Page(paragraphs.ToArray())); var t = Tools(s); Read(t, s); Cleanup(t, s, s.Blocks.Select(b => b.Id).ToArray());
                Table(t, s, "pipe", s.Blocks.Where(b => b.ObjectId == "h" || b.ObjectId == "d").Select(b => b.Id).ToArray());
                var api = new FakePage(s.Page); AgentCommitter.Find(api.Page, "d").Element(One + "T").Value = "用户改动"; var before = api.Page.ToString();
                var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
                Equal("NoChange", r.Status); Equal(0, api.Writes); Equal(0, r.MarkdownMarks); Equal(0, r.TextFixes.Count); Equal(before, api.Page.ToString());
            }
        });
        Test("cleaned markdown table rejects lost emphasis on readback and protects later edits on undo", () =>
        {
            var s = Snapshot(Page(Paragraph("h", "| **名称** | 内容 |"), Paragraph("d", "| 甲 | **重点** |"))); var t = Tools(s); Read(t, s);
            Invoke(t, "strip_markdown", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1", "p2" }, emphasis = "format" });
            Table(t, s, "pipe", "p1", "p2");
            var broken = new FakePage(s.Page);
            broken.AfterSave = () => broken.Page.Descendants(One + "OE").Last(oe => oe.Elements(One + "T").Any()).Element(One + "T").Value = "重点";
            var failed = new AgentCommitter(broken).Commit(s, CancellationToken.None);
            Equal("PartiallyApplied", failed.Status); Equal(1, failed.Unverified); Equal(0, failed.MarkdownMarks); True(!failed.CanUndo);
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None); Equal("Verified", r.Status);
            api.Page.Descendants(One + "OE").Last(oe => oe.Elements(One + "T").Any()).Element(One + "T").Value = "后来编辑"; var before = api.Page.ToString();
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal(1, undo.Conflicts); Equal(0, undo.MarkdownMarks); Equal(before, api.Page.ToString());
        });
        Test("table conversion without markdown retains its previous draft discard behavior", () =>
        {
            var s = Snapshot(Page(Paragraph("h", "名称 | 内容"), Paragraph("d", "按装 | 重点"))); var t = Tools(s); Read(t, s); Fix(t, s, "p2", "按装", "安装");
            Table(t, s, "pipe", "p1", "p2"); Equal("名称|内容|按装|重点", Texts(s.CodeConversions.Single().Table));
            Equal(0, s.CodeConversions[0].TextFixes.Count); Equal(0, s.Blocks[1].TextFixes.Count);
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
        Test("unfinished replies explain the finish reason", () =>
        {
            Rejects("MaxTokens", new AgentReply { FinishReason = "length", Done = true }.Validate);
            Rejects("连接中途结束", new AgentReply().Validate);
            Rejects("没有返回结束原因", new AgentReply { Done = true }.Validate);
            Rejects("结束原因 aborted", new AgentReply { FinishReason = "aborted", Done = true }.Validate);
            Rejects("结束原因 未知", new AgentReply { FinishReason = "<b>私有内容</b>", Done = true }.Validate);
        });
        Test("usage-only stream chunk is recorded and summarized without content", () =>
        {
            var r = new AgentReply();
            AgentChatClient.AbsorbEvent("{\"choices\":[{\"index\":0,\"delta\":{\"reasoning_content\":\"私有思考\"},\"finish_reason\":\"length\"}]}", r);
            AgentChatClient.AbsorbEvent("{\"usage\":{\"prompt_tokens\":12,\"completion_tokens\":16384,\"completion_tokens_details\":{\"reasoning_tokens\":15000}}}", r);
            var line = AgentChatClient.Summarize(r, "m", "low", 100, 1.5);
            True(line.Contains("finish=length")); True(line.Contains("12/16384/15000")); True(!line.Contains("私有思考"));
            Throws(() => AgentChatClient.AbsorbEvent("{\"id\":\"x\"}", new AgentReply()));
        });
        Test("runner retries a truncated turn with a reminder and keeps the draft", () =>
        {
            var s = Snapshot(); var api = new FakePage(s.Page); var log = new ProgressLog();
            // 第一轮和设样式前各截断一次：丢掉那轮、提醒分批，脚本接着往下走。
            var model = new TruncatingClient(new ScriptedClient(s), 0, 3);
            var r = new AgentRunner(model, new AgentCommitter(api)).RunAsync(s, "美化", log, CancellationToken.None).GetAwaiter().GetResult();
            Equal("Verified", r.Status); Equal(1, api.Writes); Equal(7, model.Sent.Count);
            Equal(2, log.Items.Count(p => p.Step?.State == AgentStepState.Note));
            // 第一轮截断后提醒并进用户需求，不出现连续两条 user 消息；截断的思考和半截工具调用不进历史。
            Equal(1, model.Sent[1].Split(new[] { "\"role\":\"user\"" }, StringSplitOptions.None).Length - 1);
            True(model.Sent[1].Contains(AgentRunner.TruncatedPrompt));
            True(model.Sent.All(json => !json.Contains("TRUNCATED_THOUGHT") && !json.Contains("\"cut")));
        });
        Test("runner stops after repeated truncation and names MaxTokens", () =>
        {
            var s = Snapshot(); var api = new FakePage(s.Page);
            var model = new TruncatingClient(new ScriptedClient(s), Enumerable.Range(0, 24).ToArray());
            Rejects("MaxTokens", () => new AgentRunner(model, new AgentCommitter(api)).RunAsync(s, "美化", null, CancellationToken.None).GetAwaiter().GetResult());
            Equal(AgentRunner.MaxTruncatedRetries + 1, model.Sent.Count); Equal(0, api.Writes);
        });
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
            // 工具结果不带 null 字段，HTML 字符不转义。
            True(model.LastJson.Contains("tool_call_id")); True(!model.LastJson.Contains("\\\":null")); True(!model.LastJson.Contains("\\u003c"));
        });
        Test("runner budget error reports the size and the setting to raise", () =>
        {
            var s = Snapshot(); s.Options.MaxRequestChars = 3000; var api = new FakePage(s.Page);
            Rejects("MaxRequestChars", () => new AgentRunner(new ScriptedClient(s), new AgentCommitter(api)).RunAsync(s, "美化", null, CancellationToken.None).GetAwaiter().GetResult());
            Rejects("上限 3000", () => new AgentRunner(new ScriptedClient(s), new AgentCommitter(api)).RunAsync(s, "美化", null, CancellationToken.None).GetAwaiter().GetResult());
            Equal(0, api.Writes);
        });
        Test("agent budget defaults fit a 1M-token model and the serializer allows them", () =>
        {
            var defaults = AgentOptions.Parse(null);
            Equal(1000000, defaults.MaxRequestChars); Equal(200000, defaults.MaxPageChars);
            var capped = AgentOptions.Parse(XElement.Parse("<Agent><MaxRequestChars>9999999</MaxRequestChars><MaxPageChars>5000000</MaxPageChars></Agent>"));
            Equal(3000000, capped.MaxRequestChars); Equal(1000000, capped.MaxPageChars);
            var large = new string('字', 700000);
            Equal(large, (string)((IDictionary<string, object>)AgentChatClient.Parse(AgentJson.Serialize(new { text = large })))["text"]);
        });
        Test("AgentJson keeps HTML characters, prunes null fields and round-trips", () =>
        {
            const string html = "<b>A&B</b> 'q' \"x\" \\u003c 中文";
            var json = AgentJson.Serialize(new { text = html });
            True(json.Contains("<b>A&B</b> 'q'")); Equal(1, json.Split(new[] { "u003c" }, StringSplitOptions.None).Length - 1);
            Equal(html, (string)((IDictionary<string, object>)AgentChatClient.Parse(json))["text"]);
            // 作为字符串再包一层也不会变成 \\u003c。
            var nested = AgentJson.Serialize(new { content = json });
            True(nested.Contains("<b>A&B</b>")); Equal(json, (string)((IDictionary<string, object>)AgentChatClient.Parse(nested))["content"]);
            var result = AgentJson.ToolResult(new { ok = true, none = (string)null, tags = new string[0], blocks = new[] { new { id = "p1", parent_id = (string)null } } });
            Equal("{\"ok\":true,\"tags\":[],\"blocks\":[{\"id\":\"p1\"}]}", result);
        });
        Test("read_blocks gives raw runs only for paragraphs with inline formatting", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "纯文字 A&amp;B"), Paragraph("b", "有<b>加粗</b>的段落"))); var t = Tools(s);
            var read = AgentChatClient.Parse(Json(Invoke(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1", "p2" } })));
            var blocks = (object[])AiClient.Get(read, "blocks");
            True(AiClient.Get(blocks[0], "runs") == null); Equal("纯文字 A&B", (string)AiClient.Get(blocks[0], "text"));
            Equal("有<b>加粗</b>的段落", string.Concat(((object[])AiClient.Get(blocks[1], "runs")).Cast<string>()));
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
            Equal("读取段落 · 2 段", finished[1].Text);
            Equal("设置段落样式 · 一级标题 · 1 段", finished[2].Text);
            Equal("检查格式草稿", finished[3].Text);
            Equal("写回并验证", finished[4].Text);
        });
        Test("step descriptions summarize arguments, results and failures", () =>
        {
            Equal(("读取段落 · 2 段", AgentStepState.Done), AgentTools.DescribeStep("read_blocks", "{\"block_ids\":[\"a\",\"b\"]}", "{}"));
            Equal(("读取段落 · 1 段 · 跳过受保护 2 段", AgentStepState.Done),
                AgentTools.DescribeStep("read_blocks", "{\"block_ids\":[\"a\",\"b\",\"c\"]}", "{\"blocks\":[{}],\"skipped\":[{},{}]}"));
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
        Test("849-paragraph page rejects incomplete finish and resumes the same draft", () =>
        {
            var s = Snapshot(LongCoveragePage()); var api = new FakePage(s.Page);
            var t = new AgentTools(s, new AgentCommitter(api), CancellationToken.None);
            Equal(849, s.Blocks.Count); Equal(456, s.Blocks.Count(b => b.Editable));
            var overview = Invoke(t, "get_page_overview", new { });
            Equal(849, (int)ToolField(overview, "total")); Equal(100, (int)ToolField(overview, "next_offset"));
            Invoke(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = s.Blocks.Take(100).Where(b => b.Editable).Select(b => b.Id).ToArray() });
            Style(t, s); var revision = s.Revision;
            var refused = Invoke(t, "finish_edit", new { snapshot_id = s.SnapshotId, draft_revision = revision });
            Equal(false, (bool)ToolField(refused, "ok")); Equal(396, (int)ToolField(refused, "unread_count"));
            Equal(100, ((string[])ToolField(refused, "next_read_block_ids")).Length);
            Equal("p102", ((string[])ToolField(refused, "next_read_block_ids"))[0]);
            Equal(0, api.Writes); True(!s.Frozen && t.Report == null); Equal(revision, s.Revision); True(s.Blocks[0].Changed);
            while (t.UnreadCount > 0)
            {
                var batch = (string[])ToolField(Invoke(t, "get_pending_changes", new { snapshot_id = s.SnapshotId }), "next_read_block_ids");
                Invoke(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = batch });
            }
            Invoke(t, "finish_edit", new { snapshot_id = s.SnapshotId, draft_revision = s.Revision });
            Equal("Verified", t.Report.Status); Equal(1, api.Writes); Equal(1, t.Report.Applied); Equal(0, t.Report.UnreadCount);
            Equal(0, (int)ToolField(t.Report.ToToolResult(), "unread_count"));
        });
        Test("coverage read batches respect 100-item pagination boundaries", () =>
        {
            foreach (var count in new[] { 100, 101, 200, 201 })
            {
                var s = Snapshot(Page(Enumerable.Range(1, count).Select(i => Paragraph("a" + i, "正文 " + i)).ToArray())); var t = Tools(s);
                var seen = 0;
                for (var offset = 0; offset < count; offset += 100)
                {
                    var overview = Invoke(t, "get_page_overview", new { offset });
                    Equal(Math.Min(100, count - offset), ((object[])ToolField(overview, "blocks")).Length);
                    Equal(offset + 100 < count ? (int?)(offset + 100) : null, (int?)ToolField(overview, "next_offset"));
                }
                while (t.UnreadCount > 0)
                {
                    var pending = Invoke(t, "get_pending_changes", new { snapshot_id = s.SnapshotId });
                    var ids = (string[])ToolField(pending, "next_read_block_ids");
                    Equal(Math.Min(100, count - seen), ids.Length);
                    Invoke(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = ids }); seen += ids.Length;
                }
                Equal(count, seen);
                Invoke(t, "finish_edit", new { snapshot_id = s.SnapshotId, draft_revision = s.Revision });
                Equal("NoChange", t.Report.Status); Equal(0, t.Report.UnreadCount);
            }
        });
        Test("coverage includes candidate code but skips protected content and respects selection", () =>
        {
            var mono = Paragraph("m", "int value = 1;"); mono.SetAttributeValue("style", "font-family:Consolas");
            var page = Page(Paragraph("a", "正文"), mono, Paragraph("blank", ""),
                Paragraph("inline", "调用 <span style='font-family:Consolas'>Run()</span> 方法"), Paragraph("html", "<img src='x'>"));
            var s = Snapshot(page); var t = Tools(s);
            Equal(2, t.UnreadCount); Equal("unhighlighted_code", s.Blocks[1].ProtectedReason);
            Invoke(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" } });
            var refused = Invoke(t, "finish_edit", new { snapshot_id = s.SnapshotId, draft_revision = s.Revision });
            Equal("p2", ((string[])ToolField(refused, "next_read_block_ids")).Single());
            Invoke(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p2" } }); Equal(0, t.UnreadCount);
            foreach (var id in new[] { "a", "m" })
            {
                var selected = new AgentPageSnapshot(page.ToString(), new HashSet<string> { id }, new AgentOptions()); var selectedTools = Tools(selected);
                Equal(1, selectedTools.UnreadCount);
                Invoke(selectedTools, "read_blocks", new { snapshot_id = selected.SnapshotId, block_ids = new[] { "p1" } });
                Invoke(selectedTools, "finish_edit", new { snapshot_id = selected.SnapshotId, draft_revision = selected.Revision });
                Equal("NoChange", selectedTools.Report.Status); Equal(0, selectedTools.Report.UnreadCount);
            }
            var disabled = new AgentPageSnapshot(page.ToString(), null, new AgentOptions { EnableCodeHighlight = false });
            Equal(1, Tools(disabled).UnreadCount);
        });
        Test("coverage excludes generated and converted paragraphs and preserves layout undo", () =>
        {
            var s = Snapshot(Page(Paragraph("h", "标题"), Paragraph("r1", "名称|内容"), Paragraph("r2", "甲|乙"), Paragraph("tail", "结尾")));
            var api = new FakePage(s.Page); var committer = new AgentCommitter(api); var t = new AgentTools(s, committer, CancellationToken.None);
            Invoke(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1", "p2", "p3" } });
            Table(t, s, "pipe", "p2", "p3"); Insert(t, s, "p1", "摘要");
            var pending = Invoke(t, "get_pending_changes", new { snapshot_id = s.SnapshotId });
            Equal("p4", ((string[])ToolField(pending, "unread")).Single()); Equal(1, (int)ToolField(pending, "unread_count"));
            Invoke(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p4" } });
            Invoke(t, "finish_edit", new { snapshot_id = s.SnapshotId, draft_revision = s.Revision });
            Equal("Verified", t.Report.Status); Equal(1, t.Report.Inserted); Equal(1, t.Report.TextTables); Equal(0, t.Report.UnreadCount);
            var undo = committer.Undo(s.PageId, t.Report, s.Options, CancellationToken.None);
            Equal("Verified", undo.Status); Equal(Texts(s.Page), Texts(api.Page)); Equal(0, undo.UnreadCount);
        });
        Test("coverage survives markdown deletion and follows moved paragraph order", () =>
        {
            var s = Snapshot(Page(Paragraph("h", "# 标题"), Paragraph("f1", "```markdown"), Paragraph("a", "正文"), Paragraph("f2", "```"), Paragraph("tail", "结尾")));
            var t = Tools(s);
            Invoke(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1", "p2", "p3", "p4" } });
            Cleanup(t, s, "p1", "p2", "p3", "p4");
            Equal(1, t.UnreadCount); Equal("p5", ((string[])ToolField(Invoke(t, "get_pending_changes", new { snapshot_id = s.SnapshotId }), "unread")).Single());
            // 被整体移动的父段下有未读取子段；下一批按新的草稿顺序返回。
            var parent = Paragraph("parent", "父段"); parent.Add(new XElement(One + "OEChildren", Paragraph("child", "子段")));
            var moved = Snapshot(Page(parent, Paragraph("b", "后段"), Paragraph("c", "尾段"))); var mt = Tools(moved);
            Invoke(mt, "read_blocks", new { snapshot_id = moved.SnapshotId, block_ids = new[] { "p1", "p4" } });
            Move(mt, moved, new[] { "p1" }, "p4", "after");
            Equal("p3,p2", string.Join(",", (string[])ToolField(Invoke(mt, "get_pending_changes", new { snapshot_id = moved.SnapshotId }), "next_read_block_ids")));
        });
        Test("coverage does not enlarge a local edit or repeat an unchanged style", () =>
        {
            var s = Snapshot(); var api = new FakePage(s.Page); var committer = new AgentCommitter(api); var t = new AgentTools(s, committer, CancellationToken.None);
            var untouched = AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "b"), api.Page);
            Read(t, s); Style(t, s); var revision = s.Revision; Style(t, s); Equal(revision, s.Revision);
            Invoke(t, "finish_edit", new { snapshot_id = s.SnapshotId, draft_revision = s.Revision });
            Equal(1, t.Report.Applied); Equal(untouched, AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "b"), api.Page));
            Equal("Verified", committer.Undo(s.PageId, t.Report, s.Options, CancellationToken.None).Status);
        });
        Test("coverage cannot bypass revision checks or accept a model partial-commit flag", () =>
        {
            var s = Snapshot(); var t = Tools(s);
            Invoke(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" } }); Style(t, s);
            Rejects("修订号过期", () => Invoke(t, "finish_edit", new { snapshot_id = s.SnapshotId, draft_revision = 0 }));
            Throws(() => Invoke(t, "finish_edit", new { snapshot_id = s.SnapshotId, draft_revision = s.Revision, allow_partial = true }));
            True(!s.Frozen && t.Report == null); Equal(1, t.UnreadCount);
        });
        Test("runner recovers from premature finish and completes remaining read batches", () =>
        {
            var s = Snapshot(LongCoveragePage()); var api = new FakePage(s.Page); var model = new CoverageClient(s);
            var r = new AgentRunner(model, new AgentCommitter(api)).RunAsync(s, "只突出第一段", null, CancellationToken.None).GetAwaiter().GetResult();
            True(model.SawRejection); Equal("Verified", r.Status); Equal(0, r.UnreadCount); Equal(1, api.Writes); Equal(1, r.Applied);
            True(s.Blocks.Where(b => b.Editable).All(b => b.Read));
        });
        Test("cancelling after a refused finish never writes the retained draft", () =>
        {
            var s = Snapshot(LongCoveragePage()); var api = new FakePage(s.Page);
            using (var cancel = new CancellationTokenSource())
            {
                var model = new CoverageClient(s, cancel.Cancel);
                Throws<OperationCanceledException>(() => new AgentRunner(model, new AgentCommitter(api)).RunAsync(s, "美化", null, cancel.Token).GetAwaiter().GetResult());
                True(model.SawRejection); Equal(0, api.Writes); True(!s.Frozen);
            }
        });
        Test("readback format mismatch logs only object identity and a fixed reason", () =>
        {
            const string privateText = "NOTE_PRIVATE_do_not_log";
            var objectId = "verify-" + Guid.NewGuid().ToString("N");
            var s = Snapshot(Page(Paragraph(objectId, privateText))); var api = new FakePage(s.Page); var t = new AgentTools(s, new AgentCommitter(api), CancellationToken.None);
            Read(t, s); Style(t, s);
            api.AfterSave = () => AgentCommitter.Find(api.Page, objectId).SetAttributeValue("alignment", "right");
            var oldLog = File.Exists(AddInLog.LogPath) ? File.ReadAllText(AddInLog.LogPath) : "";
            Invoke(t, "finish_edit", new { snapshot_id = s.SnapshotId, draft_revision = s.Revision });
            Equal("PartiallyApplied", t.Report.Status); Equal(1, t.Report.Unverified); Equal(0, t.Report.Applied); True(!t.Report.CanUndo);
            var log = File.ReadAllText(AddInLog.LogPath); var added = log.StartsWith(oldLog, StringComparison.Ordinal) ? log.Substring(oldLog.Length) : log;
            True(added.Contains("id=" + objectId)); True(added.Contains("reason=semantic_format_mismatch")); True(!added.Contains(privateText));
            Equal("Agent 核验失败：kind=paragraph id=- reason=missing_object。", AgentCommitter.VerificationDiagnostic("paragraph", "正文\nApiKey=private", "missing_object"));
        });
        Test("runner turn budget discards uncommitted draft", () =>
        {
            var s = Snapshot(); s.Options.MaxTurns = 3; var api = new FakePage(s.Page);
            // 最后一轮只能 finish_edit，脚本仍去设样式，被拒绝后报轮数用完。
            Rejects("MaxTurns", () => new AgentRunner(new ScriptedClient(s), new AgentCommitter(api)).RunAsync(s, "美化", null, CancellationToken.None).GetAwaiter().GetResult());
            Equal(0, api.Writes);
        });
        Test("runner wraps up and commits the draft on the last turn", () =>
        {
            var s = Snapshot(); s.Options.MaxTurns = 4; var api = new FakePage(s.Page); var model = new WrapUpClient(s);
            var r = new AgentRunner(model, new AgentCommitter(api)).RunAsync(s, "美化", null, CancellationToken.None).GetAwaiter().GetResult();
            Equal("PartiallyApplied", r.Status); Equal(1, r.UnreadCount); Equal(1, api.Writes);
            Equal(1, model.ToolCounts.Last()); True(model.ToolCounts.Take(3).All(n => n > 1));
            True(model.LastJson.Contains("只剩 2 轮")); True(model.LastJson.Contains("这是最后一轮"));
            True(r.Message.Contains("MaxTurns"));
            True(r.Message.Contains("仍有 1 段未读取")); Equal(0, r.Conflicts); Equal(0, r.Unverified);
            var undo = new AgentCommitter(api).Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal("Verified", undo.Status); Equal(0, undo.UnreadCount); Equal(Texts(s.Page), Texts(api.Page));
        });
        Test("last-turn complete coverage remains Verified and an empty draft stays NoChange", () =>
        {
            foreach (var change in new[] { true, false })
            {
                var s = Snapshot(); s.Options.MaxTurns = 4; var api = new FakePage(s.Page);
                var r = new AgentRunner(new WrapUpClient(s, change, change), new AgentCommitter(api)).RunAsync(s, "美化", null, CancellationToken.None).GetAwaiter().GetResult();
                Equal(change ? "Verified" : "NoChange", r.Status); Equal(change ? 0 : 1, r.UnreadCount); Equal(change ? 1 : 0, api.Writes);
            }
        });
        Test("last-turn unread count preserves conflict and unknown-write outcomes", () =>
        {
            foreach (var conflict in new[] { true, false })
            {
                var s = Snapshot(); s.Options.MaxTurns = 4; var api = new FakePage(s.Page);
                if (conflict) AgentCommitter.Find(api.Page, "a").Element(One + "T").Value = "用户后续编辑";
                else api.FailReadAfterSave = true;
                var r = new AgentRunner(new WrapUpClient(s), new AgentCommitter(api)).RunAsync(s, "美化", null, CancellationToken.None).GetAwaiter().GetResult();
                Equal(conflict ? "NoChange" : "CommitOutcomeUnknown", r.Status); Equal(1, r.UnreadCount);
                Equal(conflict ? 1 : 0, r.Conflicts); Equal(conflict ? 0 : 1, api.Writes); True(!r.CanUndo);
                if (conflict) Equal("用户后续编辑", new AgentRichText(AgentCommitter.Find(api.Page, "a")).Text);
            }
        });
        Test("last turn text-only reply is NoChange, not an error", () =>
        {
            var s = Snapshot(); s.Options.MaxTurns = 2; var api = new FakePage(s.Page);
            var r = new AgentRunner(new TextOnlyClient(), new AgentCommitter(api)).RunAsync(s, "美化", null, CancellationToken.None).GetAwaiter().GetResult();
            Equal("NoChange", r.Status); Equal(0, api.Writes); Equal(2, r.UnreadCount); True(r.Message.Contains("最大轮数"));
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
        Test("HTTP stream progress shows finished reasoning sentences and the tool being prepared", () =>
        {
            var handler = new StubHttp("data: {\"choices\":[{\"index\":0,\"delta\":{\"reasoning_content\":\"**先读取段落**\\n\\n再确定标题层\"}}]}\n\n" +
                "data: {\"choices\":[{\"index\":0,\"delta\":{\"reasoning_content\":\"级\"}}]}\n\n" +
                "data: {\"choices\":[{\"index\":0,\"delta\":{\"tool_calls\":[{\"index\":0,\"id\":\"x\",\"function\":{\"name\":\"read_blocks\",\"arguments\":\"{}\"}}]},\"finish_reason\":\"tool_calls\"}]}\n\ndata: [DONE]\n\n", "text/event-stream");
            var c = AiConfigStore.Parse(XElement.Parse("<AiConfig><ApiUrl>https://test.invalid</ApiUrl><ApiKey>test</ApiKey></AiConfig>"));
            var log = new ProgressLog();
            using (var http = new HttpClient(handler))
                new AgentChatClient(c, "test", "medium", http).CompleteAsync(new List<object>(), new object[0], log, CancellationToken.None).GetAwaiter().GetResult();
            // 第一段立即报告，只显示说完的一句、去掉加粗；后面的被限频吞掉，收完再报一次最终状态，这时最后一句也算说完。
            Equal("模型正在思考…", log.Items[0].Status); Equal("先读取段落", log.Items[0].Thinking);
            Equal("模型正在准备：读取段落", log.Items.Last().Status); Equal("先读取段落\n再确定标题层级", log.Items.Last().Thinking);
        });
        Test("HTTP stream progress leaves the excerpt unchanged until a sentence is finished", () =>
        {
            var handler = new StubHttp("data: {\"choices\":[{\"index\":0,\"delta\":{\"reasoning_content\":\"先读取页面概\"}}]}\n\n" +
                "data: {\"choices\":[{\"index\":0,\"delta\":{\"content\":\"好的。\"},\"finish_reason\":\"stop\"}]}\n\ndata: [DONE]\n\n", "text/event-stream");
            var c = AiConfigStore.Parse(XElement.Parse("<AiConfig><ApiUrl>https://test.invalid</ApiUrl><ApiKey>test</ApiKey></AiConfig>"));
            var log = new ProgressLog();
            using (var http = new HttpClient(handler))
                new AgentChatClient(c, "test", "medium", http).CompleteAsync(new List<object>(), new object[0], log, CancellationToken.None).GetAwaiter().GetResult();
            // 半句不显示（null 表示不变）；回复只有语气词时退回思考，思考已经结束，最后一句也算说完。
            Equal(null, log.Items[0].Thinking);
            Equal("模型正在回复…", log.Items.Last().Status); Equal("先读取页面概", log.Items.Last().Thinking);
        });
        Test("reasoning excerpt names tools and blocks in Chinese and drops English or mixed sentences", () =>
        {
            var reply = new AgentReply();
            reply.ReasoningBuffer.Append("The user wants a cleaner page. 先调用get_page_overview看看结构。\n" +
                "p3 到 p5 是标题，用 set_paragraph_style 设为 heading2。\n需要 check 一下 draft_revision。\n");
            Equal("先调用「读取页面概况」看看结构。\n第 3–5 段是标题，用「设置段落样式」设为二级标题。", AgentChatClient.DescribeStream(reply).Thinking);
            // 收完了（有结束原因），最后一句没有句号也显示；前后的「第」「段落」不重复。
            var ids = new AgentReply { FinishReason = "stop" };
            ids.ReasoningBuffer.Append("把 t2 和 i1 移到 n1 后面。第 p4 段和 p7 段落的 style 不同");
            Equal("把第 2 个表格和第 1 张图片移到新插入的第 1 段后面。\n第 4 段和第 7 段的样式不同", AgentChatClient.DescribeStream(ids).Thinking);
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
        Test("HTTP failures distinguish gateway, authentication, URL, limits and request parameters", () =>
        {
            var cases = new[] { (400, "请求参数"), (401, "ApiKey"), (403, "访问权限"), (404, "ApiUrl"), (422, "兼容选项"),
                (429, "限流"), (500, "服务暂时不可用"), (502, "服务暂时不可用"), (503, "服务暂时不可用"), (504, "上游服务超时") };
            foreach (var (status, hint) in cases)
            {
                var message = HttpFailure(status, "<html>PRIVATE_PAGE_CONTENT SECRET_API_KEY</html>", "text/html");
                True(message.Contains("HTTP " + status) && message.Contains(hint) && message.Contains("未提交草稿"));
                True(!message.Contains("PRIVATE_PAGE_CONTENT") && !message.Contains("SECRET_API_KEY"));
                if (status >= 500) True(!message.Contains("工具调用"));
            }
        });
        Test("HTTP 503 model_not_found is recognized without echoing the upstream message", () =>
        {
            foreach (var status in new[] { 404, 503 })
            {
                var message = HttpFailure(status, "{\"error\":{\"code\":\"MODEL_NOT_FOUND\",\"message\":\"PRIVATE_PAGE_CONTENT SECRET_API_KEY\"}}");
                True(message.Contains("模型不存在或未开通") && message.Contains("HTTP " + status));
                True(!message.Contains("PRIVATE_PAGE_CONTENT") && !message.Contains("SECRET_API_KEY"));
                True(!message.Contains("工具调用") && !message.Contains("服务暂时不可用"));
            }
        });
        Test("malformed, oversized and unrecognized HTTP error bodies use the status safely", () =>
        {
            foreach (var body in new[] { "", "{invalid PRIVATE_PAGE_CONTENT", new string('x', 20000),
                "{\"error\":{\"code\":\"PRIVATE_PAGE_CONTENT\",\"message\":\"model_not_found\"}}", "{\"error\":{\"code\":{\"private\":\"data\"}}}" })
            {
                var message = HttpFailure(503, body);
                True(message.Contains("服务暂时不可用") && !message.Contains("PRIVATE_PAGE_CONTENT") && !message.Contains("模型不存在"));
            }
        });
        Test("HTTP 503 stops the agent without retrying or writing a prepared draft", () =>
        {
            var s = Prepared(); var api = new FakePage(s.Page); var before = api.Page.ToString();
            var handler = new StubHttp("{\"error\":{\"code\":\"upstream_unavailable\"}}", "application/json") { Status = HttpStatusCode.ServiceUnavailable };
            var config = AiConfigStore.Parse(XElement.Parse("<AiConfig><ApiUrl>https://test.invalid</ApiUrl><ApiKey>test</ApiKey></AiConfig>"));
            using (var http = new HttpClient(handler))
                Rejects("服务暂时不可用", () => new AgentRunner(new AgentChatClient(config, "test", "none", http), new AgentCommitter(api))
                    .RunAsync(s, "排版", null, CancellationToken.None).GetAwaiter().GetResult());
            Equal(1, handler.Requests); Equal(0, api.Writes); Equal(before, api.Page.ToString()); True(!s.Frozen);
        });
        Test("HTTP error diagnostics have a deadline and respect user cancellation", () =>
        {
            foreach (var cancel in new[] { false, true })
            {
                var content = new WaitingErrorContent();
                var handler = new StubHttp("", "application/json") { Status = HttpStatusCode.ServiceUnavailable, ResponseContent = content };
                var config = AiConfigStore.Parse(XElement.Parse("<AiConfig><ApiUrl>https://test.invalid</ApiUrl><ApiKey>test</ApiKey></AiConfig>"));
                using (var cancellation = new CancellationTokenSource())
                using (var http = new HttpClient(handler))
                {
                    if (cancel) cancellation.CancelAfter(100);
                    try
                    {
                        new AgentChatClient(config, "test", "none", http).CompleteAsync(new List<object>(), new object[0], null, cancellation.Token).GetAwaiter().GetResult();
                        throw new Exception("Expected HTTP failure");
                    }
                    catch (OperationCanceledException) { True(cancel && cancellation.IsCancellationRequested); }
                    catch (AiException ex) { True(!cancel && ex.Message.Contains("HTTP 503") && ex.Message.Contains("服务暂时不可用")); }
                    Equal(1, handler.Requests); True(content.Disposed);
                }
            }
        });
        Test("configuration absent Agent node preserves defaults", () =>
        {
            var c = AiConfigStore.Parse(XElement.Parse("<AiConfig><Agent><MaxTurns>999</MaxTurns><ReplayReasoning>false</ReplayReasoning></Agent></AiConfig>"));
            Equal(60, c.Agent.MaxTurns); True(!c.Agent.ReplayReasoning); True(c.Agent.EnableMixedOutlines);
            Equal(24, AgentOptions.Parse(null).MaxTurns); Equal(96, AgentOptions.Parse(null).MaxToolCalls);
        });
        Test("Agent default request reads legacy configurations and validates complete templates", () =>
        {
            Equal(AgentOptions.DefaultRequestText, AiConfigStore.Parse(XElement.Parse("<AiConfig/>")).Agent.DefaultRequest);
            foreach (var text in new[] { null, "", " \r\n\t ", new string('字', AgentOptions.MaxRequestLength + 1) })
            {
                var element = new XElement("Agent", text == null ? null : new XElement("DefaultRequest", text));
                var options = AgentOptions.Parse(element);
                Equal(AgentOptions.DefaultRequestText, options.DefaultRequest); True(options.IsDefault);
            }
            foreach (var text in new[] { "字", new string('字', AgentOptions.MaxRequestLength), "  正文 <11> & 标题\r\n第二行\r第三行  " })
            {
                var normalized = AgentOptions.NormalizeRequest(text);
                var options = AgentOptions.Parse(new XElement("Agent", new XElement("DefaultRequest", text)));
                Equal(normalized, options.DefaultRequest); True(!options.IsDefault);
                Equal(normalized, AgentOptions.Parse(new XElement("Agent", options.ToElements())).DefaultRequest);
            }
            Equal(AgentOptions.DefaultRequestText, (string)AiConfigStore.BuildDefaultDocument().Root.Element("Agent").Element("DefaultRequest"));
        });
        Test("Agent default request is saved in place while retaining unknown nodes and comments", () =>
        {
            var document = XDocument.Parse("<AiConfig><!-- 保留 --><Custom>x</Custom></AiConfig>");
            AiConfigStore.Apply(document, AiConfigStore.Default);
            True(document.Root.Element("Agent") == null);
            var config = AiConfigStore.Parse(new XElement("AiConfig", new XElement("Agent",
                new XElement("DefaultRequest", "  正文 <11> & 标题\r\n第二行  "))));
            AiConfigStore.Apply(document, config);
            Equal(config.Agent.DefaultRequest, AiConfigStore.Parse(document.Root).Agent.DefaultRequest);
            Equal("x", (string)document.Root.Element("Custom"));
            True(document.ToString().Contains("<!-- 保留 -->"));
            document.Root.Element("Agent").Add(new XComment("内部注释"), new XElement("Unknown", "keep"));
            AiConfigStore.Apply(document, AiConfigStore.Default);
            Equal(AgentOptions.DefaultRequestText, (string)document.Root.Element("Agent").Element("DefaultRequest"));
            Equal("keep", (string)document.Root.Element("Agent").Element("Unknown"));
            True(document.ToString().Contains("<!--内部注释-->"));
        });
        Test("AI settings default request supports normalization, validation and restoring without changing other options", () =>
        {
            var config = AiConfigStore.Parse(XElement.Parse("<AiConfig><Agent><DefaultRequest>我的模板</DefaultRequest><MaxTurns>30</MaxTurns><EnableMoves>false</EnableMoves></Agent></AiConfig>"));
            var window = new OneNoteCodeHelper.Views.AiSettingsWindow(config, null, IntPtr.Zero);
            try
            {
                bool Dirty() => (bool)window.GetType().GetProperty("IsDirty", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(window);
                True(!Dirty()); Equal("我的模板", window.AgentDefaultRequestBox.Text);
                Equal(AgentOptions.MaxRequestLength, window.AgentDefaultRequestBox.MaxLength);
                window.AgentDefaultRequestBox.Text = "  正文 <11> & 标题\r\n第二行  ";
                True(Dirty()); True(window.TryBuildConfig(out var built));
                Equal("正文 <11> & 标题\n第二行", built.Agent.DefaultRequest);
                foreach (var text in new[] { " \r\n ", new string('字', AgentOptions.MaxRequestLength + 1) })
                {
                    window.AgentDefaultRequestBox.Text = text;
                    True(!window.TryBuildConfig(out _));
                    Equal("默认需求要填 1–8000 字的内容。", window.StatusText.Text);
                    Equal(System.Windows.Visibility.Visible, window.AgentPage.Visibility);
                }
                window.AgentDefaultRequestBox.Text = new string('字', AgentOptions.MaxRequestLength);
                True(window.TryBuildConfig(out built)); Equal(AgentOptions.MaxRequestLength, built.Agent.DefaultRequest.Length);
                window.RestoreAgentDefaultRequestButton.RaiseEvent(new System.Windows.RoutedEventArgs(System.Windows.Controls.Button.ClickEvent));
                True(Dirty()); True(window.TryBuildConfig(out built));
                Equal(AgentOptions.DefaultRequestText, built.Agent.DefaultRequest);
                Equal(30, built.Agent.MaxTurns); True(!built.Agent.EnableMoves);
                Equal("我的模板", config.Agent.DefaultRequest);
                window.AgentDefaultRequestBox.Text = "我的模板"; True(!Dirty());
            }
            finally { window.CloseWithoutSaving(); }
        });
        Test("Agent request preserves edits across choices and refreshes, then reopens with the latest saved template", () =>
        {
            AiConfig Config(string request) => AiConfigStore.Parse(new XElement("AiConfig", new XElement("Agent", new XElement("DefaultRequest", request))));
            var page = Page(Paragraph("a", "正文"));
            var config = Config("正文 11 磅\n标题加粗");
            var latest = Config("第四版模板");
            var window = new OneNoteCodeHelper.Views.AgentWindow(new FakePage(page), "page", page.ToString(), config, new AddInSettings(), IntPtr.Zero);
            try
            {
                Equal(config.Agent.DefaultRequest, AgentOptions.NormalizeRequest(window.RequestText.Text));
                Equal(AgentOptions.MaxRequestLength, window.RequestText.MaxLength);
                window.ApplyConfig(Config("新模板")); Equal("新模板", window.RequestText.Text);
                window.RequestText.Text = "本次只整理标题";
                window.FunctionPicker.SelectedIndex = 1;
                window.ApplyConfig(Config("第三版模板"));
                window.FunctionPicker.SelectedIndex = 0;
                Equal("本次只整理标题", window.RequestText.Text);
                window.RequestText.Clear(); window.ApplyConfig(latest); Equal("", window.RequestText.Text);
                window.FunctionPicker.SelectedIndex = 1;
                window.FunctionPicker.SelectedIndex = 0;
                Equal("", window.RequestText.Text);
                Equal("正文 11 磅\n标题加粗", config.Agent.DefaultRequest);
            }
            finally { window.Close(); }
            // 保存到独立临时文件，验证重开使用持久化的最新模板，不碰本机 AI 配置。
            var path = Path.Combine(Path.GetTempPath(), "onenote-agent-request-" + Guid.NewGuid().ToString("N") + ".xml");
            try
            {
                AiConfigStore.Save(latest, path);
                var saved = AiConfigStore.Parse(XDocument.Load(path).Root);
                var reopened = new OneNoteCodeHelper.Views.AgentWindow(new FakePage(page), "page", page.ToString(), saved, new AddInSettings(), IntPtr.Zero);
                try { Equal(latest.Agent.DefaultRequest, AgentOptions.NormalizeRequest(reopened.RequestText.Text)); }
                finally { reopened.Close(); }
            }
            finally { File.Delete(path); }
        });
        Test("Agent sends only the current request with its complete system protocol for both page and selection scopes", () =>
        {
            var page = Page(Paragraph("a", "正文"));
            var config = AiConfigStore.Parse(XElement.Parse("<AiConfig><Agent><DefaultRequest>模板不应追加</DefaultRequest></Agent></AiConfig>"));
            var window = new OneNoteCodeHelper.Views.AgentWindow(new FakePage(page), "page", page.ToString(), config, new AddInSettings(), IntPtr.Zero);
            try
            {
                window.RequestText.Text = "标题加粗\n正文 11 磅";
                var request = AgentOptions.NormalizeRequest(window.RequestText.Text);
                foreach (var selection in new ISet<string>[] { null, new HashSet<string> { "a" } })
                {
                    var snapshot = new AgentPageSnapshot(page.ToString(), selection, config.Agent);
                    var api = new FakePage(page); var client = new TruncatingClient(new ScriptedClient(snapshot));
                    new AgentRunner(client, new AgentCommitter(api)).RunAsync(snapshot, request, null, CancellationToken.None).GetAwaiter().GetResult();
                    var sent = (object[])AgentChatClient.Serializer().DeserializeObject(client.Sent[0]);
                    Equal(2, sent.Length);
                    var system = (Dictionary<string, object>)sent[0]; var user = (Dictionary<string, object>)sent[1];
                    Equal("system", system["role"]); Equal("user", user["role"]); Equal(request, user["content"]);
                    True(((string)system["content"]).StartsWith(AgentRunner.SystemPrompt(Tools(snapshot)), StringComparison.Ordinal));
                    True(!client.Sent[0].Contains(config.Agent.DefaultRequest));
                }
            }
            finally { window.Close(); }
        });
        Test("AI settings window shows every option and builds the same configuration back", () =>
        {
            var c = AiConfigStore.Parse(XElement.Parse("<AiConfig><ApiUrl>https://gw.test/v1</ApiUrl><ApiKey>k</ApiKey><TimeoutSeconds>90</TimeoutSeconds><MaxTokens>0</MaxTokens>" +
                "<Agent><MaxTurns>30</MaxTurns><EnableMoves>false</EnableMoves><SendThinking>false</SendThinking><FontFamily>Arial</FontFamily></Agent>" +
                "<Models><Model id='m1'/><Model id='m2'/></Models><Functions><Function name='甲' removeExtraBlankLines='true'><Prompt>第一行\n第二行</Prompt></Function><Function name='乙'><Prompt>p</Prompt></Function></Functions></AiConfig>"));
            var window = new OneNoteCodeHelper.Views.AiSettingsWindow(c, null, IntPtr.Zero);
            Equal(AgentOptions.Switches.Length, window.FormatSwitchesPanel.Children.Count + window.StructureSwitchesPanel.Children.Count + window.CompatSwitchesPanel.Children.Count);
            Equal(AgentOptions.Numbers.Length, window.AgentNumbersPanel.Children.Count);
            var ok = window.TryBuildConfig(out var built);
            Equal("", window.StatusText.Text); True(ok);
            Equal("https://gw.test/v1|k|90|0", $"{built.ApiUrl}|{built.ApiKey}|{built.TimeoutSeconds}|{built.MaxTokens}");
            Equal("m1,m2", string.Join(",", built.Models.Select(m => m.Id)));
            Equal("甲=第一行\n第二行=True|乙=p=False", string.Join("|", built.Functions.Select(f => f.Name + "=" + f.Prompt + "=" + f.RemoveExtraBlankLines)));
            foreach (var option in AgentOptions.Numbers) Equal(option.Get(c.Agent), option.Get(built.Agent));
            foreach (var option in AgentOptions.Switches) Equal(option.Get(c.Agent), option.Get(built.Agent));
            Equal("Arial", built.Agent.FontFamily);

            string Rejected() => window.TryBuildConfig(out _) ? "accepted" : window.StatusText.Text;
            window.ApiUrlBox.Text = "ftp://gw.test";
            Equal("接口地址要填以 http:// 或 https:// 开头的完整地址。", Rejected());
            window.ApiUrlBox.Text = "https://gw.test/v1";
            window.TimeoutBox.Text = "5";
            Equal("请求超时要填 10–3600 之间的整数。", Rejected());
            window.TimeoutBox.Text = "90";
            window.FunctionList.SelectedIndex = 1;
            Equal("乙|p|False", $"{window.FunctionNameBox.Text}|{window.FunctionPromptBox.Text}|{window.FunctionBlankLinesBox.IsChecked}");
            window.FunctionNameBox.Text = " 甲 ";
            Equal("有两个文字功能都叫「甲」，改一个名字。", Rejected());
            window.FunctionNameBox.Text = AiConfigStore.AgentFunctionName;
            Equal("「" + AiConfigStore.AgentFunctionName + "」是 Agent 自己的名字，文字功能换一个名字。", Rejected());
        });
        Test("Agent window has the AI settings entry in its header", () =>
        {
            var page = Page(Paragraph("a", "正文"));
            var window = new OneNoteCodeHelper.Views.AgentWindow(new FakePage(page), "page", page.ToString(), AiConfigStore.Default, new AddInSettings(), IntPtr.Zero);
            True(window.SettingsButton.IsEnabled);
            True(window.SettingsIcon.Source != null);
            window.Close();
        });
        Test("Insert code window has its header icon and previews the pasted code", () =>
        {
            var window = new OneNoteCodeHelper.Views.InsertCodeWindow(new PageEditor(null), new AddInSettings { LanguageId = "python" }, IntPtr.Zero);
            try
            {
                True(window.AppIcon.Source != null);
                True(!window.InsertButton.IsEnabled);
                window.CodeBox.Text = "def f():\n    return 1";
                window.UpdatePreview();
                True(window.Preview.Document != null);
                True(window.InsertButton.IsEnabled);
                Equal("2 行，将按 Python 高亮。", window.StatusText.Text);
            }
            finally { window.Close(); }
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
            Rejects("p5（highlighted_code）", () => Invoke(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p5" } }));
            Rejects("p9 不存在", () => Invoke(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1", "p9" } }));
            // 受保护段落夹在范围里时跳过并说明原因，其余照常读取。
            var mixed = AgentChatClient.Parse(Json(Invoke(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1", "p4", "p5" } })));
            Equal("p1", (string)AiClient.Get(((object[])AiClient.Get(mixed, "blocks")).Single(), "id"));
            Equal("p4:protected_code,p5:highlighted_code", string.Join(",", ((object[])AiClient.Get(mixed, "skipped")).Select(x => AiClient.Get(x, "id") + ":" + AiClient.Get(x, "reason"))));
            True(s.Blocks[0].Read && !s.Blocks[3].Read && !s.Blocks[4].Read);
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
        Test("markdown restores separated numbered lists with their own starts, rich text and undo", () =>
        {
            foreach (var split in new[] { false, true })
            {
                var lines = new List<XElement> { Paragraph("h1", "核心改造") };
                lines.AddRange(Enumerable.Range(1, 8).Select(i => Paragraph("a" + i, i + ". <b>改造</b><a href='https://example.com/" + i + "'>渠道</a>")));
                lines.AddRange(new[] { Paragraph("h2", "YAML 字段建议"), Listed("bullet", "appId", "2"), Paragraph("body", "注意事项"), Paragraph("h3", "验证重点") });
                lines.AddRange(Enumerable.Range(1, 5).Select(i => Paragraph("b" + i, i + ". 验证渠道")));
                var s = Snapshot(Page(lines.ToArray())); var t = Tools(s); Read(t, s);
                var ids = s.Blocks.Where(b => b.Text.Length > 1 && char.IsDigit(b.Text[0])).Select(b => b.Id).ToArray();
                Cleanup(t, s, ids); Cleanup(t, s, ids); // 第二次清理没有标记，起点仍必须保留。
                if (split) foreach (var id in ids) NumberList(t, s, id);
                else NumberList(t, s, ids.Reverse().ToArray()); // 模型传参顺序不能改变分组。
                var revision = s.Revision; NumberList(t, s, ids); Equal(revision, s.Revision);
                var expected = "1.,2.,3.,4.,5.,6.,7.,8.,1.,2.,3.,4.,5.";
                var api = new FakePage(s.Page); api.AfterSave = () => RenderNumbering(api.Page);
                var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
                Equal("Verified", r.Status); Equal(13, r.Applied); Equal(13, r.MarkdownMarks);
                Equal(expected, string.Join(",", api.Page.Descendants(One + "Number").Select(n => (string)n.Attribute("text"))));
                Equal(2, api.Page.Descendants(One + "Number").Count(n => n.Attribute("restartNumberingAt") != null));
                Equal("1", (string)AgentCommitter.Find(api.Page, "b1").Descendants(One + "Number").Single().Attribute("restartNumberingAt"));
                True(AgentCommitter.Find(api.Page, "a1").Element(One + "T").Value.Contains("<b>改造</b>"));
                True(AgentCommitter.Find(api.Page, "a1").Element(One + "T").Value.Contains("https://example.com/1"));
                Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status);
                Equal(Texts(s.Page), Texts(api.Page)); True(!api.Page.Descendants(One + "Number").Any());
                Equal("bullet", AgentMarks.ListKind(AgentCommitter.Find(api.Page, "bullet")));
            }
        });
        Test("markdown all-one numbering is independent of batch order and preserves rich text through undo", () =>
        {
            var orders = new[] { new[] { "p1", "p2", "p3" }, new[] { "p1", "p3", "p2" }, new[] { "p2", "p1", "p3" },
                new[] { "p2", "p3", "p1" }, new[] { "p3", "p1", "p2" }, new[] { "p3", "p2", "p1" } };
            var batches = orders.Select(order => order.Select(id => new[] { id }).ToArray()).Concat(new[] {
                new[] { new[] { "p3" }, new[] { "p2", "p1" } }, new[] { new[] { "p3", "p2" }, new[] { "p1" } },
                new[] { new[] { "p3", "p1", "p2" } } });
            foreach (var batch in batches)
            {
                var s = Snapshot(Page(Paragraph("a", "1.&nbsp;<b>第一</b><a href='https://example.com'>项</a>"),
                    Paragraph("b", "1. 第二项"), Paragraph("c", "1. 第三项"))); var t = Tools(s); Read(t, s); Cleanup(t, s, "p1", "p2", "p3");
                foreach (var ids in batch) NumberList(t, s, ids);
                var revision = s.Revision; NumberList(t, s, "p3", "p1", "p2"); Equal(revision, s.Revision);
                var api = new FakePage(s.Page); api.AfterSave = () => RenderNumbering(api.Page);
                var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
                Equal("Verified", r.Status); Equal(3, r.Applied); Equal(3, r.MarkdownMarks);
                Equal("1.,2.,3.", string.Join(",", api.Page.Descendants(One + "Number").Select(n => (string)n.Attribute("text"))));
                Equal(1, api.Page.Descendants(One + "Number").Count(n => n.Attribute("restartNumberingAt") != null));
                var expected = Paragraph("a", "<b>第一</b><a href='https://example.com'>项</a>");
                Equal(new AgentRichText(expected).Signature(Page(expected), true), new AgentRichText(AgentCommitter.Find(api.Page, "a")).Signature(api.Page, true));
                Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status); Equal(Texts(s.Page), Texts(api.Page));
            }
        });
        Test("markdown numbering recalculates after removing or switching an item without changing native lists", () =>
        {
            foreach (var kind in new[] { "none", "bullet" })
            {
                var native = Paragraph("native", "原生编号"); AgentMarks.SetList(native, "number", 7);
                native.Descendants(One + "Number").Single().SetAttributeValue("fontSize", "11.0");
                native.Descendants(One + "Number").Single().SetAttributeValue("text", "7.");
                var s = Snapshot(Page(native, Paragraph("a", "1. 第一项"), Paragraph("b", "2. 第二项"), Paragraph("c", "3. 第三项")));
                var t = Tools(s); Read(t, s); Cleanup(t, s, "p2", "p3", "p4"); NumberList(t, s, "p2", "p3", "p4");
                var result = Json(Invoke(t, "set_list", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p2" }, list = kind }));
                Equal("p2,p3", string.Join(",", ((object[])((IDictionary<string, object>)AgentChatClient.Parse(result))["changed"]).Cast<string>()));
                Equal(kind, AgentMarks.ListKind(s.Blocks[1].Draft));
                Equal("2", (string)s.Blocks[2].Draft.Descendants(One + "Number").Single().Attribute("restartNumberingAt"));
                True(XNode.DeepEquals(s.Blocks[0].Original.Element(One + "List"), s.Blocks[0].Draft.Element(One + "List")));
                var revision = s.Revision; Invoke(t, "set_list", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p2" }, list = kind }); Equal(revision, s.Revision);
                NumberList(t, s, "p2"); Equal(null, (string)s.Blocks[2].Draft.Descendants(One + "Number").Single().Attribute("restartNumberingAt"));
                var api = new FakePage(s.Page); api.AfterSave = () => RenderNumbering(api.Page);
                var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None); Equal("Verified", r.Status);
                Equal("7.,1.,2.,3.", string.Join(",", api.Page.Descendants(One + "Number").Select(n => (string)n.Attribute("text"))));
                Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status); Equal(Texts(s.Page), Texts(api.Page));
            }
        });
        Test("markdown numbering rejects a bad batch before changing already restored neighbors", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "1. 第一项"), Paragraph("b", "1. 第二项"), Paragraph("c", "1. 第三项")));
            var t = Tools(s); Read(t, s); Cleanup(t, s, "p1", "p2", "p3"); NumberList(t, s, "p3");
            var draft = s.CreateDraftPage().ToString(); var revision = s.Revision;
            Throws(() => NumberList(t, s, "p2", "missing")); Equal(draft, s.CreateDraftPage().ToString()); Equal(revision, s.Revision);
            NumberList(t, s, "p2"); Equal(null, (string)s.Blocks[2].Draft.Descendants(One + "Number").Single().Attribute("restartNumberingAt"));
        });
        Test("markdown preserves non-one starts and consecutive one markers across separate groups", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "5)&nbsp;第五项"), Paragraph("b", "6. 第六项"), Paragraph("gap", "正文"),
                Paragraph("c", "1. 第一项"), Paragraph("d", "1. 第二项"), Paragraph("gap2", "另一组"), Paragraph("e", "0. 第零项")));
            var t = Tools(s); Read(t, s); var ids = new[] { "p1", "p2", "p4", "p5", "p7" }; Cleanup(t, s, ids);
            foreach (var id in ids.Reverse()) NumberList(t, s, id);
            var api = new FakePage(s.Page); api.AfterSave = () => RenderNumbering(api.Page);
            Equal("Verified", new AgentCommitter(api).Commit(s, CancellationToken.None).Status);
            Equal("5.,6.,1.,2.,0.", string.Join(",", api.Page.Descendants(One + "Number").Select(n => (string)n.Attribute("text"))));
        });
        Test("markdown number groups do not cross outlines, cells or paragraph parents", () =>
        {
            var grid = GridPage();
            var cells = grid.Descendants(One + "Cell").Take(2).ToArray();
            cells[0].Element(One + "OEChildren").ReplaceNodes(Paragraph("c1", "1. 单元格一"), Paragraph("c2", "2. 单元格二"));
            cells[1].Element(One + "OEChildren").ReplaceNodes(Paragraph("other1", "1. 另一单元格"), Paragraph("other2", "2. 第二项"));
            var parent = Paragraph("a1", "1. 父项"); parent.Add(new XElement(One + "OEChildren", Paragraph("child1", "1. 子项"), Paragraph("child2", "2. 子项二")));
            var page = Boxes(Box("A", 100, parent, Paragraph("a2", "2. 父项二")), Box("B", 300, Paragraph("b1", "1. 第二框"), Paragraph("b2", "2. 第二项"), grid.Descendants(One + "Table").Single().Parent));
            var s = Snapshot(page); var t = Tools(s); Read(t, s);
            var ids = s.Blocks.Where(b => b.Editable && char.IsDigit(b.Text[0])).Select(b => b.Id).ToArray(); Cleanup(t, s, ids);
            foreach (var id in ids.Reverse()) NumberList(t, s, id);
            var api = new FakePage(s.Page); api.AfterSave = () => RenderNumbering(api.Page);
            Equal("Verified", new AgentCommitter(api).Commit(s, CancellationToken.None).Status);
            Equal("1.,1.,2.,2.,1.,2.,1.,2.,1.,2.", string.Join(",", api.Page.Descendants(One + "Number").Select(n => (string)n.Attribute("text"))));
            Equal(5, api.Page.Descendants(One + "Number").Count(n => n.Attribute("restartNumberingAt") != null));
        });
        Test("selected markdown numbers retain their first value without changing unselected content", () =>
        {
            var page = Page(Paragraph("a", "3. 未选中"), Paragraph("b", "4. 第四项"), Paragraph("c", "5. 第五项"), Paragraph("d", "6. 未选中"));
            var s = new AgentPageSnapshot(page.ToString(), new HashSet<string> { "b", "c" }, new AgentOptions()); var t = Tools(s); Read(t, s);
            Cleanup(t, s, "p1", "p2"); NumberList(t, s, "p2"); NumberList(t, s, "p1");
            var api = new FakePage(s.Page); api.AfterSave = () => RenderNumbering(api.Page); var c = new AgentCommitter(api);
            var r = c.Commit(s, CancellationToken.None); Equal("Verified", r.Status);
            Equal("4.,5.", string.Join(",", api.Page.Descendants(One + "Number").Select(n => (string)n.Attribute("text"))));
            Equal(1, api.Page.Descendants(One + "Number").Count(n => n.Attribute("restartNumberingAt") != null));
            True(XNode.DeepEquals(AgentCommitter.Find(page, "a"), AgentCommitter.Find(api.Page, "a")));
            True(XNode.DeepEquals(AgentCommitter.Find(page, "d"), AgentCommitter.Find(api.Page, "d")));
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status); Equal(Texts(page), Texts(api.Page));
        });
        Test("markdown number restoration keeps existing native numbering and ordinary list behavior", () =>
        {
            var existing = Paragraph("a", "5. 保留原生样式");
            existing.AddFirst(new XElement(One + "List", new XElement(One + "Number", new XAttribute("numberSequence", "1"), new XAttribute("numberFormat", "(##)"), new XAttribute("restartNumberingAt", "7"))));
            var s = Snapshot(Page(existing, Paragraph("b", "普通正文"), Paragraph("c", "1. **只清强调**"))); var t = Tools(s); Read(t, s);
            Cleanup(t, s, "p1");
            Invoke(t, "strip_markdown", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p3" }, kinds = new[] { "emphasis" } });
            NumberList(t, s, "p1", "p2", "p3");
            True(XNode.DeepEquals(s.Blocks[0].Original.Element(One + "List"), s.Blocks[0].Draft.Element(One + "List")));
            foreach (var b in s.Blocks.Skip(1)) Equal(null, (string)b.Draft.Descendants(One + "Number").Single().Attribute("restartNumberingAt"));
        });
        Test("numbering controls are verified, while generated text and fonts remain transient", () =>
        {
            foreach (var attribute in new[] { "restartNumberingAt", "numberSequence", "numberFormat" })
            {
                var s = Snapshot(Page(Paragraph("a", "5. 第五项"))); var t = Tools(s); Read(t, s); Cleanup(t, s, "p1"); NumberList(t, s, "p1");
                var api = new FakePage(s.Page);
                api.AfterSave = () => api.Page.Descendants(One + "Number").Single().SetAttributeValue(attribute, attribute == "numberFormat" ? "##)" : attribute == "numberSequence" ? "1" : null);
                var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
                Equal(1, r.Unverified); Equal(0, r.Applied); Equal(0, r.MarkdownMarks); True(!r.CanUndo);
            }
            var normalized = Paragraph("n", "正文"); AgentMarks.SetList(normalized, "number"); normalized.Descendants(One + "Number").Single().SetAttributeValue("restartNumberingAt", "5");
            var readback = new XElement(normalized); var number = readback.Descendants(One + "Number").Single();
            number.SetAttributeValue("numberSequence", "00"); number.SetAttributeValue("restartNumberingAt", "05"); number.SetAttributeValue("fontSize", "11.0"); number.SetAttributeValue("text", "5.");
            Equal(AgentMarks.Projection(normalized, Page(normalized)), AgentMarks.Projection(readback, Page(readback)));
            number.SetAttributeValue("restartNumberingAt", "6"); True(AgentMarks.DraftKey(normalized) != AgentMarks.DraftKey(readback));
        });
        Test("markdown number restoration skips concurrent edits and undo protects later numbering changes", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "5. 第五项"))); var t = Tools(s); Read(t, s); Cleanup(t, s, "p1"); NumberList(t, s, "p1");
            var api = new FakePage(s.Page); AgentCommitter.Find(api.Page, "a").Element(One + "T").Value += "用户补充";
            var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None); Equal(1, r.Conflicts); Equal(0, api.Writes);
            api = new FakePage(s.Page); c = new AgentCommitter(api); r = c.Commit(s, CancellationToken.None); Equal("Verified", r.Status);
            api.Page.Descendants(One + "Number").Single().SetAttributeValue("restartNumberingAt", "8");
            var writes = api.Writes; var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None); Equal(1, undo.Conflicts); Equal(writes, api.Writes);
        });
        Test("a conflicting markdown list start also skips dependent items but other lists still commit", () =>
        {
            var s = Snapshot(Page(Paragraph("a1", "1. 第一组"), Paragraph("a2", "2. 第二项"), Paragraph("gap", "正文"),
                Paragraph("b1", "1. 第二组"), Paragraph("b2", "1. 后续项"), Paragraph("b3", "1. 第三项")));
            var t = Tools(s); Read(t, s); var ids = new[] { "p1", "p2", "p4", "p5", "p6" }; Cleanup(t, s, ids);
            foreach (var id in ids.Reverse()) NumberList(t, s, id);
            var api = new FakePage(s.Page); AgentCommitter.Find(api.Page, "b1").Element(One + "T").Value += "用户编辑";
            var protectedText = Texts(api.Page).Split('|').Skip(3).ToArray();
            api.AfterSave = () => RenderNumbering(api.Page); var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
            Equal("PartiallyApplied", r.Status); Equal(2, r.Applied); Equal(3, r.Conflicts); Equal(2, r.MarkdownMarks);
            Equal("1.,2.", string.Join(",", api.Page.Descendants(One + "Number").Select(n => (string)n.Attribute("text"))));
            Equal(string.Join("|", protectedText), string.Join("|", Texts(api.Page).Split('|').Skip(3)));
        });
        Test("set_list numbers paragraphs with a style, verifies despite OneNote attributes and undo removes it", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "第一项"), Paragraph("b", "第二项"), Paragraph("c", "结尾"))); var t = Tools(s); Read(t, s);
            Invoke(t, "set_list", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1", "p2" }, list = "number" });
            Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" }, preset_id = "body" });
            Equal(2, s.Revision);
            Invoke(t, "set_list", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1", "p2" }, list = "number" }); Equal(2, s.Revision);
            True(Json(Invoke(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p2" } })).Contains("\"list\":\"number\""));
            var api = new FakePage(s.Page);
            api.AfterSave = () => { foreach (var n in api.Page.Descendants(One + "Number")) { n.SetAttributeValue("fontSize", "11.0"); n.SetAttributeValue("text", "1."); } };
            var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(2, r.Applied);
            Equal("number", AgentMarks.ListKind(AgentCommitter.Find(api.Page, "b")));
            Equal("List", AgentCommitter.Find(api.Page, "a").Elements().First().Name.LocalName);
            api.AfterSave = null;
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal("Verified", undo.Status); Equal(2, undo.Applied); True(!api.Page.Descendants(One + "List").Any());
            // 去掉列表要重建段落：文字和顺序不变，ID 换成 OneNote 新分配的。
            True(AgentCommitter.Find(api.Page, "a") == null && AgentCommitter.Find(api.Page, "b") == null);
            Equal("第一项|第二项|结尾", string.Join("|", api.Page.Descendants(One + "OE").Select(AgentCode.PlainText)));
        });
        Test("set_list none rebuilds only the listed paragraph; children keep IDs; undo restores the bullet in place", () =>
        {
            var parent = Listed("a", "父项", "2"); parent.Add(new XElement(One + "OEChildren", Paragraph("b", "子项")));
            var s = Snapshot(Page(parent, Paragraph("c", "结尾"))); var t = Tools(s); Read(t, s);
            Invoke(t, "set_list", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" }, list = "none" });
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(1, r.Applied); True(AgentCommitter.Find(api.Page, "a") == null);
            var rebuilt = api.Page.Descendants(One + "OE").First();
            Equal("父项", AgentCode.PlainText(rebuilt)); Equal("none", AgentMarks.ListKind(rebuilt));
            Equal("b", (string)rebuilt.Element(One + "OEChildren").Element(One + "OE").Attribute("objectID"));
            Equal((string)rebuilt.Attribute("objectID"), r.Undo.Single().ObjectId);
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal("Verified", undo.Status); Equal("2", (string)AgentCommitter.Find(api.Page, (string)rebuilt.Attribute("objectID")).Descendants(One + "Bullet").Single().Attribute("bullet"));
        });
        Test("set_list keeps existing bullets and rejects title, unread and code paragraphs", () =>
        {
            var code = Paragraph("code", "int x = 1;"); code.SetAttributeValue("style", "font-family:Consolas");
            var p = Page(Listed("l", "已有列表", "13"), Paragraph("a", "正文"), code); p.AddFirst(new XElement(One + "Title", Paragraph("title", "标题")));
            var s = Snapshot(p); var t = Tools(s);
            Rejects("请先完整读取", () => Invoke(t, "set_list", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p3" }, list = "bullet" }));
            Read(t, s);
            Rejects("页面标题", () => Invoke(t, "set_list", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" }, list = "bullet" }));
            Rejects("代码段落", () => Invoke(t, "set_list", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p3", "p4" }, list = "bullet" }));
            Invoke(t, "set_list", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p2" }, list = "bullet" });
            Equal(0, s.Revision); Equal("13", (string)s.Blocks[1].Draft.Descendants(One + "Bullet").Single().Attribute("bullet"));
        });
        Test("set_tag reuses the page to-do definition, adds one for important, verifies and undo removes tags", () =>
        {
            var p = Page(Paragraph("a", "买菜"), Paragraph("b", "交报告"), Paragraph("c", "关键结论")); p.AddFirst(TagDef("0", 3, "待办事项"));
            var s = Snapshot(p); var t = Tools(s); Read(t, s);
            Invoke(t, "set_tag", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1", "p2" }, tag = "todo" });
            Invoke(t, "set_tag", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p2" }, tag = "todo", completed = true });
            Invoke(t, "set_tag", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p3" }, tag = "important" });
            Invoke(t, "set_tag", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p3" }, tag = "important" });
            Equal(3, s.Revision); Equal(2, s.DraftTags.Elements().Count()); Equal(1, s.Blocks[1].Draft.Elements(One + "Tag").Count());
            Rejects("completed 只能", () => Invoke(t, "set_tag", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p3" }, tag = "important", completed = true }));
            True(Json(Invoke(t, "get_page_overview", new { })).Contains("\"tags\":[\"todo:done\"]"));
            var api = new FakePage(s.Page);
            api.AfterSave = () => { foreach (var tag in api.Page.Descendants(One + "Tag")) tag.SetAttributeValue("creationDate", "2026-09-28T00:00:00.000Z"); };
            var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(3, r.Applied); Equal(2, api.Page.Elements(One + "TagDef").Count());
            var important = api.Page.Elements(One + "TagDef").Single(d => (string)d.Attribute("symbol") == "13");
            Equal((string)important.Attribute("index"), (string)AgentCommitter.Find(api.Page, "c").Element(One + "Tag").Attribute("index"));
            Equal("0", (string)AgentCommitter.Find(api.Page, "a").Element(One + "Tag").Attribute("index"));
            Equal("true", (string)AgentCommitter.Find(api.Page, "b").Element(One + "Tag").Attribute("completed"));
            api.AfterSave = null;
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status); True(!api.Page.Descendants(One + "Tag").Any());
        });
        Test("set_tag none removes only the three supported tags", () =>
        {
            var oe = Paragraph("a", "事项"); oe.AddFirst(Tag("0"), Tag("1"));
            var p = Page(oe); p.AddFirst(TagDef("0", 3, "待办事项"), TagDef("1", 99, "自定义"));
            var s = Snapshot(p); var t = Tools(s); Read(t, s);
            Invoke(t, "set_tag", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" }, tag = "none" });
            Equal("1", (string)s.Blocks[0].Draft.Elements(One + "Tag").Single().Attribute("index"));
            Equal("other", AgentMarks.Describe(s.Blocks[0].Draft, s.DraftTags).Single());
            Equal("Verified", new AgentCommitter(new FakePage(s.Page)).Commit(s, CancellationToken.None).Status);
        });
        Test("concurrent tag added to a staged paragraph skipped", () => Conflict(p => AgentCommitter.Find(p, "a").AddFirst(Tag("0"))));
        Test("OneNote changing the bullet of an unspecified paragraph is not reported verified", () =>
        {
            var s = Prepared(Page(Paragraph("a", "第一段"), Listed("b", "列表项", "2"))); var api = new FakePage(s.Page);
            api.AfterSave = () => AgentCommitter.Find(api.Page, "b").Descendants(One + "Bullet").Single().SetAttributeValue("bullet", "5");
            Equal("CommitOutcomeUnknown", new AgentCommitter(api).Commit(s, CancellationToken.None).Status);
        });
        Test("TagDef renumbering keeps untouched tags verified", () =>
        {
            var tagged = Paragraph("b", "待办"); tagged.AddFirst(Tag("0"));
            var p = Page(Paragraph("a", "第一段"), tagged); p.AddFirst(TagDef("0", 3, "待办事项"));
            var s = Prepared(p); var api = new FakePage(s.Page);
            api.AfterSave = () => { api.Page.Element(One + "TagDef").SetAttributeValue("index", "7"); AgentCommitter.Find(api.Page, "b").Element(One + "Tag").SetAttributeValue("index", "7"); };
            Equal("Verified", new AgentCommitter(api).Commit(s, CancellationToken.None).Status);
        });
        Test("set_table_style changes borders, header row and shading, verifies and undo restores", () =>
        {
            var s = Snapshot(GridPage()); var t = Tools(s);
            Equal(1, s.Tables.Count); Equal("t1", s.Blocks[0].TableId); Equal("名称 | 说明", s.Tables[0].Summary);
            var style = new { borders = false, header_row = true, header_shading = "#DEEAF6" };
            Invoke(t, "set_table_style", new { snapshot_id = s.SnapshotId, table_ids = new[] { "t1" }, style }); Equal(1, s.Revision);
            Invoke(t, "set_table_style", new { snapshot_id = s.SnapshotId, table_ids = new[] { "t1" }, style }); Equal(1, s.Revision);
            True(Json(Invoke(t, "get_pending_changes", new { snapshot_id = s.SnapshotId })).Contains("\"tables_changed\":[\"t1\"]"));
            var api = new FakePage(s.Page);
            api.AfterSave = () =>
            {
                // OneNote 省略值为 false 的开关，颜色写成小写。
                var written = api.Page.Descendants(One + "Table").Single(); written.Attribute("bordersVisible")?.Remove();
                foreach (var cell in written.Element(One + "Row").Elements(One + "Cell")) cell.SetAttributeValue("shadingColor", "#deeaf6");
            };
            var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(1, r.Tables); Equal(1, r.TableUndo.Count); Equal(1, api.Writes);
            var table = api.Page.Descendants(One + "Table").Single();
            True(!TableLook.Flag(table, "bordersVisible")); True(TableLook.Flag(table, "hasHeaderRow"));
            Equal("", TableLook.Shade((string)table.Elements(One + "Row").Last().Element(One + "Cell").Attribute("shadingColor")));
            api.AfterSave = null;
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal("Verified", undo.Status); Equal(1, undo.Tables);
            table = api.Page.Descendants(One + "Table").Single();
            True(TableLook.Flag(table, "bordersVisible")); True(!TableLook.Flag(table, "hasHeaderRow"));
            Equal("none", TableLook.Read(table).ShadingName);
        });
        Test("table style: code boxes protected, appearance edits conflict, cell text edits do not", () =>
        {
            var box = new XElement(One + "OE", new XAttribute("objectID", "w"), CodeBlockBuilder.BuildTable("a = 1", LanguageRegistry.Find("python"), CodeThemes.Light, new AddInSettings()));
            var n = 0; foreach (var e in box.Descendants().Where(e => e.Name.LocalName == "Table" || e.Name.LocalName == "Row" || e.Name.LocalName == "Cell" || e.Name.LocalName == "OE")) e.SetAttributeValue("objectID", "box" + n++);
            var coded = Snapshot(Page(Paragraph("a", "正文"), box));
            Equal("highlighted_code", coded.Tables.Single().ProtectedReason); True(!Json(Tools(coded).Definitions).Contains("set_table_style"));
            var s = Snapshot(GridPage()); var t = Tools(s);
            Invoke(t, "set_table_style", new { snapshot_id = s.SnapshotId, table_ids = new[] { "t1" }, style = new { header_shading = "#F2F2F2" } });
            var api = new FakePage(s.Page); api.Page.Descendants(One + "Table").Single().SetAttributeValue("bordersVisible", "false");
            var r = new AgentCommitter(api).Commit(s, CancellationToken.None); Equal("NoChange", r.Status); Equal(1, r.Conflicts); Equal(0, api.Writes);
            api = new FakePage(s.Page); AgentCommitter.Find(api.Page, "d1").Element(One + "T").Value = "改过";
            r = new AgentCommitter(api).Commit(s, CancellationToken.None); Equal("Verified", r.Status); Equal(1, r.Tables);
            Equal("改过", new AgentRichText(AgentCommitter.Find(api.Page, "d1")).Text);
        });
        Test("read_image_text returns recognized text only when OneNote has OCR data", () =>
        {
            var image = new XElement(One + "OE", new XAttribute("objectID", "img"), new XElement(One + "Image", new XAttribute("objectID", "image1"),
                new XElement(One + "OCRData", new XAttribute("lang", "en-US"), new XElement(One + "OCRText", new XCData("OCR sample " + new string('x', 4100))))));
            var s = Snapshot(Page(Paragraph("a", "正文"), image)); var t = Tools(s);
            Equal(1, s.Images.Count); Equal(AgentImage.MaxChars, s.Images[0].Text.Length);
            var json = Json(Invoke(t, "read_image_text", new { snapshot_id = s.SnapshotId, image_ids = new[] { "i1" } }));
            True(json.Contains("OCR sample")); True(json.Contains("\"truncated\":true"));
            Rejects("没有识别出的文字", () => Invoke(t, "read_image_text", new { snapshot_id = s.SnapshotId, image_ids = new[] { "i2" } }));
            True(Json(Invoke(t, "get_page_overview", new { })).Contains("\"images\":[{\"id\":\"i1\""));
            True(AgentRunner.SystemPrompt(t).Contains("read_image_text"));
            True(!Json(Tools(Snapshot()).Definitions).Contains("read_image_text"));
        });
        Test("switches remove list, tag and table tools and their prompts", () =>
        {
            var c = AiConfigStore.Parse(XElement.Parse("<AiConfig><Agent><EnableTags>false</EnableTags></Agent></AiConfig>"));
            True(!c.Agent.EnableTags); True(c.Agent.EnableLists); True(c.Agent.EnableTableStyles);
            var off = Tools(new AgentPageSnapshot(GridPage().ToString(), null, new AgentOptions { EnableLists = false, EnableTags = false, EnableTableStyles = false }));
            var defs = Json(off.Definitions); var prompt = AgentRunner.SystemPrompt(off);
            True(!defs.Contains("set_list") && !defs.Contains("set_tag") && !defs.Contains("set_table_style"));
            True(!prompt.Contains("set_list") && !prompt.Contains("set_tag") && !prompt.Contains("set_table_style"));
            var on = Tools(Snapshot(GridPage()));
            True(Json(on.Definitions).Contains("set_table_style")); True(AgentRunner.SystemPrompt(on).Contains("set_list"));
        });
        Test("step descriptions for list, tag, table and image tools", () =>
        {
            Equal(("设置列表 · 编号 · 2 段", AgentStepState.Done), AgentTools.DescribeStep("set_list", "{\"block_ids\":[\"a\",\"b\"],\"list\":\"number\"}", "{\"ok\":true}"));
            Equal(("设置标记 · 待办已完成 · 1 段", AgentStepState.Done), AgentTools.DescribeStep("set_tag", "{\"block_ids\":[\"a\"],\"tag\":\"todo\",\"completed\":true}", "{\"ok\":true}"));
            Equal(("设置表格样式 · 1 个表格", AgentStepState.Done), AgentTools.DescribeStep("set_table_style", "{\"table_ids\":[\"t1\"]}", "{\"ok\":true}"));
            Equal(("读取图片文字 · 2 张", AgentStepState.Done), AgentTools.DescribeStep("read_image_text", "{\"image_ids\":[\"i1\",\"i2\"]}", "{}"));
        });
        Test("code spacing normalizes both sides for zero, one and many ordinary or monospace blank paragraphs", () =>
        {
            foreach (var count in new[] { 0, 1, 3 }) foreach (var mono in new[] { false, true })
            {
                var blanks = Enumerable.Range(0, count).Select(i => Paragraph("e" + i, i % 2 == 0 ? "&nbsp; " : " ")).ToList();
                if (mono) foreach (var blank in blanks) blank.SetAttributeValue("style", "font-family:Consolas");
                var s = Snapshot(Page(new[] { Paragraph("a", "正文") }.Concat(blanks).Concat(new[] { SpacingBox("code"), Paragraph("b", "结尾") }).ToArray()));
                var t = Tools(s); Read(t, s); var result = Normalize(t, s);
                Equal(count == 3 ? 2 : 0, AiClient.Get(result, "removed_paragraphs"));
                Equal(count == 0 ? 2 : 1, AiClient.Get(result, "inserted_paragraphs"));
                Equal(count == 1 ? "e0" : count == 3 ? "e0" : null, (string)s.Layout.Element(One + "Outline").Element(One + "OEChildren").Elements(One + "OE").ElementAt(1).Attribute("objectID"));
                var revision = s.Revision; True((bool)AiClient.Get(Normalize(t, s), "noop")); Equal(revision, s.Revision);
                var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
                Equal("Verified", r.Status); Equal(1, r.OutlineUndo.Count);
                Equal("正文||#||结尾", SpacingTexts(api.Page));
                Equal("x = 1\n\ny = 2", string.Join("\n", api.Page.Descendants(One + "Table").Single().Descendants(One + "OE").Select(AgentCode.PlainText)));
                True(!api.LastXml.Contains("trim-leading") && !api.LastXml.Contains("trim-trailing") && !api.LastXml.Contains("urn:onenote-code-helper:agent"));
                Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status);
                Equal(SpacingTexts(s.Page), SpacingTexts(api.Page));
            }
        });
        Test("code spacing preserves a single soft blank and trims excess soft blanks on either edge", () =>
        {
            foreach (var count in new[] { 1, 3 })
            {
                var html = string.Concat(Enumerable.Repeat("<br>&nbsp;", count));
                var s = Snapshot(Page(Paragraph("a", "<b>正文</b>" + html), SpacingBox("code"), Paragraph("b", string.Concat(Enumerable.Repeat("&nbsp;<br>", count)) + "<a href='https://example.com'>链接</a>")));
                var t = Tools(s); Read(t, s); var result = Normalize(t, s);
                Equal((count - 1) * 2, AiClient.Get(result, "removed_soft_lines")); Equal(0, AiClient.Get(result, "inserted_paragraphs"));
                if (count == 1) { Equal(0, s.Revision); continue; }
                Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1", s.Blocks.Last().Id }, preset_id = "body" });
                var read = Json(Invoke(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" } })); True(!read.Contains("\\n\\u00a0\\n"));
                var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
                Equal("Verified", r.Status); Equal(4, r.RemovedSoftLines); Equal(2, r.Applied); Equal(1, r.OutlineUndo.Count);
                Equal("正文\n\u00a0", new AgentRichText(AgentCommitter.Find(api.Page, "a")).Text);
                Equal("\u00a0\n链接", new AgentRichText(AgentCommitter.Find(api.Page, "b")).Text);
                True(AgentCommitter.Find(api.Page, "a").ToString().Contains("<b>")); True(AgentCommitter.Find(api.Page, "b").ToString().Contains("https://example.com"));
                Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status);
                Equal(new AgentRichText(s.Blocks.First().Original).Text, new AgentRichText(AgentCommitter.Find(api.Page, "a")).Text);
            }
        });
        Test("code spacing counts mixed paragraph and soft blanks and keeps the first existing blank", () =>
        {
            var blank = Paragraph("e1", "&nbsp;<br> <br>&nbsp;"); blank.SetAttributeValue("style", "font-family:Consolas");
            var s = Snapshot(Page(Paragraph("a", "<b>正文</b><br><br>"), blank, Paragraph("e2", ""), SpacingBox("code"), Paragraph("b", "<br><br>结尾")));
            var t = Tools(s); Read(t, s); var result = Normalize(t, s);
            Equal(1, AiClient.Get(result, "removed_paragraphs")); Equal(5, AiClient.Get(result, "removed_soft_lines"));
            Equal("正文||#|\n结尾", SpacingTexts(s.CreateDraftPage()));
            var api = new FakePage(s.Page); var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(1, r.Removed); Equal(5, r.RemovedSoftLines);
            True(AgentCommitter.Find(api.Page, "e1") != null); Equal("\u00a0", new AgentRichText(AgentCommitter.Find(api.Page, "e1")).Text);
        });
        Test("code spacing handles scheduled conversions and leaves code interior and unrelated gaps unchanged", () =>
        {
            var s = CodePrepared(); var t = Tools(s); var result = Normalize(t, s);
            Equal(2, AiClient.Get(result, "inserted_paragraphs"));
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(1, r.CodeBlocks); Equal(2, r.InsertedBlankLines); Equal(1, r.OutlineUndo.Count);
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status); Equal(Texts(s.Page), Texts(api.Page));
            var p = Page(Paragraph("edge", ""), Paragraph("a", "甲<br><br><br>乙"), Paragraph("e1", ""), Paragraph("e2", ""), Paragraph("b", "正文"),
                SpacingBox("c1"), Paragraph("between1", ""), Paragraph("between2", ""), SpacingBox("c2"), Paragraph("tail", ""));
            var other = Snapshot(p); var ot = Tools(other); Read(ot, other); Normalize(ot, other);
            foreach (var id in new[] { "edge", "e1", "e2", "between1", "between2", "tail" }) True(AgentCommitter.Find(other.Layout, id) != null);
            Equal("甲\n\n\n乙", new AgentRichText(AgentCommitter.Find(other.CreateDraftPage(), "a")).Text);
        });
        Test("code spacing handles nested flows, ordinary cells and identifiable blank OEs without T", () =>
        {
            var parent = Paragraph("parent", "父段落"); parent.Add(new XElement(One + "OEChildren", Paragraph("child", "子段落<br><br>"), SpacingBox("nested")));
            var s = Snapshot(Page(parent, new XElement(One + "OE", new XAttribute("objectID", "noT1")), new XElement(One + "OE", new XAttribute("objectID", "noT2")), SpacingBox("outer"), Paragraph("b", "结尾")));
            var t = Tools(s); Read(t, s); Normalize(t, s);
            Equal("子段落\n", new AgentRichText(AgentCommitter.Find(s.CreateDraftPage(), "child")).Text);
            True(AgentCommitter.Find(s.Layout, "child").Parent.Parent == AgentCommitter.Find(s.Layout, "parent"));
            True(AgentCommitter.Find(s.Layout, "noT1") != null && AgentCommitter.Find(s.Layout, "noT2") != null); // 两个代码块之间不清理。
            var noT = Snapshot(Page(Paragraph("a", "正文"), new XElement(One + "OE", new XAttribute("objectID", "e1")), new XElement(One + "OE", new XAttribute("objectID", "e2")), SpacingBox("c")));
            var nt = Tools(noT); Read(nt, noT); Equal(1, AiClient.Get(Normalize(nt, noT), "removed_paragraphs"));
            Equal("Verified", new AgentCommitter(new FakePage(noT.Page)).Commit(noT, CancellationToken.None).Status);
            var table = new XElement(One + "OE", new XAttribute("objectID", "wrapper"), new XElement(One + "Table", new XAttribute("objectID", "grid"),
                new XElement(One + "Columns", new XElement(One + "Column", new XAttribute("index", 0), new XAttribute("width", 300))),
                new XElement(One + "Row", new XElement(One + "Cell", new XElement(One + "OEChildren", Paragraph("ca", "文字"), SpacingBox("cellCode"), Paragraph("cb", "结尾"))))));
            var cell = Snapshot(Page(table)); var ct = Tools(cell); Read(ct, cell); Equal(2, AiClient.Get(Normalize(ct, cell), "inserted_paragraphs"));
            Equal("Verified", new AgentCommitter(new FakePage(cell.Page)).Commit(cell, CancellationToken.None).Status);
        });
        Test("code spacing respects selection boundaries, protected blanks and disabled capabilities", () =>
        {
            var page = Page(Paragraph("a", "正文"), Paragraph("e1", ""), Paragraph("e2", ""), SpacingBox("code"), Paragraph("b", "结尾"));
            var selected = new AgentPageSnapshot(page.ToString(), new HashSet<string> { "a", "e1", "code1", "code3" }, new AgentOptions());
            var t = Tools(selected); Read(t, selected); var result = Normalize(t, selected);
            Equal(0, selected.Revision); True(Json(result).Contains("outside_selection"));
            var full = new AgentPageSnapshot(page.ToString(), new HashSet<string> { "a", "e1", "e2", "code1", "code3", "b" }, new AgentOptions());
            var ft = Tools(full); Read(ft, full); Equal(1, AiClient.Get(Normalize(ft, full), "removed_paragraphs"));
            Equal("Verified", new AgentCommitter(new FakePage(full.Page)).Commit(full, CancellationToken.None).Status);
            var marked = Paragraph("marked", ""); marked.Add(new XElement(One + "List", new XElement(One + "Bullet", new XAttribute("bullet", "2"))));
            var protectedPage = Snapshot(Page(Paragraph("a", "正文"), marked, Paragraph("empty", ""), SpacingBox("code")));
            var pt = Tools(protectedPage); Read(pt, protectedPage); True(Json(Normalize(pt, protectedPage)).Contains("protected_blank")); Equal(0, protectedPage.Revision);
            var singleMarked = Snapshot(Page(Paragraph("a", "正文"), new XElement(marked), SpacingBox("code")));
            var mt = Tools(singleMarked); Read(mt, singleMarked); True(Json(Normalize(mt, singleMarked)).Contains("protected_blank")); Equal(0, singleMarked.Revision);
            foreach (var removing in new[] { true, false })
            {
                var p = removing ? Page(Paragraph("a", "正文<br><br>"), SpacingBox("c")) : Page(Paragraph("a", "正文"), SpacingBox("c"));
                var off = new AgentPageSnapshot(p.ToString(), null, new AgentOptions { EnableBlankLineRemoval = !removing, EnableInsert = removing });
                var ot = Tools(off); Read(ot, off); True(Json(Normalize(ot, off)).Contains(removing ? "removal_disabled" : "insert_disabled")); Equal(0, off.Revision);
            }
            var disabled = Tools(new AgentPageSnapshot(page.ToString(), null, new AgentOptions { EnableBlankLineRemoval = false, EnableInsert = false }));
            True(!disabled.Has("normalize_code_spacing")); True(!AgentRunner.SystemPrompt(disabled).Contains("normalize_code_spacing"));
        });
        Test("code spacing soft edits skip the entire conflicted frame while independent style changes can commit", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "正文<br><br><br>"), Paragraph("e1", ""), Paragraph("e2", ""), SpacingBox("code"), Paragraph("b", "结尾")));
            var t = Tools(s); Read(t, s); Normalize(t, s);
            Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" }, preset_id = "body" });
            var api = new FakePage(s.Page); AgentCommitter.Find(api.Page, "b").Element(One + "T").Value = "用户后来编辑";
            var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
            True(r.Conflicts > 0); Equal(0, r.Removed); Equal(0, r.RemovedSoftLines); Equal(0, r.InsertedBlankLines); Equal(1, r.Applied);
            Equal("正文\n\n\n", new AgentRichText(AgentCommitter.Find(api.Page, "a")).Text); True(AgentCommitter.Find(api.Page, "e2") != null);
            Equal("用户后来编辑", AgentCode.PlainText(AgentCommitter.Find(api.Page, "b")));
        });
        Test("finish rejects code spacing changed after normalization and permits an idempotent refresh", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "正文"), Paragraph("e", ""), SpacingBox("code"), Paragraph("b", "结尾")));
            var api = new FakePage(s.Page); var t = new AgentTools(s, new AgentCommitter(api), CancellationToken.None); Read(t, s); Normalize(t, s);
            Invoke(t, "remove_blank_lines", new { snapshot_id = s.SnapshotId, mode = "all" });
            True(Json(Invoke(t, "finish_edit", new { snapshot_id = s.SnapshotId, draft_revision = s.Revision })).Contains("normalize_code_spacing"));
            Equal(0, api.Writes); True(!s.Frozen);
            Normalize(t, s); var revision = s.Revision; Normalize(t, s); Equal(revision, s.Revision);
            Invoke(t, "finish_edit", new { snapshot_id = s.SnapshotId, draft_revision = s.Revision }); Equal("Verified", t.Report.Status);
        });
        Test("code spacing never reports an unverified write as successful and undo preserves later edits", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "正文<br><br>"), SpacingBox("code"))); var t = Tools(s); Read(t, s); Normalize(t, s);
            var unread = new FakePage(s.Page) { FailReadAfterSave = true }; var ur = new AgentCommitter(unread).Commit(s, CancellationToken.None);
            Equal("CommitOutcomeUnknown", ur.Status); Equal(0, ur.RemovedSoftLines); Equal(0, ur.OutlineUndo.Count);
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            AgentCommitter.Find(api.Page, "a").Element(One + "T").Value = "用户编辑";
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None); True(undo.Conflicts > 0); Equal("用户编辑", AgentCode.PlainText(AgentCommitter.Find(api.Page, "a")));
        });
        Test("code spacing preserves rich runs, emoji, text fixes and links through soft-line trimming", () =>
        {
            var p = Paragraph("a", "<br><br><b>按装😀</b>", "<a href='https://example.com'>链接</a><br><br>");
            var s = Snapshot(Page(SpacingBox("first"), p, SpacingBox("second"))); var t = Tools(s); Read(t, s);
            var id = s.Blocks.Single(b => b.ObjectId == "a").Id;
            Fix(t, s, id, "按装", "安装"); Normalize(t, s);
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(2, r.RemovedSoftLines); Equal(1, r.TextFixes.Count);
            Equal("\n安装😀链接\n", new AgentRichText(AgentCommitter.Find(api.Page, "a")).Text);
            Equal(2, AgentCommitter.Find(api.Page, "a").Elements(One + "T").Count());
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status);
            Equal("\n\n按装😀链接\n\n", new AgentRichText(AgentCommitter.Find(api.Page, "a")).Text);
        });
        Test("code spacing observes insertion quota, unknown blank identity, code fonts and ordinary empty tables", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "正文"), SpacingBox("code"))); var t = Tools(s); Read(t, s);
            for (var i = 0; i < AgentTools.MaxInserted; i++) Insert(t, s, "p1", "摘要");
            var before = new XElement(s.Layout); True(Json(Normalize(t, s)).Contains("insert_limit")); True(XNode.DeepEquals(before, s.Layout));
            var unknown = Snapshot(Page(Paragraph("a", "正文"), new XElement(One + "OE"), Paragraph("e", ""), SpacingBox("code")));
            var ut = Tools(unknown); Read(ut, unknown); True(Json(Normalize(ut, unknown)).Contains("protected_blank")); Equal(0, unknown.Revision);
            var missingId = Snapshot(Page(Paragraph("a", "正文"), new XElement(One + "OE", new XElement(One + "T", "")), Paragraph("e", ""), SpacingBox("code")));
            var it = Tools(missingId); Read(it, missingId); True(Json(Normalize(it, missingId)).Contains("protected_blank")); Equal(0, missingId.Revision);
            var separate = Snapshot(Boxes(Box("A", 100, Paragraph("a", "正文")), Box("B", 300, SpacingBox("code"))));
            var st = Tools(separate); Read(st, separate); Normalize(st, separate); Equal(0, separate.Revision);
            var unsupported = Snapshot(Page(Paragraph("a", "<b>不完整 HTML"), SpacingBox("code")));
            True(Json(Normalize(Tools(unsupported), unsupported)).Contains("protected_text")); Equal(0, unsupported.Revision);
            var ordinary = SpacingBox("ordinary"); ordinary.Descendants(One + "OE").Attributes("style").Remove();
            var empty = SpacingBox("empty"); foreach (var line in empty.Descendants(One + "T")) line.ReplaceNodes(new XCData("&nbsp;"));
            var other = Snapshot(Page(Paragraph("a", "正文"), ordinary, Paragraph("b", "正文"), empty)); var ot = Tools(other); Read(ot, other);
            Normalize(ot, other); Equal(0, other.Revision);
            var raw = Paragraph("raw", "print(1)"); raw.SetAttributeValue("style", "font-family:Consolas");
            var rawPage = new AgentPageSnapshot(Page(Paragraph("a", "正文"), raw, Paragraph("b", "结尾")).ToString(), null, new AgentOptions { EnableCodeHighlight = false });
            var rt = Tools(rawPage); Read(rt, rawPage); Normalize(rt, rawPage); Equal(0, rawPage.Revision);
        });
        Test("code spacing follows cross-frame groups and detects corrupted readback and cancellation", () =>
        {
            var s = Snapshot(Boxes(Box("A", 100, Paragraph("a", "正文<br><br>"), SpacingBox("code")), Box("B", 300, Paragraph("b", "第二框"))));
            var t = Tools(s); Read(t, s); Normalize(t, s);
            var id = s.Blocks.Single(b => b.ObjectId == "a").Id; var target = s.Blocks.Single(b => b.ObjectId == "b").Id;
            Move(t, s, new[] { id }, target, "after"); Normalize(t, s);
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(2, r.OutlineUndo.Count); Equal(r.OutlineUndo[0].Group, r.OutlineUndo[1].Group);
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status);
            var corrupted = Snapshot(Page(Paragraph("a", "正文<br><br>"), SpacingBox("code"))); var ct = Tools(corrupted); Read(ct, corrupted); Normalize(ct, corrupted);
            var broken = new FakePage(corrupted.Page) { AfterSave = () => { } };
            broken.AfterSave = () => AgentCommitter.Find(broken.Page, "a").Element(One + "T").Value = "正文<br><br>";
            Equal("CommitOutcomeUnknown", new AgentCommitter(broken).Commit(corrupted, CancellationToken.None).Status);
            var cancelled = new FakePage(corrupted.Page); var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            try { new AgentCommitter(cancelled).Commit(corrupted, cancellation.Token); } catch (OperationCanceledException) { }
            Equal(0, cancelled.Writes);
        });
        Test("code spacing tool reports meaningful steps and exposes capability even without current blanks", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "正文"), SpacingBox("code"))); var t = Tools(s);
            True(t.Has("normalize_code_spacing")); True(t.Has("remove_blank_lines"));
            var overview = AgentChatClient.Parse(Json(Invoke(t, "get_page_overview", new { })));
            Equal(true, AiClient.Get(overview, "code_spacing"));
            // 允许插入时提前提供清理工具，供同一任务清理之后补入的空行。
            Equal(true, AiClient.Get(overview, "blank_lines"));
            True(AgentRunner.SystemPrompt(t).Contains("只处理交界处"));
            Equal(("规范化代码框间隔 · 删除 2 个空段落、3 个段内空行，补入 1 行", AgentStepState.Done), AgentTools.DescribeStep("normalize_code_spacing", "{}",
                "{\"ok\":true,\"removed_paragraphs\":2,\"removed_soft_lines\":3,\"inserted_paragraphs\":1,\"skipped\":[]}"));
        });
        Test("code spacing verifies native empty-T normalization without ignoring blank font changes", () =>
        {
            foreach (var changeFont in new[] { false, true })
            {
                var s = Snapshot(Page(Paragraph("a", "正文"), SpacingBox("code"))); var t = Tools(s); Read(t, s); Normalize(t, s);
                var api = new FakePage(s.Page);
                api.AfterSave = () =>
                {
                    var blank = api.Page.Descendants(One + "OE").Single(e => ((string)e.Attribute("objectID") ?? "").StartsWith("new-", StringComparison.Ordinal));
                    blank.Element(One + "T").ReplaceNodes(new XCData(""));
                    if (changeFont) blank.SetAttributeValue("style", "font-family:Arial;font-size:40pt;color:#222222");
                };
                var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
                Equal(changeFont ? "PartiallyApplied" : "Verified", r.Status); Equal(changeFont ? 0 : 1, r.InsertedBlankLines);
                Equal(changeFont ? 1 : 0, r.Unverified);
            }
        });
        Test("code spacing verifies normalized existing soft blanks and restores them without overwriting later edits", () =>
        {
            foreach (var nativeBlank in new[] { "", " ", "&nbsp;" })
            foreach (var laterEdit in new[] { false, true })
            {
                var s = Snapshot(Page(Paragraph("a", "正文"), Paragraph("e", "&nbsp;<br>&nbsp;<br>&nbsp;"), SpacingBox("code")));
                var t = Tools(s); Read(t, s); Equal(2, AiClient.Get(Normalize(t, s), "removed_soft_lines"));
                var api = new FakePage(s.Page);
                api.AfterSave = () => AgentCommitter.Find(api.Page, "e").Element(One + "T").ReplaceNodes(new XCData(nativeBlank));
                var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
                Equal("Verified", r.Status); Equal(2, r.RemovedSoftLines); Equal(1, r.OutlineUndo.Count);
                var after = Snapshot(api.Page); foreach (var b in after.Blocks) b.Read = true;
                True(XNode.DeepEquals(after.Layout, AgentCodeSpacing.Build(after).Layout));
                api.AfterSave = null;
                if (laterEdit) AgentCommitter.Find(api.Page, "a").Element(One + "T").Value = "用户后来编辑";
                var writes = api.Writes; var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
                if (laterEdit)
                {
                    True(undo.Conflicts > 0); Equal(writes, api.Writes);
                    Equal("用户后来编辑", AgentCode.PlainText(AgentCommitter.Find(api.Page, "a")));
                }
                else
                {
                    Equal("Verified", undo.Status);
                    Equal("\u00a0\n\u00a0\n\u00a0", new AgentRichText(AgentCommitter.Find(api.Page, "e")).Text);
                }
            }
        });
        Test("code spacing rejects extra soft lines, text, links, missing T and unreadable HTML in blank readback", () =>
        {
            foreach (var existing in new[] { false, true })
            foreach (var html in new[] { "&nbsp;<br>", "&nbsp;<br><br>", "回读产生正文", "<a href='https://example.com' style='color:#222222'>&nbsp;</a>", "<b>未闭合", null })
            {
                var s = Snapshot(existing ? Page(Paragraph("a", "正文"), Paragraph("e", "&nbsp;<br>&nbsp;<br>&nbsp;"), SpacingBox("code")) :
                    Page(Paragraph("a", "正文"), SpacingBox("code")));
                var t = Tools(s); Read(t, s); Normalize(t, s);
                var api = new FakePage(s.Page);
                api.AfterSave = () =>
                {
                    var blank = existing ? AgentCommitter.Find(api.Page, "e") : api.Page.Descendants(One + "OE")
                        .Single(e => ((string)e.Attribute("objectID") ?? "").StartsWith("new-", StringComparison.Ordinal));
                    if (html == null) blank.Elements(One + "T").Remove();
                    else blank.Element(One + "T").ReplaceNodes(new XCData(html));
                };
                var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
                Equal(1, api.Writes); Equal("CommitOutcomeUnknown", r.Status);
                Equal(0, r.RemovedSoftLines); Equal(0, r.InsertedBlankLines); Equal(0, r.OutlineUndo.Count);
            }
        });
        Test("code spacing still verifies the appearance of normalized existing blanks", () =>
        {
            var changes = new Action<XElement>[]
            {
                e => e.SetAttributeValue("style", "font-family:Arial;font-size:11pt;color:#222222"),
                e => e.SetAttributeValue("style", "font-family:Calibri;font-size:40pt;color:#222222"),
                e => e.SetAttributeValue("style", "font-family:Calibri;font-size:11pt;color:#FF0000"),
                e => e.SetAttributeValue("alignment", "right"),
                e => e.SetAttributeValue("spaceAfter", "20")
            };
            foreach (var change in changes)
            {
                var blank = Paragraph("e", "&nbsp;<br>&nbsp;<br>&nbsp;"); blank.SetAttributeValue("style", "font-family:Calibri;font-size:11pt;color:#222222");
                var s = Snapshot(Page(Paragraph("a", "正文"), blank, SpacingBox("code"))); var t = Tools(s); Read(t, s); Normalize(t, s);
                var api = new FakePage(s.Page);
                api.AfterSave = () => { var e = AgentCommitter.Find(api.Page, "e"); e.Element(One + "T").ReplaceNodes(new XCData("")); change(e); };
                var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
                Equal("PartiallyApplied", r.Status); Equal(1, r.Unverified); Equal(0, r.RemovedSoftLines);
            }
        });
        Test("code spacing rejects changes to later inline styles of uniform blank readback", () =>
        {
            foreach (var existing in new[] { false, true })
            foreach (var style in new[] { "font-size:80pt", "font-family:Calibri", "color:#FF0000", "font-weight:bold", "font-style:italic", "text-decoration:underline" })
            {
                var s = Snapshot(existing ? Page(Paragraph("a", "正文"), Paragraph("e", "&nbsp;<br>&nbsp;<br>&nbsp;"), SpacingBox("code")) :
                    Page(Paragraph("a", "正文"), SpacingBox("code")));
                var t = Tools(s); Read(t, s); Normalize(t, s);
                var api = new FakePage(s.Page);
                api.AfterSave = () =>
                {
                    var blank = existing ? AgentCommitter.Find(api.Page, "e") : api.Page.Descendants(One + "OE")
                        .Single(e => ((string)e.Attribute("objectID") ?? "").StartsWith("new-", StringComparison.Ordinal));
                    blank.Element(One + "T").ReplaceNodes(new XCData("&nbsp;<span style='" + style + "'>&nbsp;</span>"));
                };
                var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
                if (r.Status != "PartiallyApplied") throw new Exception($"Changed blank was verified: existing={existing}, style={style}, status={r.Status}");
                Equal(1, r.Unverified);
                Equal(0, r.RemovedSoftLines); Equal(0, r.InsertedBlankLines);
            }
        });
        Test("code spacing strictly verifies every style of an existing mixed blank", () =>
        {
            foreach (var style in new[] { "font-size:20pt", "font-size:80pt", "font-family:Arial;font-size:20pt", "font-size:20pt;color:#FF0000",
                "font-size:20pt;font-weight:bold", "font-size:20pt;font-style:italic", "font-size:20pt;text-decoration:underline", null })
            {
                var blank = Paragraph("e", "&nbsp;<span style='font-size:20pt'>&nbsp;</span><br>&nbsp;<br>&nbsp;");
                blank.SetAttributeValue("style", "font-family:Calibri;font-size:11pt;color:#222222");
                var s = Snapshot(Page(Paragraph("a", "正文"), blank, SpacingBox("code"))); var t = Tools(s); Read(t, s); Normalize(t, s);
                var api = new FakePage(s.Page);
                api.AfterSave = () => AgentCommitter.Find(api.Page, "e").Element(One + "T").ReplaceNodes(new XCData(style == null ? "" :
                    "&nbsp;<span style='" + style + "'>&nbsp;</span>"));
                var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
                var preserved = style == "font-size:20pt";
                Equal(preserved ? "Verified" : "PartiallyApplied", r.Status); Equal(preserved ? 0 : 1, r.Unverified);
                Equal(preserved ? 2 : 0, r.RemovedSoftLines); Equal(0, r.InsertedBlankLines);
            }
        });
        Test("mixed blank styles survive native run splitting, verify, undo and later edit protection", () =>
        {
            foreach (var laterEdit in new[] { false, true })
            {
                var blank = Paragraph("e", "&nbsp;<span style='font-size:20pt'><b>&nbsp;</b></span><br>&nbsp;<br>&nbsp;");
                blank.SetAttributeValue("style", "font-family:Calibri;font-size:11pt;color:#222222");
                var s = Snapshot(Page(Paragraph("a", "正文"), blank, SpacingBox("code"))); var t = Tools(s); Read(t, s); Normalize(t, s);
                var api = new FakePage(s.Page);
                api.AfterSave = () =>
                {
                    var e = AgentCommitter.Find(api.Page, "e"); e.Elements(One + "T").Remove();
                    e.Add(new XElement(One + "T", new XCData("&nbsp;")),
                        new XElement(One + "T", new XCData("<strong><span style='font-size:20pt'>&nbsp;</span></strong>")));
                };
                var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
                Equal("Verified", r.Status); Equal(2, r.RemovedSoftLines); Equal(1, r.OutlineUndo.Count);
                Equal(AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(s.CreateDraftPage(), "e"), s.CreateDraftPage()),
                    AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "e"), api.Page));
                api.AfterSave = null;
                if (laterEdit) AgentCommitter.Find(api.Page, "e").SetAttributeValue("style", "font-family:Arial;font-size:30pt");
                var writes = api.Writes; var beforeUndo = api.Page.ToString();
                var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
                if (laterEdit) { True(undo.Conflicts > 0); Equal(writes, api.Writes); Equal(beforeUndo, api.Page.ToString()); }
                else
                {
                    Equal("Verified", undo.Status);
                    Equal(AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(s.Page, "e"), s.Page),
                        AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "e"), api.Page));
                }
            }
        });
        Test("blank projection permits uniform native runs but keeps mixed styles and empty-T formatting", () =>
        {
            string Format(params XElement[] runs)
            {
                var e = new XElement(One + "OE", new XAttribute("objectID", "e"), new XAttribute("style", "font-family:Calibri;font-size:11pt;color:#222222"), runs);
                var page = Page(e); return AgentPageSnapshot.BlankFormat(AgentCommitter.Find(page, "e"), page);
            }
            XElement Run(string html, string style = null) => new XElement(One + "T", style == null ? null : new XAttribute("style", style), new XCData(html));
            var uniform = Format(Run("&nbsp;"));
            Equal(uniform, Format(Run(""))); Equal(uniform, Format(Run(" ")));
            Equal(uniform, Format(Run("<span style='font-size:11pt'>&nbsp;</span>"), Run("&nbsp;")));
            Equal(uniform, Format(Run("", "font-size:11pt")));
            True(uniform != Format(Run("", "font-size:80pt")));
            True(uniform != Format(Run("&nbsp;"), Run("", "font-size:80pt")));
            var mixed = Format(Run("&nbsp;<span style='font-size:20pt'><b>&nbsp;</b></span>"));
            Equal(mixed, Format(Run("&nbsp;"), Run("<strong>&nbsp;</strong>", "font-size:20pt")));
            True(mixed != Format(Run("&nbsp;"), Run("<strong>&nbsp;</strong>", "font-size:80pt")));
            True(mixed != Format(Run("")));
        });
        Test("code spacing skips protected text on both sides and still normalizes independent boundaries", () =>
        {
            foreach (var beforeCode in new[] { false, true })
            foreach (var selected in new[] { false, true })
            foreach (var gap in new[] { (Paragraphs: 0, Soft: 0), (Paragraphs: 1, Soft: 0), (Paragraphs: 3, Soft: 0), (Paragraphs: 0, Soft: 3), (Paragraphs: 1, Soft: 2) })
            {
                var text = "调用 <span style='font-family:Consolas'>foo()</span>";
                var soft = string.Concat(Enumerable.Repeat(beforeCode ? "<br>&nbsp;" : "&nbsp;<br>", gap.Soft));
                var protectedText = Paragraph("protected", beforeCode ? text + soft : soft + text);
                var blanks = Enumerable.Range(0, gap.Paragraphs).Select(i => Paragraph("gap" + i, "&nbsp;")).ToList();
                foreach (var blank in blanks) blank.SetAttributeValue("style", "font-family:Consolas");
                var page = Page(beforeCode ? new[] { protectedText }.Concat(blanks).Concat(new[] { SpacingBox("code"), Paragraph("normal", "普通正文") }).ToArray() :
                    new[] { Paragraph("normal", "普通正文"), SpacingBox("code") }.Concat(blanks).Concat(new[] { protectedText }).ToArray());
                var selection = selected ? new HashSet<string>(page.Descendants(One + "OE").Attributes("objectID").Select(a => a.Value)) : null;
                var s = new AgentPageSnapshot(page.ToString(), selection, new AgentOptions());
                var api = new FakePage(s.Page); var t = new AgentTools(s, new AgentCommitter(api), CancellationToken.None); Read(t, s);
                Equal("protected_code", s.Blocks.Single(b => b.ObjectId == "protected").ProtectedReason);
                var result = Normalize(t, s);
                Equal(1, AiClient.Get(result, "inserted_paragraphs")); Equal(0, AiClient.Get(result, "removed_paragraphs")); Equal(0, AiClient.Get(result, "removed_soft_lines"));
                var settled = gap.Paragraphs + gap.Soft == 1;
                True(Json(result).Contains(settled ? "\"skipped\":[]" : "\"reason\":\"protected_text\""));
                Equal(0, AgentCodeSpacing.TrimCount(AgentCommitter.Find(s.Layout, "protected")));
                foreach (var blank in blanks) True(AgentCommitter.Find(s.Layout, (string)blank.Attribute("objectID")) != null);
                Invoke(t, "finish_edit", new { snapshot_id = s.SnapshotId, draft_revision = s.Revision });
                Equal("Verified", t.Report.Status); Equal(1, t.Report.InsertedBlankLines); Equal(settled ? 0 : 1, t.Report.SpacingSkipped.Length);
                Equal(AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(s.Page, "protected"), s.Page),
                    AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "protected"), api.Page));
                Equal("Verified", new AgentCommitter(api).Undo(s.PageId, t.Report, s.Options, CancellationToken.None).Status);
                Equal(Texts(s.Page), Texts(api.Page));
            }
        });
        Test("code spacing normalization does not exempt untouched or linked blank paragraphs", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "正文"), Paragraph("e", "&nbsp;<br>&nbsp;<br>&nbsp;"), SpacingBox("code"),
                Paragraph("gap", ""), Paragraph("b", "结尾"), Paragraph("untouched", "&nbsp;")));
            var t = Tools(s); Read(t, s); Normalize(t, s);
            var api = new FakePage(s.Page);
            api.AfterSave = () => { foreach (var id in new[] { "e", "untouched" }) AgentCommitter.Find(api.Page, id).Element(One + "T").ReplaceNodes(new XCData("")); };
            var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
            Equal("CommitOutcomeUnknown", r.Status); Equal(0, r.RemovedSoftLines); Equal(0, r.OutlineUndo.Count);
            var linked = Snapshot(Page(Paragraph("a", "正文"), Paragraph("e", "<a href='https://example.com'>&nbsp;</a><br>&nbsp;"), SpacingBox("code")));
            var lt = Tools(linked); Read(lt, linked); Normalize(lt, linked);
            var linkApi = new FakePage(linked.Page);
            linkApi.AfterSave = () => AgentCommitter.Find(linkApi.Page, "e").Element(One + "T").ReplaceNodes(new XCData(""));
            Equal("CommitOutcomeUnknown", new AgentCommitter(linkApi).Commit(linked, CancellationToken.None).Status);
            var styled = Snapshot(Page(Paragraph("a", "正文"), Paragraph("e", "&nbsp;"))); var st = Tools(styled); Read(st, styled); Style(st, styled);
            var styleApi = new FakePage(styled.Page);
            styleApi.AfterSave = () => AgentCommitter.Find(styleApi.Page, "e").Element(One + "T").ReplaceNodes(new XCData(""));
            Equal("CommitOutcomeUnknown", new AgentCommitter(styleApi).Commit(styled, CancellationToken.None).Status);
        });
        Test("blank format projection preserves soft-line counts while tolerating single-line whitespace normalization", () =>
        {
            string Format(string html) { var page = Page(Paragraph("e", html)); return AgentPageSnapshot.BlankFormat(AgentCommitter.Find(page, "e"), page); }
            var expected = Format("&nbsp;");
            foreach (var html in new[] { "", " ", "&nbsp;" }) Equal(expected, Format(html));
            foreach (var html in new[] { "&nbsp;<br>", "&nbsp;<br><br>" }) True(expected != Format(html));
        });
        Test("code spacing trims survive the appearance-only heading fallback", () =>
        {
            var parent = Paragraph("a", "<br><br>父标题"); parent.Add(new XElement(One + "OEChildren", Paragraph("b", "子段正文")));
            var s = Snapshot(Page(SpacingBox("code"), parent)); var t = Tools(s); Read(t, s); Normalize(t, s);
            var id = s.Blocks.Single(b => b.ObjectId == "a").Id;
            True(Json(Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { id }, preset_id = "heading1" })).Contains("\"appearance_only\":[\"" + id + "\"]"));
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var child = AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "b"), api.Page);
            var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(1, r.Applied); Equal(1, r.RemovedSoftLines);
            Equal("\n父标题", new AgentRichText(AgentCommitter.Find(api.Page, "a")).Text);
            Equal(child, AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "b"), api.Page));
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status);
            Equal("\n\n父标题", new AgentRichText(AgentCommitter.Find(api.Page, "a")).Text);
        });
        Test("structural tools after code spacing revert it so moved text and removed blanks stay unchanged", () =>
        {
            // 段内删行和补入的空段一起撤回；移走的段落不再挨着代码框，保持原样。
            var s = Snapshot(Page(Paragraph("a", "正文<br><br><br>"), SpacingBox("code"), Paragraph("b", "结尾"), Paragraph("c", "别处")));
            var api = new FakePage(s.Page); var t = new AgentTools(s, new AgentCommitter(api), CancellationToken.None); Read(t, s); Normalize(t, s);
            Equal(1, s.Inserted.Count);
            Move(t, s, new[] { s.Blocks.Single(b => b.ObjectId == "a").Id }, s.Blocks.Single(b => b.ObjectId == "c").Id, "after");
            Equal(0, s.Inserted.Count); Equal(0, s.Layout.Descendants(One + "OE").Sum(AgentCodeSpacing.TrimCount));
            True(Json(Invoke(t, "finish_edit", new { snapshot_id = s.SnapshotId, draft_revision = s.Revision })).Contains("normalize_code_spacing"));
            Normalize(t, s); Invoke(t, "finish_edit", new { snapshot_id = s.SnapshotId, draft_revision = s.Revision });
            Equal("Verified", t.Report.Status); Equal(0, t.Report.RemovedSoftLines); Equal(1, t.Report.InsertedBlankLines); Equal(1, t.Report.Moved);
            Equal("#||结尾|别处|正文", SpacingTexts(api.Page)); Equal("正文\n\n\n", new AgentRichText(AgentCommitter.Find(api.Page, "a")).Text);
            // 删掉的空段也恢复：移走正文后，代码框前的空行在文本框开头，不再处理。
            var blanks = Snapshot(Page(Paragraph("a", "正文"), Paragraph("e1", ""), Paragraph("e2", ""), SpacingBox("code"), Paragraph("b", "结尾")));
            var blankApi = new FakePage(blanks.Page); var bt = new AgentTools(blanks, new AgentCommitter(blankApi), CancellationToken.None); Read(bt, blanks);
            Equal(1, AiClient.Get(Normalize(bt, blanks), "removed_paragraphs")); True(AgentCommitter.Find(blanks.Layout, "e2") == null);
            Move(bt, blanks, new[] { blanks.Blocks.Single(b => b.ObjectId == "a").Id }, blanks.Blocks.Single(b => b.ObjectId == "b").Id, "after");
            True(AgentCommitter.Find(blanks.Layout, "e2") != null);
            Normalize(bt, blanks); Invoke(bt, "finish_edit", new { snapshot_id = blanks.SnapshotId, draft_revision = blanks.Revision });
            Equal("Verified", bt.Report.Status); Equal(0, bt.Report.Removed);
            True(AgentCommitter.Find(blankApi.Page, "e1") != null && AgentCommitter.Find(blankApi.Page, "e2") != null);
        });
        Test("structural calls without changes or with errors keep the code spacing draft", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "正文"), Paragraph("e1", ""), Paragraph("e2", ""), SpacingBox("code"), Paragraph("b", "结尾")));
            var t = Tools(s); Read(t, s); Normalize(t, s);
            var layout = new XElement(s.Layout); var revision = s.Revision; var changes = s.LayoutChanges.Count;
            var a = s.Blocks.Single(b => b.ObjectId == "a").Id;
            True(Json(Cleanup(t, s, a)).Contains("\"noop\":[\"" + a + "\"]"));
            True(XNode.DeepEquals(layout, s.Layout)); Equal(revision, s.Revision); Equal(1, s.Inserted.Count); Equal(changes, s.LayoutChanges.Count);
            Throws(() => Move(t, s, new[] { a }, "missing", "after"));
            True(XNode.DeepEquals(layout, s.Layout)); Equal(revision, s.Revision); Equal(1, s.Inserted.Count); Equal(changes, s.LayoutChanges.Count);
            True((bool)AiClient.Get(Normalize(t, s), "noop"));
            Equal("Verified", new AgentCommitter(new FakePage(s.Page)).Commit(s, CancellationToken.None).Status);
        });
        Test("final-round finish renormalizes code spacing instead of discarding the task", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "正文<br><br><br>"), SpacingBox("code"), Paragraph("b", "结尾"), Paragraph("c", "别处")));
            var api = new FakePage(s.Page); var t = new AgentTools(s, new AgentCommitter(api), CancellationToken.None); Read(t, s); Normalize(t, s);
            Move(t, s, new[] { s.Blocks.Single(b => b.ObjectId == "a").Id }, s.Blocks.Single(b => b.ObjectId == "c").Id, "after");
            t.AllowIncompleteFinish = true;
            Invoke(t, "finish_edit", new { snapshot_id = s.SnapshotId, draft_revision = s.Revision });
            Equal("Verified", t.Report.Status); Equal(1, t.Report.InsertedBlankLines); Equal(0, t.Report.RemovedSoftLines); Equal(1, t.Report.Moved);
            Equal("#||结尾|别处|正文", SpacingTexts(api.Page));
        });
        Test("code spacing does not report settled boundaries or unconverted code next to a code box", () =>
        {
            var page = Page(Paragraph("a", "正文"), Paragraph("e", ""), SpacingBox("code"), Paragraph("e2", ""), Paragraph("b", "结尾"));
            var selected = new AgentPageSnapshot(page.ToString(), new HashSet<string> { "a", "e", "code1", "code3" }, new AgentOptions());
            var t = Tools(selected); Read(t, selected); var result = Json(Normalize(t, selected));
            True(result.Contains("\"skipped\":[]") && result.Contains("\"noop\":true"));
            var raw = Paragraph("raw", "print(1)"); raw.SetAttributeValue("style", "font-family:Consolas");
            var code = Snapshot(Page(Paragraph("a", "正文"), Paragraph("e", ""), SpacingBox("code"), raw));
            var ct = Tools(code); Read(ct, code); True(Json(Normalize(ct, code)).Contains("\"skipped\":[]"));
            // 选区外的文字紧贴代码框时仍须说明。
            var tight = new AgentPageSnapshot(Page(Paragraph("a", "正文"), Paragraph("e", ""), SpacingBox("code"), Paragraph("b", "结尾")).ToString(),
                new HashSet<string> { "a", "e", "code1", "code3" }, new AgentOptions());
            var tt = Tools(tight); Read(tt, tight); True(Json(Normalize(tt, tight)).Contains("outside_selection"));
            // 已排入的表格转换和已有表格一样不算文字，不报跳过。
            var grid = Snapshot(Page(Paragraph("t1", "a\tb"), Paragraph("t2", "c\td"), SpacingBox("code"), Paragraph("b", "结尾")));
            var gt = Tools(grid); Read(gt, grid);
            Invoke(gt, "text_to_table", new { snapshot_id = grid.SnapshotId, block_ids = grid.Blocks.Where(b => b.ObjectId == "t1" || b.ObjectId == "t2").Select(b => b.Id).ToArray(), delimiter = "tab" });
            var gridResult = Normalize(gt, grid); True(Json(gridResult).Contains("\"skipped\":[]")); Equal(1, AiClient.Get(gridResult, "inserted_paragraphs"));
        });
        Test("text tools after code spacing count occurrences in the visible text and keep trim records valid", () =>
        {
            foreach (var occurrence in new[] { 1, 2 })
            {
                var s = Snapshot(Page(SpacingBox("code"), Paragraph("b", "&nbsp;<br>&nbsp;<br>正文&nbsp;结尾")));
                var api = new FakePage(s.Page); var t = new AgentTools(s, new AgentCommitter(api), CancellationToken.None); Read(t, s);
                Equal(1, AiClient.Get(Normalize(t, s), "removed_soft_lines"));
                var id = s.Blocks.Single(b => b.ObjectId == "b").Id;
                // 第 1 处是可见的那行空白，不是已删掉的首行；第 2 处在正文中间。
                var expected = occurrence == 1 ? "X\n正文 结尾" : " \n正文，结尾";
                var fixedResult = Json(Invoke(t, "fix_text", new { snapshot_id = s.SnapshotId, fixes = new[] { new { block_id = id, quote = " ", occurrence, replacement = occurrence == 1 ? "X" : "，" } } }));
                True(fixedResult.Contains(Json(expected)));
                True(Json(Invoke(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = new[] { id } })).Contains(Json(expected)));
                Equal(expected, new AgentRichText(AgentCommitter.Find(s.CreateDraftPage(), "b")).Text);
                // 可见首行不再是空白时要补一行空段落；正文中间的修改不影响间隔。
                var renormalized = Normalize(t, s); Equal(occurrence == 1 ? 1 : 0, AiClient.Get(renormalized, "inserted_paragraphs"));
                Invoke(t, "finish_edit", new { snapshot_id = s.SnapshotId, draft_revision = s.Revision });
                Equal("Verified", t.Report.Status); Equal(1, t.Report.TextFixes.Count);
                Equal(expected, new AgentRichText(AgentCommitter.Find(api.Page, "b")).Text);
            }
            var styled = Snapshot(Page(SpacingBox("code"), Paragraph("b", "&nbsp;<br>&nbsp;<br>正文&nbsp;结尾")));
            var styleApi = new FakePage(styled.Page); var st = new AgentTools(styled, new AgentCommitter(styleApi), CancellationToken.None); Read(st, styled); Normalize(st, styled);
            var block = styled.Blocks.Single(b => b.ObjectId == "b");
            Invoke(st, "set_text_style", new { snapshot_id = styled.SnapshotId, targets = new[] { new { block_id = block.Id, quote = " ", occurrence = 2, style = new { bold = true } } } });
            var html = string.Concat(block.Draft.Elements(One + "T").Select(x => x.Value));
            True(html.IndexOf("bold", StringComparison.Ordinal) > html.IndexOf("正文", StringComparison.Ordinal));
            Invoke(st, "finish_edit", new { snapshot_id = styled.SnapshotId, draft_revision = styled.Revision });
            Equal("Verified", st.Report.Status); Equal(1, st.Report.RemovedSoftLines);
        });
        Test("remove_blank_lines collapses runs and edges, verifies and undo restores the blank lines", () =>
        {
            var s = Snapshot(Page(Paragraph("e1", ""), Paragraph("a", "第一段"), Paragraph("e2", ""), Paragraph("e3", ""), Paragraph("b", "第二段"), Paragraph("e4", "")));
            var t = Tools(s);
            True(Json(Invoke(t, "remove_blank_lines", new { snapshot_id = s.SnapshotId, mode = "collapse" })).Contains("\"removed\":[\"p1\",\"p4\",\"p6\"]"));
            Equal(1, s.Revision); True(!Json(Invoke(t, "get_page_overview", new { })).Contains("\"id\":\"p1\""));
            True(Json(Invoke(t, "get_pending_changes", new { snapshot_id = s.SnapshotId })).Contains("\"removed\":[\"p1\",\"p4\",\"p6\"]"));
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(3, r.Removed); Equal(1, r.OutlineUndo.Count); Equal(1, api.Writes); True(r.Message.Contains("删除空行 3 行"));
            Equal("第一段||第二段", Texts(api.Page));
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal("Verified", undo.Status); Equal(1, undo.Outlines); True(undo.Message.Contains("恢复文本框结构 1 个"));
            Equal("|第一段|||第二段|", Texts(api.Page));
            True(AgentCommitter.Find(api.Page, "e2") != null && AgentCommitter.Find(api.Page, "e1") == null && AgentCommitter.Find(api.Page, "a") != null);
        });
        Test("remove_blank_lines skips marked blanks, code ranges and blanks outside the selection", () =>
        {
            var tagged = Paragraph("g", ""); tagged.AddFirst(Tag("0"));
            var p = Page(Paragraph("a", "正文"), Paragraph("e1", ""), Listed("l", "", "2"), Paragraph("e2", ""), tagged,
                Paragraph("c1", "x = 1"), Paragraph("c2", ""), Paragraph("c3", "y = 2"), Paragraph("b", "结尾"));
            p.AddFirst(TagDef("0", 3, "待办事项"));
            var s = Snapshot(p); var t = Tools(s); Read(t, s);
            Code(t, s, "python", "p6", "p7", "p8");
            True(Json(Invoke(t, "remove_blank_lines", new { snapshot_id = s.SnapshotId, mode = "all" })).Contains("\"removed\":[\"p2\",\"p4\"]"));
            var api = new FakePage(s.Page); var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(2, r.Removed); Equal(1, r.CodeBlocks); Equal(0, r.CodeUndo.Count); Equal(1, r.OutlineUndo.Count);
            // all 也不会把一摞段落删空；选区外的空行不删。
            var blank = Snapshot(Page(Paragraph("e1", ""), Paragraph("e2", "")));
            True(Json(Invoke(Tools(blank), "remove_blank_lines", new { snapshot_id = blank.SnapshotId, mode = "all" })).Contains("\"removed\":[\"p2\"]"));
            var selected = new AgentPageSnapshot(Page(Paragraph("a", "正文"), Paragraph("e1", ""), Paragraph("e2", ""), Paragraph("b", "结尾")).ToString(),
                new HashSet<string> { "a", "e1" }, new AgentOptions());
            True(Json(Invoke(Tools(selected), "remove_blank_lines", new { snapshot_id = selected.SnapshotId, mode = "collapse" })).Contains("\"removed\":[\"p2\"]"));
            Equal("正文||结尾", Texts(new FakePage(selected.Page).Commit(selected)));
        });
        Test("move_blocks moves a paragraph with its children, keeps IDs, verifies and undo restores order", () =>
        {
            var conclusion = Paragraph("c", "结论"); conclusion.Add(new XElement(One + "OEChildren", Paragraph("c1", "细节")));
            var s = Snapshot(Page(Paragraph("a", "背景"), Paragraph("b", "过程"), conclusion)); var t = Tools(s); Read(t, s);
            Move(t, s, new[] { "p3" }, "p1", "before"); Equal(1, s.Revision);
            Move(t, s, new[] { "p3" }, "p1", "before"); Equal(1, s.Revision);
            var overview = Json(Invoke(t, "get_page_overview", new { }));
            True(overview.IndexOf("\"id\":\"p3\"", StringComparison.Ordinal) < overview.IndexOf("\"id\":\"p1\"", StringComparison.Ordinal));
            True(overview.Contains("\"parent_id\":\"p3\""));
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(1, r.Moved); True(r.Message.Contains("移动 1 段"));
            Equal("结论|细节|背景|过程", Texts(api.Page)); Equal("c,c1,a,b", Ids(api.Page));
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal("Verified", undo.Status); Equal("背景|过程|结论|细节", Texts(api.Page)); Equal("a,b,c,c1", Ids(api.Page));
        });
        Test("move_blocks rejects unread, cross-cell, nested, title and self targets without side effects", () =>
        {
            var parent = Paragraph("n", "父"); parent.Add(new XElement(One + "OEChildren", Paragraph("m", "子")));
            var p = GridPage(); p.Element(One + "Outline").Element(One + "OEChildren").Add(Paragraph("a", "正文"), parent);
            p.AddFirst(new XElement(One + "Title", Paragraph("title", "标题")));
            var s = Snapshot(p); var t = Tools(s);
            Rejects("请先完整读取", () => Move(t, s, new[] { "p6" }, "p7", "after"));
            Read(t, s);
            Rejects("单元格", () => Move(t, s, new[] { "p6" }, "p2", "before"));
            Rejects("不要同时列出", () => Move(t, s, new[] { "p7", "p8" }, "p6", "before"));
            Rejects("下面", () => Move(t, s, new[] { "p7" }, "p8", "after"));
            Rejects("页面标题", () => Move(t, s, new[] { "p6" }, "p1", "after"));
            Rejects("不能是要移动的段落", () => Move(t, s, new[] { "p6" }, "p6", "after"));
            Equal(0, s.Revision); Equal(0, s.LayoutChanges.Count);
        });
        Test("set_indent in and out keep document order; undo restores nesting", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "一"), Paragraph("b", "二"), Paragraph("c", "三"), Paragraph("d", "四"))); var t = Tools(s); Read(t, s);
            Indent(t, s, "in", "p2", "p3");
            True(Json(Invoke(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p3" } })).Contains("\"parent_id\":\"p1\""));
            Indent(t, s, "out", "p2");
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(2, r.Indented); Equal("一|二|三|四", Texts(api.Page));
            var b = AgentCommitter.Find(api.Page, "b");
            Equal(0, b.Ancestors(One + "OE").Count()); Equal("c", (string)b.Element(One + "OEChildren").Element(One + "OE").Attribute("objectID"));
            Equal(0, AgentCommitter.Find(api.Page, "a").Elements(One + "OEChildren").Count());
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status);
            Equal(0, api.Page.Descendants(One + "OEChildren").Count(e => e.Parent.Name == One + "OE")); Equal("a,b,c,d", Ids(api.Page));
        });
        Test("set_indent rejects the first paragraph, outermost outdent and inherited parent style", () =>
        {
            var heading = Paragraph("h", "标题"); heading.SetAttributeValue("style", "font-size:20pt");
            var s = Snapshot(Page(Paragraph("a", "首段"), heading, Paragraph("b", "正文"))); var t = Tools(s); Read(t, s);
            Rejects("前面没有", () => Indent(t, s, "in", "p1"));
            Rejects("最外层", () => Indent(t, s, "out", "p2"));
            Rejects("继承", () => Indent(t, s, "in", "p3"));
            Equal(0, s.Revision);
        });
        Test("insert_blocks adds styled plain paragraphs that can be targets, commits and undo removes them", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "第一段"), Paragraph("b", "第二段"))); var t = Tools(s); Read(t, s);
            True(Json(Invoke(t, "insert_blocks", new { snapshot_id = s.SnapshotId, target_id = "p1", position = "before", paragraphs = new object[] {
                new { text = "摘要", preset_id = "heading2" }, new { text = "要点 <b>&", preset_id = "body", list = "bullet" } } })).Contains("\"inserted\":[\"n1\",\"n2\"]"));
            Insert(t, s, "n2", "补充");
            var overview = Json(Invoke(t, "get_page_overview", new { }));
            True(overview.IndexOf("\"id\":\"n3\"", StringComparison.Ordinal) < overview.IndexOf("\"id\":\"p1\"", StringComparison.Ordinal)); True(overview.Contains("\"inserted\":true"));
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(3, r.Inserted); True(!api.LastXml.Contains("urn:onenote-code-helper"));
            Equal("摘要|要点 <b>&|补充|第一段|第二段", Texts(api.Page));
            var lines = api.Page.Descendants(One + "OE").ToList();
            True(lines[1].Element(One + "T").Value.Contains("要点 &lt;b&gt;&amp;")); Equal("bullet", AgentMarks.ListKind(lines[1]));
            var index = (string)lines[0].Attribute("quickStyleIndex");
            Equal("h2", (string)api.Page.Elements(One + "QuickStyleDef").Single(d => (string)d.Attribute("index") == index).Attribute("name"));
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status); Equal("第一段|第二段", Texts(api.Page));
        });
        Test("insert_blocks rejects bad text, presets, targets and limits without side effects", () =>
        {
            var p = Page(Paragraph("a", "正文")); p.AddFirst(new XElement(One + "Title", Paragraph("title", "标题")));
            var s = Snapshot(p); var t = Tools(s);
            Rejects("换行", () => Insert(t, s, "p2", "第一行\n第二行"));
            Rejects("空白", () => Insert(t, s, "p2", "   "));
            Rejects("字符串无效", () => Invoke(t, "insert_blocks", new { snapshot_id = s.SnapshotId, target_id = "p2", position = "after", paragraphs = new[] { new { text = "x", preset_id = "page_title" } } }));
            Rejects("页面标题", () => Insert(t, s, "p1", "新段落"));
            Rejects("不存在", () => Insert(t, s, "n9", "新段落"));
            var many = Enumerable.Range(1, 21).Select(i => new { text = "第" + i + "项", preset_id = "body" }).ToArray();
            Rejects("数组长度", () => Invoke(t, "insert_blocks", new { snapshot_id = s.SnapshotId, target_id = "p2", position = "after", paragraphs = many }));
            Equal(0, s.Revision);
            for (var i = 0; i < 2; i++) Invoke(t, "insert_blocks", new { snapshot_id = s.SnapshotId, target_id = "p2", position = "after", paragraphs = many.Take(20).ToArray() });
            Rejects("最多插入", () => Invoke(t, "insert_blocks", new { snapshot_id = s.SnapshotId, target_id = "p2", position = "after", paragraphs = many.Take(20).ToArray() }));
            Equal(2, s.Revision); Equal(40, s.Inserted.Count);
        });
        Test("insert_blocks blank items add body-look blank lines that verify as blank lines and undo removes them", () =>
        {
            var s = Snapshot(); var t = Tools(s); Read(t, s);
            var result = Json(Invoke(t, "insert_blocks", new { snapshot_id = s.SnapshotId, target_id = "p1", position = "after", paragraphs = new object[] { new { blank = true } } }));
            True(result.Contains("\"inserted\":[\"n1\"]")); True(result.Contains("\"blank_lines\":[\"n1\"]"));
            True(Json(Invoke(t, "get_page_overview", new { })).Contains("\"blank\":true"));
            var draft = AgentLayout.Find(s.Layout, "n1");
            Equal(null, (string)draft.Attribute("quickStyleIndex")); Equal("0", (string)draft.Attribute("spaceBefore")); Equal("0", (string)draft.Attribute("spaceAfter"));
            Insert(t, s, "n1", "补充");
            True(Json(Invoke(t, "get_pending_changes", new { snapshot_id = s.SnapshotId })).Contains("\"inserted\":[{\"id\":\"n1\",\"text\":\"\"},{\"id\":\"n2\""));
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(1, r.InsertedBlankLines); Equal(1, r.Inserted); Equal(1, r.OutlineUndo.Count);
            True(r.Message.Contains("补空行 1 行")); True(r.Message.Contains("插入 1 段"));
            Equal("第一段||补充|第二段", Texts(api.Page));
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status); Equal("第一段|第二段", Texts(api.Page));
        });
        Test("insert_blocks blank items reject extra fields, missing text and empty items; blank lines count toward the quota", () =>
        {
            var s = Snapshot(); var t = Tools(s); Read(t, s);
            object Add(params object[] items) => Invoke(t, "insert_blocks", new { snapshot_id = s.SnapshotId, target_id = "p1", position = "after", paragraphs = items });
            Rejects("空行项", () => Add(new { blank = true, text = "x" }));
            Rejects("空行项", () => Add(new { blank = true, preset_id = "body" }));
            Rejects("要么", () => Add(new { preset_id = "body" }));
            Rejects("要么", () => Add(new { blank = false }));
            Rejects("字段不符合", () => Add(new { }));
            Rejects("blank: true", () => Add(new { text = "  ", preset_id = "body" }));
            Equal(0, s.Revision);
            var blanks = Enumerable.Range(0, 20).Select(i => (object)new { blank = true }).ToArray();
            for (var i = 0; i < 2; i++) Add(blanks);
            Rejects("最多插入", () => Add(blanks));
            Equal(2, s.Revision); Equal(40, s.Inserted.Count);
        });
        Test("blank cleanup is available before insertion and removes generated blanks in the same task", () =>
        {
            foreach (var removal in new[] { false, true }) foreach (var insertion in new[] { false, true }) foreach (var existing in new[] { false, true })
            {
                var page = existing ? Page(Paragraph("a", "正文"), Paragraph("e", "")) : Page(Paragraph("a", "正文"));
                var tools = Tools(new AgentPageSnapshot(page.ToString(), null, new AgentOptions { EnableBlankLineRemoval = removal, EnableInsert = insertion }));
                var available = removal && (insertion || existing);
                Equal(available, tools.Has("remove_blank_lines"));
                Equal(available, AiClient.Get(AgentChatClient.Parse(Json(Invoke(tools, "get_page_overview", new { }))), "blank_lines"));
                Equal(available, AgentRunner.SystemPrompt(tools).Contains(AgentRunner.BlankLinePrompt));
            }
            var s = Snapshot(); var t = Tools(s); Read(t, s);
            InsertBlanks(t, s, "p1", 2);
            Equal("n1,n2", string.Join(",", (string[])ToolField(RemoveBlanks(t, s), "removed")));
            Equal("第一段|第二段", Texts(s.Layout)); Equal(2, s.Inserted.Count);
            True(Json(Invoke(t, "get_pending_changes", new { snapshot_id = s.SnapshotId })).Contains("\"inserted\":[]"));
            var revision = s.Revision; RemoveBlanks(t, s); Equal(revision, s.Revision);
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(0, r.Removed); Equal(0, r.InsertedBlankLines); Equal("第一段|第二段", Texts(api.Page));
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status); Equal("第一段|第二段", Texts(api.Page));
        });
        Test("blank cleanup collapses mixed original and generated blanks and counts only original deletions", () =>
        {
            foreach (var mode in new[] { "collapse", "all" })
            {
                var s = Snapshot(Page(Paragraph("head", ""), Paragraph("a", "第一段"), Paragraph("middle", ""), Paragraph("b", "第二段"), Paragraph("tail", "")));
                var t = Tools(s); Read(t, s);
                InsertBlanks(t, s, "p2", 1, "before"); InsertBlanks(t, s, "p4", 2, "before"); InsertBlanks(t, s, "p4");
                var removed = (string[])ToolField(RemoveBlanks(t, s, mode), "removed");
                Equal(mode == "collapse" ? 6 : 7, removed.Length); True(new[] { "n1", "n2", "n3", "n4" }.All(removed.Contains));
                var expected = mode == "collapse" ? "第一段||第二段" : "第一段|第二段";
                Equal(expected, Texts(s.Layout));
                var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
                Equal("Verified", r.Status); Equal(mode == "collapse" ? 2 : 3, r.Removed); Equal(0, r.InsertedBlankLines); Equal(expected, Texts(api.Page));
                Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status); Equal(Texts(s.Page), Texts(api.Page));
            }
        });
        Test("collapse keeps one generated blank and reports only the surviving insertion", () =>
        {
            var s = Snapshot(); var t = Tools(s); Read(t, s); InsertBlanks(t, s, "p1", 2);
            Equal("n2", string.Join(",", (string[])ToolField(RemoveBlanks(t, s, "collapse"), "removed")));
            var pending = Json(Invoke(t, "get_pending_changes", new { snapshot_id = s.SnapshotId }));
            True(pending.Contains("\"inserted\":[{\"id\":\"n1\",\"text\":\"\"}]"));
            True(!Json(Invoke(t, "get_page_overview", new { })).Contains("\"id\":\"n2\""));
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(0, r.Removed); Equal(1, r.InsertedBlankLines); Equal("第一段||第二段", Texts(api.Page));
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status); Equal("第一段|第二段", Texts(api.Page));
        });
        Test("deleted generated blanks retain their IDs and consume the task insertion quota", () =>
        {
            var s = Snapshot(); var t = Tools(s); var total = 0;
            foreach (var count in new[] { 20, 20, 10 })
            {
                var added = (string[])ToolField(InsertBlanks(t, s, "p1", count), "inserted");
                Equal("n" + (total + 1), added[0]); total += count; Equal("n" + total, added.Last());
                Equal(count, ((string[])ToolField(RemoveBlanks(t, s), "removed")).Length);
                Equal(total, s.Inserted.Count); Equal("第一段|第二段", Texts(s.Layout));
            }
            Equal(50, s.Inserted.Select(i => i.Id).Distinct().Count());
            var state = MergeDraftState(s);
            Rejects("不存在", () => InsertBlanks(t, s, "n1"));
            Rejects("最多插入", () => InsertBlanks(t, s, "p1")); Equal(state, MergeDraftState(s));
        });
        Test("generated blank cleanup preserves unselected, marked and code blanks", () =>
        {
            var tagged = Paragraph("tagged", ""); tagged.AddFirst(Tag("0"));
            var parent = Paragraph("parent", ""); parent.Add(new XElement(One + "OEChildren", Paragraph("child", "细节")));
            var mono = Paragraph("mono", "&nbsp;"); mono.SetAttributeValue("style", "font-family:Consolas");
            var page = Page(Paragraph("a", "正文"), Paragraph("outside", ""), Paragraph("b", "结尾"), Listed("listed", "", "2"), tagged, parent, mono, SpacingBox("code"));
            page.AddFirst(TagDef("0", 3, "待办事项"));
            var selection = new HashSet<string>(page.Descendants(One + "OE").Select(e => (string)e.Attribute("objectID"))); selection.Remove("outside");
            var s = new AgentPageSnapshot(page.ToString(), selection, new AgentOptions()); var t = Tools(s); var original = new XElement(s.Layout);
            InsertBlanks(t, s, "p1", 2);
            Equal("n1,n2", string.Join(",", (string[])ToolField(RemoveBlanks(t, s, "collapse"), "removed")));
            RemoveBlanks(t, s); True(XNode.DeepEquals(original, s.Layout));
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(0, r.Removed); Equal(0, r.InsertedBlankLines); Equal(Texts(s.Page), Texts(api.Page));
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status); Equal(Texts(s.Page), Texts(api.Page));
            // 新空行成为上级段落后，也不能连同它的子段落一起删除。
            var nested = Snapshot(); var nt = Tools(nested); InsertBlanks(nt, nested, "p1"); Insert(nt, nested, "n1", "新细节");
            Indent(nt, nested, "in", "n2"); InsertBlanks(nt, nested, "p2", 1, "before");
            Equal("n3", string.Join(",", (string[])ToolField(RemoveBlanks(nt, nested), "removed")));
            True(AgentLayout.Find(nested.Layout, "n2").Ancestors(One + "OE").Contains(AgentLayout.Find(nested.Layout, "n1")));
        });
        Test("generated blank cleanup follows cross-frame moves and merges with grouped commit, conflict and undo", () =>
        {
            foreach (var operation in new[] { "move", "merge" }) foreach (var editTime in new[] { "none", "during", "after" })
            {
                var s = Snapshot(Boxes(Box("A", 100, Paragraph("a", "甲")), Box("B", 300, Paragraph("b", "乙"), Paragraph("e", ""), Paragraph("b2", "乙二"))));
                var t = Tools(s); Read(t, s); InsertBlanks(t, s, "p2", 2);
                if (operation == "merge") Merge(t, s, "B", "p1", "after");
                else Move(t, s, new[] { "n1", "n2", "p3" }, "p1", "after");
                var removed = (string[])ToolField(RemoveBlanks(t, s), "removed"); Equal(3, removed.Length); True(removed.Contains("p3"));
                var api = new FakePage(s.Page); var c = new AgentCommitter(api);
                if (editTime == "during") AgentCommitter.Find(api.Page, "b").Element(One + "T").Value = "处理中编辑";
                var r = c.Commit(s, CancellationToken.None);
                if (editTime == "during")
                {
                    Equal("NoChange", r.Status); True(r.Conflicts > 0); Equal(0, api.Writes); Equal(0, r.Removed); Equal(0, r.InsertedBlankLines);
                    Equal("甲|处理中编辑||乙二", Texts(api.Page)); continue;
                }
                Equal("Verified", r.Status); Equal(1, api.Writes); Equal(1, r.Removed); Equal(0, r.InsertedBlankLines); Equal(2, r.OutlineUndo.Count);
                Equal("甲|乙|乙二", Texts(api.Page));
                if (editTime == "after") api.Page.Descendants(One + "OE").Single(e => AgentCode.PlainText(e) == "乙").Element(One + "T").Value = "提交后编辑";
                var beforeUndo = api.Page.ToString(); var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
                if (editTime == "after") { Equal("NoChange", undo.Status); True(undo.Conflicts > 0); Equal(1, api.Writes); Equal(beforeUndo, api.Page.ToString()); }
                else { Equal("Verified", undo.Status); Equal(2, api.Page.Elements(One + "Outline").Count()); Equal(Texts(s.Page), Texts(api.Page)); }
            }
        });
        Test("manual blank cleanup reverts code spacing first and safely renormalizes before commit", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "正文"), SpacingBox("code"), Paragraph("b", "结尾")));
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var t = new AgentTools(s, c, CancellationToken.None); Read(t, s);
            InsertBlanks(t, s, "p1", 1, "before"); InsertBlanks(t, s, "p1", 2);
            Normalize(t, s); Equal(4, s.Inserted.Count); True(AgentLayout.Find(s.Layout, "n3") == null);
            Equal("n1,n2,n3", string.Join(",", (string[])ToolField(RemoveBlanks(t, s), "removed")));
            Equal(3, s.Inserted.Count); Equal("正文|#|结尾", SpacingTexts(s.Layout)); True(AgentCommitter.Find(s.Layout, "code2") != null);
            True(Json(Invoke(t, "finish_edit", new { snapshot_id = s.SnapshotId, draft_revision = s.Revision })).Contains("normalize_code_spacing")); Equal(0, api.Writes);
            Normalize(t, s); Invoke(t, "finish_edit", new { snapshot_id = s.SnapshotId, draft_revision = s.Revision });
            Equal("Verified", t.Report.Status); Equal(0, t.Report.Removed); Equal(2, t.Report.InsertedBlankLines); Equal("正文||#||结尾", SpacingTexts(api.Page));
            Equal("Verified", c.Undo(s.PageId, t.Report, s.Options, CancellationToken.None).Status); Equal("正文|#|结尾", SpacingTexts(api.Page));
            // 没有手动空行可删时，清理无改动，已经规范化的草稿应完整保留。
            var unchanged = Snapshot(s.Page); var ut = Tools(unchanged); Read(ut, unchanged); Normalize(ut, unchanged);
            var state = MergeDraftState(unchanged); Equal(0, ((string[])ToolField(RemoveBlanks(ut, unchanged), "removed")).Length); Equal(state, MergeDraftState(unchanged));
        });
        Test("cancelled generated blank cleanup preserves the complete code spacing draft", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "正文"), SpacingBox("code"), Paragraph("b", "结尾"))); var api = new FakePage(s.Page);
            using (var cancel = new CancellationTokenSource())
            {
                var t = new AgentTools(s, new AgentCommitter(api), cancel.Token); Read(t, s); InsertBlanks(t, s, "p1", 2); Normalize(t, s);
                var state = MergeDraftState(s); cancel.Cancel();
                Throws<OperationCanceledException>(() => RemoveBlanks(t, s)); Equal(state, MergeDraftState(s)); Equal(0, api.Writes);
            }
        });
        Test("text_to_table converts pipe rows keeping links and bold, verifies and undo restores paragraphs", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "对比如下："), Paragraph("r1", "| 名称 | 说明 |"), Paragraph("r2", "|---|:--:|"),
                Paragraph("r3", "| <b>甲</b> | 见<a href='https://example.com'>链接</a> |"), Paragraph("b", "结尾")));
            var t = Tools(s); Read(t, s);
            var result = Json(Table(t, s, "pipe", "p2", "p3", "p4"));
            True(result.Contains("\"rows\":2")); True(result.Contains("\"columns\":2"));
            Rejects("表格转换", () => Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p2" }, preset_id = "body" }));
            True(Json(Invoke(t, "get_pending_changes", new { snapshot_id = s.SnapshotId })).Contains("\"text_tables\":[{\"block_ids\":[\"p2\",\"p3\",\"p4\"]}]"));
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(1, r.TextTables); Equal(1, r.CodeUndo.Count); True(r.Message.Contains("转换表格 1 个"));
            var table = api.Page.Descendants(One + "Table").Single();
            True(TableLook.Flag(table, "hasHeaderRow")); True(TableLook.Flag(table, "bordersVisible"));
            Equal("#DEEAF6", (string)table.Element(One + "Row").Element(One + "Cell").Attribute("shadingColor"));
            var cells = table.Descendants(One + "Cell").Select(cell => cell.Descendants(One + "OE").Single()).ToList();
            Equal("名称|说明|甲|见链接", string.Join("|", cells.Select(AgentCode.PlainText)));
            True(cells[2].Element(One + "T").Value.Contains("<b>甲</b>")); True(cells[3].Element(One + "T").Value.Contains("href=\"https://example.com\""));
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal("Verified", undo.Status); Equal(1, undo.TextTables); True(undo.Message.Contains("表格换回段落 1 个"));
            True(!api.Page.Descendants(One + "Table").Any()); Equal(Texts(s.Page), Texts(api.Page));
        });
        Test("text_to_table splits tabs, pads short rows and rejects invalid input", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "名称\t说明\t备注"), Paragraph("b", "甲\t乙"), Paragraph("c", "单列"), Paragraph("d", "第一行<br>第二行\tx"),
                Paragraph("e", "一\t二"), Listed("l", "列\t表", "2")));
            var t = Tools(s);
            Rejects("请先完整读取", () => Table(t, s, "tab", "p1", "p2"));
            Read(t, s);
            Rejects("两列", () => Table(t, s, "tab", "p3"));
            Rejects("连续", () => Table(t, s, "tab", "p1", "p5"));
            Rejects("项目符号", () => Table(t, s, "tab", "p6"));
            Equal(0, s.Revision);
            True(Json(Table(t, s, "tab", "p1", "p2")).Contains("\"columns\":3"));
            var api = new FakePage(s.Page); Equal("Verified", new AgentCommitter(api).Commit(s, CancellationToken.None).Status);
            var cells = api.Page.Descendants(One + "Cell").Select(cell => AgentCode.PlainText(cell.Descendants(One + "OE").Single())).ToArray();
            Equal("名称|说明|备注|甲|乙|", string.Join("|", cells));
        });
        Test("text_to_table splits space-separated rows, adds a header row and keeps the title paragraph", () =>
        {
            var s = Snapshot(Page(Paragraph("t", "远程办公和不关机后端名单"), Paragraph("a", "李云星 10.28.2.16"), Paragraph("b", "尹志远&nbsp;&nbsp;10.28.1.106"),
                Paragraph("c", "<b>周涛</b>　 10.28.2.46")));
            var t = Tools(s); Read(t, s);
            var result = Json(Invoke(t, "text_to_table", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p2", "p3", "p4" }, delimiter = "space", header = new[] { "姓名", "IP" } }));
            True(result.Contains("\"rows\":4")); True(result.Contains("\"columns\":2")); True(result.Contains("\"padded_rows\":0"));
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(1, r.TextTables);
            var table = api.Page.Descendants(One + "Table").Single();
            True(TableLook.Flag(table, "hasHeaderRow"));
            var cells = table.Descendants(One + "Cell").Select(cell => cell.Descendants(One + "OE").Single()).ToList();
            Equal("姓名|IP|李云星|10.28.2.16|尹志远|10.28.1.106|周涛|10.28.2.46", string.Join("|", cells.Select(AgentCode.PlainText)));
            True(cells[6].Element(One + "T").Value.Contains("<b>周涛</b>"));
            Equal("远程办公和不关机后端名单", AgentCode.PlainText(api.Page.Descendants(One + "OE").First()));
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal("Verified", undo.Status); True(!api.Page.Descendants(One + "Table").Any()); Equal(Texts(s.Page), Texts(api.Page));
        });
        Test("text_to_table pads uneven rows and headers with empty cells", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "李云星 10.28.2.16<br>尹志远 10.28.1.106"), Paragraph("b", "名单"), Paragraph("c", "周涛 10.28.2.46"),
                Paragraph("d", "刘书康 10.28.1.89"), Paragraph("e", "标题<br>姬仁洋 10.28.2.62")));
            var t = Tools(s); Read(t, s);
            string Named(string[] header, params string[] ids) =>
                Json(Invoke(t, "text_to_table", new { snapshot_id = s.SnapshotId, block_ids = ids, delimiter = "space", header }));
            Rejects("两列", () => Table(t, s, "space", "p2"));
            Rejects("header 的每一项", () => Named(new[] { "姓\t名", "IP" }, "p3", "p4"));
            Equal(0, s.Revision);
            var plain = Json(Invoke(t, "text_to_table", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" }, delimiter = "space", header_row = false }));
            True(plain.Contains("\"rows\":2")); True(plain.Contains("\"columns\":2")); True(plain.Contains("\"padded_rows\":0"));
            // 标题行和少给的列名补空。
            var titled = Named(new[] { "姓名" }, "p2", "p3", "p4");
            True(titled.Contains("\"rows\":4")); True(titled.Contains("\"columns\":2")); True(titled.Contains("\"padded_rows\":2"));
            // 多给的列名加一列，数据行补空。
            var wider = Named(new[] { "姓名", "IP", "备注" }, "p5");
            True(wider.Contains("\"rows\":3")); True(wider.Contains("\"columns\":3")); True(wider.Contains("\"padded_rows\":2"));
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(3, r.TextTables);
            var tables = api.Page.Descendants(One + "Table").ToList();
            Equal(3, tables.Count); True(!TableLook.Flag(tables[0], "hasHeaderRow"));
            string Cells(XElement table) => string.Join("|", table.Descendants(One + "Cell").Select(cell => AgentCode.PlainText(cell.Descendants(One + "OE").Single())));
            Equal("李云星|10.28.2.16|尹志远|10.28.1.106", Cells(tables[0]));
            Equal("姓名||名单||周涛|10.28.2.46|刘书康|10.28.1.89", Cells(tables[1]));
            Equal("姓名|IP|备注|标题|||姬仁洋|10.28.2.62|", Cells(tables[2]));
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal("Verified", undo.Status); True(!api.Page.Descendants(One + "Table").Any()); Equal(Texts(s.Page), Texts(api.Page));
        });
        Test("text_to_table joins multi-line records into rows, drops label colons and keeps links", () =>
        {
            const string login = "http://login.example.com", center = "http://center.example.com";
            var s = Snapshot(Page(Paragraph("a", "登录服"), Paragraph("b", ""), Paragraph("c", $"<a href=\"{login}\">{login}</a>"), Paragraph("d", ""),
                Paragraph("e", "中心服"), Paragraph("f", ""), Paragraph("g", center), Paragraph("h", ""),
                Paragraph("i", "充值回调地址："), Paragraph("j", ""), Paragraph("k", center + "/recharge")));
            var t = Tools(s); Read(t, s);
            var all = Enumerable.Range(1, 11).Select(i => "p" + i).ToArray();
            Rejects("lines_per_row", () => Table(t, s, "none", all));
            var result = Json(Invoke(t, "text_to_table", new { snapshot_id = s.SnapshotId, block_ids = all, delimiter = "none", lines_per_row = 2, header = new[] { "名称", "地址" } }));
            True(result.Contains("\"rows\":4")); True(result.Contains("\"columns\":2")); True(result.Contains("\"padded_rows\":0"));
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(1, r.TextTables);
            var cells = api.Page.Descendants(One + "Table").Single().Descendants(One + "Cell").Select(cell => cell.Descendants(One + "OE").Single()).ToList();
            Equal($"名称|地址|登录服|{login}|中心服|{center}|充值回调地址|{center}/recharge", string.Join("|", cells.Select(AgentCode.PlainText)));
            True(cells[3].Element(One + "T").Value.Contains("href=\"" + login + "\""));
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal("Verified", undo.Status); True(!api.Page.Descendants(One + "Table").Any()); Equal(Texts(s.Page), Texts(api.Page));
            // 缺了最后一个值的记录补空。
            var odd = Snapshot(Page(Paragraph("a", "登录服"), Paragraph("b", login), Paragraph("c", "中心服：")));
            var ot = Tools(odd); Read(ot, odd);
            True(Json(Invoke(ot, "text_to_table", new { snapshot_id = odd.SnapshotId, block_ids = new[] { "p1", "p2", "p3" }, delimiter = "none", lines_per_row = 2, header_row = false }))
                .Contains("\"padded_rows\":1"));
            var oapi = new FakePage(odd.Page); Equal("Verified", new AgentCommitter(oapi).Commit(odd, CancellationToken.None).Status);
            Equal($"登录服|{login}|中心服|", string.Join("|", oapi.Page.Descendants(One + "Cell").Select(cell => AgentCode.PlainText(cell.Descendants(One + "OE").Single()))));
        });
        Test("text_to_table builds rows given by the model and verifies them against the original", () =>
        {
            const string link = "http://c.example.com";
            var s = Snapshot(Page(Paragraph("a", "服务器：A，IP：10.0.0.1，端口：80"), Paragraph("b", ""), Paragraph("c", "服务器: B&nbsp;1  IP: 10.0.0.2"),
                Paragraph("d", $"<a href=\"{link}\">C</a>&nbsp;10.0.0.3 端口=81"), Paragraph("e", "D 端口：82")));
            var t = Tools(s); Read(t, s);
            var ids = new[] { "p1", "p2", "p3", "p4", "p5" };
            var header = new[] { "服务器", "IP", "端口" };
            Func<string[][], object> call = rows => Invoke(t, "text_to_table", new { snapshot_id = s.SnapshotId, block_ids = ids, rows, header });
            var good = new[] { new[] { "A", "10.0.0.1", "80" }, new[] { "B 1", "10.0.0.2" }, new[] { "C", "10.0.0.3", "81" }, new[] { "D", "", "82" } };
            // 改写、调换、重复、漏掉文字，或省略的标签不是列名，都不通过。
            Rejects("没有在原文里按顺序找到", () => call(new[] { new[] { "A", "10.0.0.9", "80" } }.Concat(good.Skip(1)).ToArray()));
            Rejects("没有在原文里按顺序找到", () => call(new[] { good[0], good[0] }.Concat(good.Skip(1)).ToArray()));
            Rejects("没有放进任何单元格", () => call(new[] { new[] { "10.0.0.1", "A", "80" } }.Concat(good.Skip(1)).ToArray()));
            Rejects("原文「D 端口：82」没有放进", () => call(good.Take(3).ToArray()));
            Rejects("原文「服务器：A", () => Invoke(t, "text_to_table", new { snapshot_id = s.SnapshotId, block_ids = ids, rows = good }));
            Rejects("跨行", () => call(new[] { new[] { "A\n10.0.0.1" } }));
            Rejects("二选一", () => Invoke(t, "text_to_table", new { snapshot_id = s.SnapshotId, block_ids = ids, rows = good, delimiter = "space" }));
            Rejects("二选一", () => Invoke(t, "text_to_table", new { snapshot_id = s.SnapshotId, block_ids = ids }));
            Rejects("lines_per_row 只能", () => Invoke(t, "text_to_table", new { snapshot_id = s.SnapshotId, block_ids = ids, rows = good, lines_per_row = 2 }));
            Equal(0, s.Revision);
            var result = Json(call(good));
            True(result.Contains("\"rows\":5")); True(result.Contains("\"columns\":3")); True(result.Contains("\"padded_rows\":1"));
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(1, r.TextTables);
            var cells = api.Page.Descendants(One + "Table").Single().Descendants(One + "Cell").Select(cell => cell.Descendants(One + "OE").Single()).ToList();
            Equal("服务器|IP|端口|A|10.0.0.1|80|B 1|10.0.0.2||C|10.0.0.3|81|D||82", string.Join("|", cells.Select(AgentCode.PlainText)));
            True(cells[9].Element(One + "T").Value.Contains("href=\"" + link + "\""));
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal("Verified", undo.Status); True(!api.Page.Descendants(One + "Table").Any()); Equal(Texts(s.Page), Texts(api.Page));
        });
        Test("structure, format and code conversion in one text box commit together and undo restores all at once", () =>
        {
            var s = Snapshot(CodePage()); var t = Tools(s); Read(t, s);
            Code(t, s, "python", "p2", "p3", "p4");
            Style(t, s);
            Move(t, s, new[] { "p5" }, "p1", "before");
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(1, api.Writes); Equal(1, r.Applied); Equal(1, r.CodeBlocks); Equal(1, r.Moved);
            Equal(1, r.OutlineUndo.Count); Equal(0, r.Undo.Count + r.CodeUndo.Count);
            Equal("结尾", AgentCode.PlainText(api.Page.Descendants(One + "OE").First())); True(api.Page.Descendants(One + "Table").Any());
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal("Verified", undo.Status); Equal(2, api.Writes); True(!api.Page.Descendants(One + "Table").Any());
            Equal("示例：|def f(x):||    return x + 1|结尾", Texts(api.Page));
            Equal((string)AgentCommitter.Find(s.Page, "a").Attribute("style"), (string)AgentCommitter.Find(api.Page, "a").Attribute("style"));
        });
        Test("user edit in the same text box skips structure changes but still applies formats", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "第一段"), Paragraph("b", "第二段"), Paragraph("c", "第三段"))); var t = Tools(s); Read(t, s);
            Style(t, s); Move(t, s, new[] { "p3" }, "p2", "before");
            var api = new FakePage(s.Page); AgentCommitter.Find(api.Page, "b").Element(One + "T").Value = "用户改了第二段";
            var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
            Equal("PartiallyApplied", r.Status); Equal(1, r.Applied); Equal(0, r.Moved); True(r.ConflictIds.Contains("p3"));
            Equal("第一段|用户改了第二段|第三段", Texts(api.Page)); Equal(1, r.Undo.Count); Equal(0, r.OutlineUndo.Count);
        });
        Test("undo skips a text box edited after the structure change", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "第一段"), Paragraph("b", "第二段"))); var t = Tools(s); Read(t, s);
            Move(t, s, new[] { "p2" }, "p1", "before");
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status);
            AgentCommitter.Find(api.Page, "a").Element(One + "T").Value = "之后改过";
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal("NoChange", undo.Status); Equal(1, undo.Conflicts); Equal(1, api.Writes); Equal("第二段|之后改过", Texts(api.Page));
        });
        Test("OneNote changing a moved paragraph's format is not reported verified", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "第一段"), Paragraph("b", "第二段"))); var t = Tools(s); Read(t, s);
            Move(t, s, new[] { "p2" }, "p1", "before");
            var api = new FakePage(s.Page); api.AfterSave = () => AgentCommitter.Find(api.Page, "b").SetAttributeValue("style", "font-size:30pt");
            Equal("CommitOutcomeUnknown", new AgentCommitter(api).Commit(s, CancellationToken.None).Status);
        });
        Test("structure changes that would break a staged conversion are rejected", () =>
        {
            var s = Snapshot(CodePage()); var t = Tools(s); Read(t, s); Code(t, s, "python", "p2", "p3", "p4");
            Rejects("打断", () => Indent(t, s, "in", "p5"));
            Rejects("代码框或表格转换", () => Move(t, s, new[] { "p3" }, "p1", "before"));
            Rejects("代码框或表格转换", () => Move(t, s, new[] { "p5" }, "p2", "after"));
            Equal(1, s.Revision); Equal(0, s.LayoutChanges.Count);
        });
        Test("lists in a restructured text box: removal rebuilds, undo restores both lists", () =>
        {
            var s = Snapshot(Page(Listed("l", "已有列表", "2"), Paragraph("a", "第一段"), Paragraph("e", ""), Paragraph("b", "第二段"))); var t = Tools(s); Read(t, s);
            Invoke(t, "set_list", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" }, list = "none" });
            Invoke(t, "set_list", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p2" }, list = "bullet" });
            Invoke(t, "remove_blank_lines", new { snapshot_id = s.SnapshotId, mode = "all" });
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(2, r.Applied); Equal(1, r.Removed);
            True(AgentCommitter.Find(api.Page, "l") == null); Equal("none|bullet|none", string.Join("|", api.Page.Descendants(One + "OE").Select(AgentMarks.ListKind)));
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status);
            Equal("bullet|none|none|none", string.Join("|", api.Page.Descendants(One + "OE").Select(AgentMarks.ListKind)));
            Equal("已有列表|第一段||第二段", Texts(api.Page)); True(AgentCommitter.Find(api.Page, "a") == null && AgentCommitter.Find(api.Page, "b") != null);
        });
        Test("switches remove structure tools and prompts; step descriptions", () =>
        {
            var names = new[] { "remove_blank_lines", "set_indent", "move_blocks", "insert_blocks", "text_to_table" };
            var page = Page(Paragraph("a", "正文"), Paragraph("e", ""));
            var off = Tools(new AgentPageSnapshot(page.ToString(), null, new AgentOptions { EnableBlankLineRemoval = false, EnableIndent = false, EnableMoves = false, EnableInsert = false, EnableTextTables = false }));
            var on = Tools(Snapshot(page));
            foreach (var name in names)
            {
                True(!Json(off.Definitions).Contains(name)); True(!AgentRunner.SystemPrompt(off).Contains(name));
                True(Json(on.Definitions).Contains(name)); True(AgentRunner.SystemPrompt(on).Contains(name));
            }
            var c = AiConfigStore.Parse(XElement.Parse("<AiConfig><Agent><EnableMoves>false</EnableMoves></Agent></AiConfig>"));
            True(!c.Agent.EnableMoves); True(c.Agent.EnableInsert && c.Agent.EnableIndent && c.Agent.EnableBlankLineRemoval && c.Agent.EnableTextTables);
            Equal(("删除空行 · 3 行", AgentStepState.Done), AgentTools.DescribeStep("remove_blank_lines", "{\"mode\":\"collapse\"}", "{\"ok\":true,\"removed\":[\"p1\",\"p2\",\"p3\"]}"));
            Equal(("调整缩进 · 增加缩进 · 2 段", AgentStepState.Done), AgentTools.DescribeStep("set_indent", "{\"block_ids\":[\"a\",\"b\"],\"direction\":\"in\"}", "{\"ok\":true}"));
            Equal(("移动段落 · 3 段", AgentStepState.Done), AgentTools.DescribeStep("move_blocks", "{\"block_ids\":[\"a\",\"b\",\"c\"]}", "{\"ok\":true}"));
            Equal(("插入段落 · 2 段", AgentStepState.Done), AgentTools.DescribeStep("insert_blocks", "{\"paragraphs\":[{},{}]}", "{\"ok\":true}"));
            Equal(("转换为表格 · 4 段", AgentStepState.Done), AgentTools.DescribeStep("text_to_table", "{\"block_ids\":[\"a\",\"b\",\"c\",\"d\"]}", "{\"ok\":true}"));
        });
        Test("move_blocks moves paragraphs across text boxes in one write and undo restores both boxes", () =>
        {
            var s = Snapshot(TwoBoxes()); var t = Tools(s); Read(t, s);
            True(Json(Move(t, s, new[] { "p3" }, "p2", "after")).Contains("\"changed\":[\"p3\"]"));
            var overview = Json(Invoke(t, "get_page_overview", new { }));
            True(overview.Contains("{\"id\":\"p3\",\"container_id\":\"A\"")); True(overview.Contains("{\"container_id\":\"B\",\"structure\":true,\"blocks\":1,\"summary\":\"乙二\"}"));
            True(Json(Invoke(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p4" } })).Contains("\"container_id\":\"A\""));
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(1, api.Writes); Equal(1, r.Moved); Equal(2, r.Outlines); Equal(2, r.OutlineUndo.Count);
            Equal(r.OutlineUndo[0].Group, r.OutlineUndo[1].Group);
            Equal("甲一|甲二|乙一|乙一细节", BoxTexts(api.Page, "A")); Equal("乙二", BoxTexts(api.Page, "B"));
            // 本机实测：移到另一个文本框的段落按新对象建立，原 ID 不再出现。
            True(AgentCommitter.Find(api.Page, "b1") == null && AgentCommitter.Find(api.Page, "a1") != null);
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal("Verified", undo.Status); Equal(2, undo.Outlines);
            Equal("甲一|甲二", BoxTexts(api.Page, "A")); Equal("乙一|乙一细节|乙二", BoxTexts(api.Page, "B")); True(AgentCommitter.Find(api.Page, "b2") != null);
        });
        Test("move_blocks across text boxes rejects table cells and emptying the source box", () =>
        {
            var grid = GridPage().Element(One + "Outline").Element(One + "OEChildren").Element(One + "OE");
            var s = Snapshot(Boxes(Box("A", 100, Paragraph("a", "甲"), grid), Box("B", 300, Paragraph("b", "乙"))));
            var t = Tools(s); Read(t, s);
            Rejects("单元格", () => Move(t, s, new[] { "p2" }, "p6", "after"));
            Rejects("单元格", () => Move(t, s, new[] { "p6" }, "p2", "after"));
            Rejects("merge_outlines", () => Move(t, s, new[] { "p6" }, "p1", "after"));
            Equal(0, s.Revision); Equal(0, s.LayoutChanges.Count);
        });
        Test("merge_outlines moves a whole box with its table, image and blank line, deletes it and undo rebuilds it", () =>
        {
            var s = Snapshot(TwoBoxes(Image("img", "cb-orig"), Paragraph("e", ""), GridPage().Element(One + "Outline").Element(One + "OEChildren").Element(One + "OE")));
            var t = Tools(s); Read(t, s);
            var result = Json(Merge(t, s, "B", "p2", "after"));
            True(result.Contains("\"merged\":\"B\"")); True(result.Contains("\"into\":\"A\""));
            True(!Json(Invoke(t, "get_page_overview", new { })).Contains("\"container_id\":\"B\""));
            True(Json(Invoke(t, "get_pending_changes", new { snapshot_id = s.SnapshotId })).Contains("\"merged\":[{\"from\":\"B\",\"into\":\"A\"}]"));
            var api = new FakePage(s.Page); api.Binary["cb-orig"] = "IMAGEDATA";
            var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(1, r.Merged); Equal(2, r.Outlines); Equal(1, api.Writes); Equal(0, api.Deletes); True(r.Message.Contains("合并文本框 1 个"));
            Equal(1, api.Page.Elements(One + "Outline").Count());
            Equal("甲一|甲二|乙一|乙一细节|乙二||名称|说明|甲|乙", Texts(api.Page));
            // 跨框的图片带着数据写入，没有变成坏图。
            True(api.LastXml.Contains("IMAGEDATA")); Equal("IMAGEDATA", ImageData(api, api.Page));
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal("Verified", undo.Status); Equal(2, undo.Outlines); True(undo.Message.Contains("恢复文本框结构 2 个"));
            var outlines = api.Page.Elements(One + "Outline").ToList();
            Equal(2, outlines.Count); Equal("A", (string)outlines[0].Attribute("objectID")); True((string)outlines[1].Attribute("objectID") != "B");
            Equal("甲一|甲二", BoxTexts(api.Page, "A")); Equal("乙一|乙一细节|乙二||名称|说明|甲|乙", BoxTexts(api.Page, (string)outlines[1].Attribute("objectID")));
            Equal("300", (string)outlines[1].Element(One + "Position").Attribute("y")); Equal("IMAGEDATA", ImageData(api, api.Page));
        });
        Test("merging the upper box into the lower one: OneNote orders the rebuilt box by position and undo still verifies", () =>
        {
            var s = Snapshot(TwoBoxes()); var t = Tools(s); Read(t, s);
            Merge(t, s, "A", "p5", "after");
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal("乙一|乙一细节|乙二|甲一|甲二", Texts(api.Page)); Equal("B", (string)api.Page.Elements(One + "Outline").Single().Attribute("objectID"));
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal("Verified", undo.Status); Equal("甲一|甲二|乙一|乙一细节|乙二", Texts(api.Page));
            Equal("B", (string)api.Page.Elements(One + "Outline").Last().Attribute("objectID"));
        });
        Test("merge_outlines rejects same box, protected or missing sources, cell targets and unselected paragraphs", () =>
        {
            var grid = GridPage().Element(One + "Outline").Element(One + "OEChildren").Element(One + "OE");
            var ink = Box("C", 500, Paragraph("c", "丙"), new XElement(One + "OE", new XAttribute("objectID", "ink"), new XElement(One + "InkDrawing")));
            var p = Boxes(Box("A", 100, Paragraph("a", "甲"), grid), Box("B", 300, Paragraph("b", "乙"), Paragraph("b2", "乙二")), ink);
            p.AddFirst(new XElement(One + "Title", new XAttribute("objectID", "title"), Paragraph("title-p", "标题")));
            var s = Snapshot(p); var t = Tools(s); Read(t, s);
            // p1 标题；p2 甲；p3–p6 表格；p7 乙；p8 乙二；p9 丙
            Rejects("源文本框里", () => Merge(t, s, "B", "p7", "after"));
            Rejects("受到保护", () => Merge(t, s, "C", "p2", "after"));
            Rejects("不存在", () => Merge(t, s, "title", "p2", "after"));
            Rejects("受保护的文本框", () => Merge(t, s, "B", "p9", "after"));
            Rejects("页面标题", () => Merge(t, s, "B", "p1", "after"));
            Rejects("单元格", () => Merge(t, s, "B", "p3", "after"));
            Equal(0, s.Revision);
            // 受保护的 C 不算还能合并的文本框。
            True(Json(Merge(t, s, "B", "p2", "after")).Contains("\"mergeable_left\":0"));
            Rejects("已经合并", () => Merge(t, s, "B", "p2", "after"));
            var selected = new AgentPageSnapshot(TwoBoxes().ToString(), new HashSet<string> { "a1", "b1" }, new AgentOptions());
            var st = Tools(selected); Read(st, selected);
            Rejects("没选中", () => Merge(st, selected, "B", "p1", "after"));
            // 只选中了一部分段落的 C 不能整体合并，也不计入。
            var partial = new AgentPageSnapshot(Boxes(Box("A", 100, Paragraph("a", "甲")), Box("B", 300, Paragraph("b", "乙")),
                Box("C", 500, Paragraph("c", "丙"), Paragraph("c2", "丙二"))).ToString(), new HashSet<string> { "a", "b", "c" }, new AgentOptions());
            var pt = Tools(partial); Read(pt, partial);
            True(Json(Merge(pt, partial, "B", "p1", "after")).Contains("\"mergeable_left\":0"));
        });
        Test("changes follow paragraphs into the merged box: text tables, table looks, styles and list removal", () =>
        {
            var grid = GridPage().Element(One + "Outline").Element(One + "OEChildren").Element(One + "OE");
            var s = Snapshot(Boxes(Box("A", 100, Paragraph("a", "甲")), Box("B", 300, Paragraph("r1", "名称\t说明"), Paragraph("r2", "甲\t乙"), Listed("l", "列表项", "2"), grid)));
            var t = Tools(s); Read(t, s);
            // p1 甲；p2、p3 两行文字；p4 列表项；p5–p8 表格
            Table(t, s, "tab", "p2", "p3");
            Invoke(t, "set_list", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p4" }, list = "none" });
            Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p4" }, preset_id = "quote" });
            Invoke(t, "set_table_style", new { snapshot_id = s.SnapshotId, table_ids = new[] { "t1" }, style = new { borders = false } });
            Merge(t, s, "B", "p1", "after");
            var api = new FakePage(s.Page); var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(1, r.TextTables); Equal(1, r.Tables); Equal(1, r.Applied); Equal(1, r.Merged);
            var tables = api.Page.Descendants(One + "Table").ToList();
            Equal(2, tables.Count); True(tables.All(x => x.Ancestors(One + "Outline").Single().Attribute("objectID").Value == "A"));
            True(!TableLook.Flag(tables[1], "bordersVisible")); Equal("none", AgentMarks.ListKind(api.Page.Descendants(One + "OE").First(e => AgentCode.PlainText(e) == "列表项")));
        });
        Test("remaining merges exclude unselected images and carry selected images with or without selected text", () =>
        {
            foreach (var withText in new[] { false, true }) foreach (var selectedImage in new[] { false, true })
            {
                var image = Image("img", "cb");
                if (selectedImage) image.Element(One + "Image").SetAttributeValue("selected", "all");
                var source = Box("C", 500, withText ? new[] { Paragraph("c", "丙"), image } : new[] { image });
                var selection = new HashSet<string> { "a", "b" }; if (withText) selection.Add("c");
                var s = new AgentPageSnapshot(Boxes(Box("A", 100, Paragraph("a", "甲")), Box("B", 300, Paragraph("b", "乙")), source).ToString(), selection, new AgentOptions());
                var api = new FakePage(s.Page); api.Binary["cb"] = "IMAGEDATA";
                var c = new AgentCommitter(api); var t = new AgentTools(s, c, CancellationToken.None); Read(t, s);
                var originalPage = s.Page.ToString(); var originalSource = s.Layout.Elements(One + "Outline").Single(o => (string)o.Attribute("objectID") == "C").ToString();
                RemainingMerges(Merge(t, s, "B", "p1", "after"), selectedImage ? "C" : "");
                Equal(1, s.Revision); Equal(1, s.LayoutChanges.Count); Equal(0, s.Inserted.Count); Equal(0, api.Writes);
                Equal(originalPage, s.Page.ToString()); Equal(originalSource, s.Layout.Elements(One + "Outline").Single(o => (string)o.Attribute("objectID") == "C").ToString());
                if (!selectedImage)
                {
                    var layout = s.Layout.ToString(); Rejects("选区", () => Merge(t, s, "C", "p1", "after"));
                    Equal(layout, s.Layout.ToString()); Equal(1, s.Revision); Equal(1, s.LayoutChanges.Count); Equal(0, api.Writes);
                }
                else
                {
                    RemainingMerges(Merge(t, s, "C", "p1", "after"), "");
                    var report = c.Commit(s, CancellationToken.None);
                    Equal("Verified", report.Status); Equal(2, report.Merged); Equal(1, api.Writes); Equal(3, report.OutlineUndo.Count);
                    Equal("IMAGEDATA", ImageData(api, api.Page)); Equal("Verified", c.Undo(s.PageId, report, s.Options, CancellationToken.None).Status);
                    Equal(3, api.Page.Elements(One + "Outline").Count()); Equal("IMAGEDATA", ImageData(api, api.Page));
                }
            }
        });
        Test("remaining merges exclude incompatible inherited styles and honor pending compatible formatting", () =>
        {
            foreach (var compatible in new[] { false, true })
            {
                var source = Box("C", 500, Paragraph("c", "丙")); source.SetAttributeValue("style", "font-size:30pt");
                var s = Snapshot(Boxes(Box("A", 100, Paragraph("a", "甲")), Box("B", 300, Paragraph("b", "乙")), source));
                var t = Tools(s); Read(t, s);
                if (compatible) Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p3" }, preset_id = "body" });
                var revision = s.Revision; var format = s.Blocks.Single(b => b.Id == "p3").Draft.ToString(); var styles = s.DraftStyles.ToString();
                RemainingMerges(Merge(t, s, "B", "p1", "after"), compatible ? "C" : "");
                Equal(revision + 1, s.Revision); Equal(1, s.LayoutChanges.Count); Equal(format, s.Blocks.Single(b => b.Id == "p3").Draft.ToString()); Equal(styles, s.DraftStyles.ToString());
                if (compatible)
                {
                    RemainingMerges(Merge(t, s, "C", "p1", "after"), "");
                    var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
                    Equal("Verified", r.Status); Equal(2, r.Merged); Equal(1, r.Applied); Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status);
                    Equal("甲|乙|丙", Texts(api.Page));
                }
                else
                {
                    var layout = s.Layout.ToString(); Rejects("格式", () => Merge(t, s, "C", "p1", "after"));
                    Equal(layout, s.Layout.ToString()); Equal(revision + 1, s.Revision); Equal(1, s.LayoutChanges.Count);
                }
            }
        });
        Test("remaining merges search other positions when the current nested target would change formatting", () =>
        {
            var parent = Paragraph("a", "标题"); parent.SetAttributeValue("style", "font-size:30pt");
            parent.Add(new XElement(One + "OEChildren", Paragraph("child", "子段")));
            var b = Paragraph("b", "乙"); b.SetAttributeValue("style", "font-size:30pt");
            var s = Snapshot(Boxes(Box("A", 100, parent), Box("B", 300, b), Box("C", 500, Paragraph("c", "丙"))));
            var t = Tools(s); Read(t, s);
            RemainingMerges(Merge(t, s, "B", "p2", "after"), "C");
            var layout = s.Layout.ToString(); Rejects("格式", () => Merge(t, s, "C", "p2", "after"));
            Equal(layout, s.Layout.ToString()); Equal(1, s.Revision); Equal(1, s.LayoutChanges.Count);
            RemainingMerges(Merge(t, s, "C", "p1", "before"), "");
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal("丙|标题|子段|乙", Texts(api.Page));
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status); Equal("标题|子段|乙|丙", Texts(api.Page));
            // 根级和第一个嵌套位置都不兼容，仍须继续找到另一个父段下的合法位置。
            var heading = Paragraph("h", "标题"); heading.Add(new XElement(One + "OEChildren", Paragraph("hc", "标题下")));
            var body = Paragraph("body", "正文"); body.SetAttributeValue("style", "font-size:11pt"); body.Add(new XElement(One + "OEChildren", Paragraph("bc", "正文下")));
            var destination = Box("A", 100, heading, body); destination.SetAttributeValue("style", "font-size:30pt");
            var source = Box("C", 500, Paragraph("c", "丙")); source.SetAttributeValue("style", "font-size:11pt");
            var nested = Snapshot(Boxes(destination, Box("B", 300, new XElement(b)), source)); var nt = Tools(nested); Read(nt, nested);
            RemainingMerges(Merge(nt, nested, "B", "p2", "after"), "C");
            Rejects("格式", () => Merge(nt, nested, "C", "p1", "before")); Rejects("格式", () => Merge(nt, nested, "C", "p2", "after"));
            RemainingMerges(Merge(nt, nested, "C", "p4", "before"), ""); Equal(2, nested.Revision); Equal(2, nested.LayoutChanges.Count);
            var nestedApi = new FakePage(nested.Page); var nc = new AgentCommitter(nestedApi); var nr = nc.Commit(nested, CancellationToken.None);
            Equal("Verified", nr.Status); Equal("Verified", nc.Undo(nested.PageId, nr, nested.Options, CancellationToken.None).Status);
        });
        Test("remaining merges carry a fully selected table and skip converted paragraphs as target positions", () =>
        {
            var grid = GridPage().Element(One + "Outline").Element(One + "OEChildren").Element(One + "OE");
            var s = new AgentPageSnapshot(Boxes(Box("A", 100, Paragraph("a1", "名称\t值"), Paragraph("a2", "甲\t乙"), Paragraph("a3", "目标")),
                Box("B", 300, Paragraph("b", "正文")), Box("C", 500, grid)).ToString(), new HashSet<string> { "a1", "a2", "a3", "b", "h1", "h2", "d1", "d2" }, new AgentOptions());
            var t = Tools(s); Read(t, s); Table(t, s, "tab", "p1", "p2");
            var conversion = s.CodeConversions.Single(); var styles = s.DraftStyles.ToString();
            RemainingMerges(Merge(t, s, "B", "p3", "after"), "C");
            Equal(1, s.LayoutChanges.Count); Equal(1, s.CodeConversions.Count); True(ReferenceEquals(conversion, s.CodeConversions.Single())); Equal(styles, s.DraftStyles.ToString());
            RemainingMerges(Merge(t, s, "C", "p3", "after"), "");
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(2, r.Merged); Equal(1, r.TextTables); Equal(2, api.Page.Descendants(One + "Table").Count());
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status); Equal(3, api.Page.Elements(One + "Outline").Count());
        });
        Test("remaining merge preflight handles a thousand paragraphs and leaves a cancelled merge unpublished", () =>
        {
            XElement LargePage() => Boxes(Enumerable.Range(0, 20).Select(box => Box("box" + box, box * 100,
                Enumerable.Range(0, 50).Select(i => Paragraph("line" + (box * 50 + i), "合成正文 " + i)).ToArray())).ToArray());
            var s = Snapshot(LargePage()); var t = Tools(s);
            var watch = System.Diagnostics.Stopwatch.StartNew();
            RemainingMerges(Merge(t, s, "box1", "p1", "after"), string.Join(",", Enumerable.Range(2, 18).Select(i => "box" + i)));
            watch.Stop(); Console.WriteLine($"Merge preflight: 1000 paragraphs, 20 boxes, {watch.ElapsedMilliseconds} ms");
            True(watch.Elapsed < TimeSpan.FromSeconds(10)); Equal(1, s.Revision); Equal(1, s.LayoutChanges.Count); Equal(19, s.Layout.Elements(One + "Outline").Count());
            var cancelled = Snapshot(LargePage()); var api = new FakePage(cancelled.Page); var layout = cancelled.Layout.ToString();
            using (var cancel = new CancellationTokenSource())
            {
                var ct = new AgentTools(cancelled, new AgentCommitter(api), cancel.Token);
                cancel.CancelAfter(1); watch.Restart();
                Throws<OperationCanceledException>(() => Merge(ct, cancelled, "box1", "p1", "after"));
                watch.Stop(); True(watch.Elapsed < TimeSpan.FromSeconds(5));
                Equal(layout, cancelled.Layout.ToString()); Equal(0, cancelled.Revision); Equal(0, cancelled.LayoutChanges.Count); Equal(0, api.Writes);
            }
        });
        Test("remaining merge preflight rejects incompatible sources without repeating a thousand sibling positions", () =>
        {
            var source = Box("C", 500, Paragraph("c", "不兼容")); source.SetAttributeValue("style", "font-size:30pt");
            var s = Snapshot(Boxes(Box("A", 100, Enumerable.Range(0, 997).Select(i => Paragraph("a" + i, "合成正文 " + i)).ToArray()),
                Box("B", 300, Paragraph("b", "已合并")), source, Box("D", 700, Paragraph("d", "兼容"))));
            var t = Tools(s); var watch = System.Diagnostics.Stopwatch.StartNew();
            RemainingMerges(Merge(t, s, "B", "p1", "after"), "D");
            watch.Stop(); Console.WriteLine($"Merge preflight: 1000 paragraphs, incompatible source, {watch.ElapsedMilliseconds} ms");
            True(watch.Elapsed < TimeSpan.FromSeconds(5)); Equal(1, s.Revision); Equal(1, s.LayoutChanges.Count);
            Rejects("格式", () => Merge(t, s, "C", "p1", "after"));
            RemainingMerges(Merge(t, s, "D", "p1", "after"), ""); Equal(2, s.Revision); Equal(2, s.LayoutChanges.Count);
        });
        Test("remaining merge preflight reuses format rejection across 300 nested parents", () =>
        {
            var s = Snapshot(NestedMergePage(300)); var t = Tools(s);
            var watch = System.Diagnostics.Stopwatch.StartNew();
            RemainingMerges(Merge(t, s, "C", "p1", "after"), "");
            watch.Stop(); Console.WriteLine($"Merge preflight: 602 paragraphs, 300 nested parents, {watch.ElapsedMilliseconds} ms");
            True(watch.Elapsed < TimeSpan.FromSeconds(5)); Equal(1, s.Revision); Equal(1, s.LayoutChanges.Count);
            // C 是与目标兼容的本次合并；剩余 B 的 30pt 继承样式在所有位置都不兼容。
            Rejects("格式", () => Merge(t, s, "B", "p1", "after"));
            var cancelled = Snapshot(NestedMergePage(300)); var api = new FakePage(cancelled.Page);
            var prepare = Tools(cancelled);
            Invoke(prepare, "read_blocks", new { snapshot_id = cancelled.SnapshotId, block_ids = new[] { "p1" } });
            Invoke(prepare, "set_text_style", new { snapshot_id = cancelled.SnapshotId,
                targets = new[] { new { block_id = "p1", quote = "父段", occurrence = 1, style = new { bold = true } } } });
            var state = MergeDraftState(cancelled);
            using (var cancel = new CancellationTokenSource())
            {
                var ct = new AgentTools(cancelled, new AgentCommitter(api), cancel.Token);
                cancel.CancelAfter(1);
                Throws<OperationCanceledException>(() => Merge(ct, cancelled, "C", "p1", "after"));
                Equal(state, MergeDraftState(cancelled)); Equal(0, api.Writes);
            }
        });
        Test("remaining merge format cache distinguishes QuickStyles, inline overrides and source boxes", () =>
        {
            XElement Parent(string id, string style = null)
            {
                var p = Paragraph(id, "父段"); p.SetAttributeValue("style", style);
                p.Add(new XElement(One + "OEChildren", Paragraph(id + "c", "子段"))); return p;
            }
            var same = Parent("same", "font-size:30.0pt");
            var legal = Parent("legal"); legal.SetAttributeValue("quickStyleIndex", "2");
            var destination = Box("A", 100, Parent("first"), same, legal); destination.SetAttributeValue("quickStyleIndex", "1");
            var source = Box("C", 500, Paragraph("c", "丙")); source.SetAttributeValue("style", "font-size:11pt");
            var incompatible = Box("D", 700, Paragraph("d", "丁")); incompatible.SetAttributeValue("style", "font-size:14pt");
            var inline = Box("E", 900, Paragraph("e", "<b><span style='font-size:11pt'>行内</span></b>")); inline.SetAttributeValue("style", "font-size:14pt");
            var b = Paragraph("b", "乙"); b.SetAttributeValue("style", "font-size:30pt");
            var page = Boxes(destination, Box("B", 300, b), source, incompatible, inline);
            page.AddFirst(new XElement(One + "QuickStyleDef", new XAttribute("index", "1"), new XAttribute("name", "h1"), new XAttribute("fontSize", "30")),
                new XElement(One + "QuickStyleDef", new XAttribute("index", "2"), new XAttribute("name", "p"), new XAttribute("fontSize", "11")));
            var s = Snapshot(page); var t = Tools(s); Read(t, s);
            RemainingMerges(Merge(t, s, "B", "p1", "after"), "C,E");
            Equal("C,E", FullRemainingMerges(t, s, "A"));
            var state = MergeDraftState(s); Rejects("格式", () => Merge(t, s, "C", "p2", "before")); Equal(state, MergeDraftState(s));
            RemainingMerges(Merge(t, s, "C", s.Blocks.Single(x => x.ObjectId == "legalc").Id, "after"), "E");
            RemainingMerges(Merge(t, s, "E", "p1", "before"), "");
            var api = new FakePage(s.Page); var committer = new AgentCommitter(api); var report = committer.Commit(s, CancellationToken.None);
            Equal("Verified", report.Status); Equal(3, report.Merged); Equal(1, api.Writes); Equal(4, report.OutlineUndo.Count);
            Equal("Verified", committer.Undo(s.PageId, report, s.Options, CancellationToken.None).Status);
            Equal(5, api.Page.Elements(One + "Outline").Count()); Equal(Texts(page), Texts(api.Page));
        });
        Test("remaining merge cache uses pending parent formatting and newly added style definitions", () =>
        {
            var parent = Paragraph("legal", "正文父段");
            var child = Paragraph("child", "子段"); child.SetAttributeValue("style", "font-family:Microsoft YaHei;font-size:11pt;color:automatic");
            parent.Add(new XElement(One + "OEChildren", child));
            var first = Paragraph("first", "第一个父段"); first.Add(new XElement(One + "OEChildren", Paragraph("firstc", "子段")));
            var destination = Box("A", 100, first, parent); destination.SetAttributeValue("style", "font-size:30pt");
            var source = Box("C", 500, Paragraph("c", "丙")); source.SetAttributeValue("style", "font-family:Microsoft YaHei;font-size:11pt;color:automatic");
            var b = Paragraph("b", "乙"); b.SetAttributeValue("style", "font-size:30pt");
            var d = Paragraph("d", "丁"); d.SetAttributeValue("style", "font-size:30pt");
            var page = Boxes(destination, Box("B", 300, b), source, Box("D", 700, d));
            var s = Snapshot(page); var t = Tools(s); Read(t, s);
            RemainingMerges(Merge(t, s, "B", "p1", "after"), "D");
            var parentBlock = s.Blocks.Single(x => x.ObjectId == "legal");
            Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { parentBlock.Id }, preset_id = "body" });
            True(!parentBlock.AppearanceOnly); True(s.DraftStyles.Elements().Any());
            True(AgentLayout.Find(s.Layout, parentBlock.Id).Attribute("quickStyleIndex") == null);
            var format = parentBlock.Draft.ToString(); var styles = s.DraftStyles.ToString();
            RemainingMerges(Merge(t, s, "D", "p1", "after"), "C");
            Equal("C", FullRemainingMerges(t, s, "A")); Equal(format, parentBlock.Draft.ToString()); Equal(styles, s.DraftStyles.ToString());
            RemainingMerges(Merge(t, s, "C", s.Blocks.Single(x => x.ObjectId == "child").Id, "after"), "");
            var api = new FakePage(s.Page); var committer = new AgentCommitter(api); var report = committer.Commit(s, CancellationToken.None);
            Equal("Verified", report.Status); Equal(3, report.Merged); Equal(1, report.Applied);
            Equal("Verified", committer.Undo(s.PageId, report, s.Options, CancellationToken.None).Status); Equal(Texts(page), Texts(api.Page));
        });
        Test("cached remaining merges agree with exhaustive checks for selection, images, tables and conversions", () =>
        {
            foreach (var selected in new[] { false, true }) foreach (var selectedImage in new[] { false, true })
            foreach (var selectedTable in new[] { false, true }) foreach (var conversion in new[] { "none", "code", "table" })
            {
                XElement Parent(string id, string style = null)
                {
                    var p = Paragraph(id, "父段"); p.SetAttributeValue("style", style);
                    p.Add(new XElement(One + "OEChildren", Paragraph(id + "c", "子段"))); return p;
                }
                var destination = Box("A", 100, Parent("a"), Parent("a2"), Parent("legal", "font-size:11pt"),
                    Paragraph("line1", "x\t1"), Paragraph("line2", "y\t2")); destination.SetAttributeValue("style", "font-size:30pt");
                var b = Paragraph("b", "乙"); b.SetAttributeValue("style", "font-size:30pt");
                var image = Image("img", "cb"); if (selectedImage) image.Element(One + "Image").SetAttributeValue("selected", "all");
                var c = Box("C", 500, Paragraph("c", "丙"), image); c.SetAttributeValue("style", "font-size:11pt");
                var d = Box("D", 700, GridPage().Element(One + "Outline").Element(One + "OEChildren").Element(One + "OE")); d.SetAttributeValue("style", "font-size:11pt");
                var e = Box("E", 900, Paragraph("e", "不兼容")); e.SetAttributeValue("style", "font-size:14pt");
                var f = Box("F", 1100, Paragraph("f", "兼容")); f.SetAttributeValue("style", "font-size:30pt");
                var page = Boxes(destination, Box("B", 300, b), c, d, e, f);
                var selection = selected ? new HashSet<string>(page.Descendants(One + "OE").Where(x => x.Elements(One + "T").Any()).Select(x => (string)x.Attribute("objectID"))) : null;
                if (selected && !selectedTable) selection.Remove("h2");
                var s = new AgentPageSnapshot(page.ToString(), selection, new AgentOptions()); var t = Tools(s); Read(t, s);
                var lines = new[] { "line1", "line2" }.Select(id => s.Blocks.Single(x => x.ObjectId == id).Id).ToArray();
                if (conversion == "code") Code(t, s, "python", lines);
                if (conversion == "table") Table(t, s, "tab", lines);
                var result = Merge(t, s, "B", "p1", "after");
                var state = MergeDraftState(s);
                var expected = (!selected || selectedImage ? "C," : "") + (!selected || selectedTable ? "D," : "") + "F";
                Equal(expected, FullRemainingMerges(t, s, "A")); RemainingMerges(result, expected); Equal(state, MergeDraftState(s));
            }
        });
        Test("layout diagnostics report only format failures and unexpected preflight errors leave drafts unpublished", () =>
        {
            var s = Snapshot(NestedMergePage(2)); var candidate = new XElement(s.Layout); string mismatch = null;
            AgentLayout.Merge(candidate.Elements(One + "Outline").Single(x => (string)x.Attribute("objectID") == "B"), AgentLayout.Find(candidate, "p1"), true);
            var state = MergeDraftState(s);
            Rejects("格式", () => s.CheckLayout(candidate, id => mismatch = id)); Equal(s.Blocks.Single(x => x.ObjectId == "b").Id, mismatch);
            Equal(state, MergeDraftState(s));
            var page = Boxes(Box("A", 100, Paragraph("a", "甲")), Box("B", 300, Image("img", "cb")));
            var selected = new AgentPageSnapshot(page.ToString(), new HashSet<string> { "a" }, new AgentOptions());
            candidate = new XElement(selected.Layout); mismatch = null;
            AgentLayout.Merge(candidate.Elements(One + "Outline").Last(), AgentLayout.Find(candidate, "p1"), true);
            Rejects("选区", () => selected.CheckLayout(candidate, id => mismatch = id)); Equal(null, mismatch);
            var converted = Snapshot(Boxes(Box("A", 100, Paragraph("a", "x\t1"), Paragraph("a2", "y\t2")), Box("B", 300, Paragraph("b", "乙"))));
            var ct = Tools(converted); Read(ct, converted); Table(ct, converted, "tab", "p1", "p2");
            candidate = new XElement(converted.Layout); mismatch = null;
            AgentLayout.Merge(candidate.Elements(One + "Outline").Last(), AgentLayout.Find(candidate, "p2"), true);
            Rejects("转换", () => converted.CheckLayout(candidate, id => mismatch = id)); Equal(null, mismatch);
            // 故意损坏私有短 ID，模拟准备成功后的非预期预检异常；原有草稿必须完整保留。
            var broken = Snapshot(NestedMergePage(2)); var api = new FakePage(broken.Page); var t = new AgentTools(broken, new AgentCommitter(api), CancellationToken.None);
            AgentLayout.Find(broken.Layout, broken.Blocks.Single(x => x.ObjectId == "b").Id).SetAttributeValue(AgentLayout.Key, "p1");
            state = MergeDraftState(broken);
            Throws<ArgumentException>(() => Merge(t, broken, "C", "p1", "after")); Equal(state, MergeDraftState(broken)); Equal(0, api.Writes);
        });
        Test("a user edit in either linked box skips the whole group; undo skips a group edited afterwards", () =>
        {
            var s = Snapshot(TwoBoxes()); var t = Tools(s); Read(t, s);
            Move(t, s, new[] { "p3" }, "p2", "after"); Style(t, s);
            var api = new FakePage(s.Page); AgentCommitter.Find(api.Page, "b2").Element(One + "T").Value = "用户改了乙二";
            var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
            Equal("PartiallyApplied", r.Status); Equal(0, r.Moved); Equal(1, r.Applied); True(r.ConflictIds.Contains("p3")); Equal(0, r.OutlineUndo.Count);
            Equal("甲一|甲二", BoxTexts(api.Page, "A")); Equal("乙一|乙一细节|用户改了乙二", BoxTexts(api.Page, "B"));
            var merged = Snapshot(TwoBoxes()); var mt = Tools(merged); Read(mt, merged); Merge(mt, merged, "B", "p2", "after");
            var mapi = new FakePage(merged.Page); var c = new AgentCommitter(mapi); var done = c.Commit(merged, CancellationToken.None);
            Equal("Verified", done.Status);
            AgentCommitter.Find(mapi.Page, "a1").Element(One + "T").Value = "之后改过";
            var writes = mapi.Writes;
            var undo = c.Undo(merged.PageId, done, merged.Options, CancellationToken.None);
            Equal("NoChange", undo.Status); Equal(2, undo.Conflicts); Equal(writes, mapi.Writes); Equal(1, mapi.Page.Elements(One + "Outline").Count());
        });
        Test("a merged box OneNote keeps is deleted afterwards; if that fails a blank line stays and undo still restores it", () =>
        {
            var kept = Snapshot(TwoBoxes()); var kt = Tools(kept); Read(kt, kept); Merge(kt, kept, "B", "p2", "after");
            var kapi = new FakePage(kept.Page) { KeepEmptyOutlines = true }; var kr = new AgentCommitter(kapi).Commit(kept, CancellationToken.None);
            Equal("Verified", kr.Status); Equal(1, kapi.Deletes); Equal(1, kapi.Page.Elements(One + "Outline").Count()); True(kr.OutlineUndo.Any(u => u.Deleted));
            var s = Snapshot(TwoBoxes()); var t = Tools(s); Read(t, s); Merge(t, s, "B", "p2", "after");
            var api = new FakePage(s.Page) { KeepEmptyOutlines = true, DeleteConflicts = 5 }; var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("PartiallyApplied", r.Status); Equal(1, r.Leftover); True(r.Message.Contains("没能删掉"));
            Equal("", BoxTexts(api.Page, "B")); Equal("甲一|甲二|乙一|乙一细节|乙二", BoxTexts(api.Page, "A"));
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal("Verified", undo.Status); Equal("甲一|甲二", BoxTexts(api.Page, "A")); Equal("乙一|乙一细节|乙二", BoxTexts(api.Page, "B"));
        });
        Test("merge_outlines needs two editable boxes and the move switch; step description", () =>
        {
            True(Json(Tools(Snapshot(TwoBoxes())).Definitions).Contains("merge_outlines")); True(AgentRunner.SystemPrompt(Tools(Snapshot(TwoBoxes()))).Contains("merge_outlines"));
            True(!Json(Tools(Snapshot()).Definitions).Contains("merge_outlines"));
            var off = Tools(new AgentPageSnapshot(TwoBoxes().ToString(), null, new AgentOptions { EnableMoves = false }));
            True(!Json(off.Definitions).Contains("merge_outlines")); True(!Json(off.Definitions).Contains("move_blocks")); True(!AgentRunner.SystemPrompt(off).Contains("merge_outlines"));
            Equal(("合并文本框 · 2 段", AgentStepState.Done), AgentTools.DescribeStep("merge_outlines", "{\"source_id\":\"B\"}", "{\"ok\":true,\"moved\":[\"p1\",\"p2\"]}"));
            Equal(("合并文本框 · 2 段 · 还剩 1 个", AgentStepState.Done), AgentTools.DescribeStep("merge_outlines", "{\"source_id\":\"B\"}", "{\"ok\":true,\"moved\":[\"p1\",\"p2\"],\"mergeable_left\":1}"));
            Equal(("合并文本框 · 2 段", AgentStepState.Done), AgentTools.DescribeStep("merge_outlines", "{\"source_id\":\"B\"}", "{\"ok\":true,\"moved\":[\"p1\",\"p2\"],\"mergeable_left\":0}"));
            var definitions = Json(Tools(Snapshot(TwoBoxes())).Definitions);
            True(definitions.Contains("至少一个合法位置")); True(definitions.Contains("受保护或不满足限制的框保留")); True(definitions.Contains("不要重复相同的失败调用"));
        });
        Test("three boxes merge into one with a blank line at each seam; undo rebuilds all three", () =>
        {
            var s = Snapshot(Boxes(Box("A", 100, Paragraph("a", "甲")), Box("B", 300, Paragraph("b", "乙")), Box("C", 500, Paragraph("c", "丙"))));
            var t = Tools(s); Read(t, s);
            // p1 甲；p2 乙；p3 丙
            True(Json(Merge(t, s, "B", "p1", "after")).Contains("\"mergeable_left\":1,\"mergeable_outline_ids\":[\"C\"]"));
            True(Json(Merge(t, s, "C", "p2", "after")).Contains("\"mergeable_left\":0,\"mergeable_outline_ids\":[]"));
            foreach (var target in new[] { "p2", "p3" })
                Invoke(t, "insert_blocks", new { snapshot_id = s.SnapshotId, target_id = target, position = "before", paragraphs = new object[] { new { blank = true } } });
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(2, r.Merged); Equal(2, r.InsertedBlankLines); Equal(0, r.Inserted);
            Equal(1, api.Page.Elements(One + "Outline").Count()); Equal("甲||乙||丙", BoxTexts(api.Page, "A"));
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal("Verified", undo.Status); Equal(3, api.Page.Elements(One + "Outline").Count()); Equal("甲|乙|丙", Texts(api.Page));
        });
        Test("blank line guidance and step text follow the registered tools", () =>
        {
            Equal(("插入段落 · 1 段 · 空行 1 行", AgentStepState.Done), AgentTools.DescribeStep("insert_blocks", "{\"paragraphs\":[{\"text\":\"x\"},{\"blank\":true}]}", "{\"ok\":true}"));
            Equal(("插入段落 · 空行 2 行", AgentStepState.Done), AgentTools.DescribeStep("insert_blocks", "{\"paragraphs\":[{\"blank\":true},{\"blank\":true}]}", "{\"ok\":true}"));
            True(AgentRunner.SystemPrompt(Tools(Snapshot(TwoBoxes()))).Contains(AgentRunner.MergeBlankPrompt));
            True(!AgentRunner.SystemPrompt(Tools(Snapshot())).Contains(AgentRunner.MergeBlankPrompt));
            var off = Tools(new AgentPageSnapshot(TwoBoxes().ToString(), null, new AgentOptions { EnableInsert = false }));
            True(off.Has("merge_outlines")); True(!off.Has("insert_blocks")); True(!AgentRunner.SystemPrompt(off).Contains(AgentRunner.MergeBlankPrompt));
        });
        ReviewRegressions();
        Console.WriteLine($"Agent: {_passed} passed, {_failed} failed");
        return _failed == 0 ? 0 : 1;
    }

    private static void ReviewRegressions()
    {
        Test("fix_text preserves untouched emoji and supplementary Han characters", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "😀按装说明𠮷"))); var t = Tools(s); Read(t, s);
            Fix(t, s, "p1", "按装", "安装");
            Equal("😀安装说明𠮷", s.Blocks[0].CurrentText);
            var api = new FakePage(s.Page); var c = new AgentCommitter(api);
            var done = c.Commit(s, CancellationToken.None); Equal("Verified", done.Status);
            Equal("Verified", c.Undo(s.PageId, done, s.Options, CancellationToken.None).Status);
            Equal("😀按装说明𠮷", Texts(api.Page));
        });
        Test("text_to_table preserves emoji in plain text, links and multiple runs", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "😀|", "<a href='https://example.com'><b>𠮷👍🏽</b></a>")));
            var t = Tools(s); Read(t, s); Table(t, s, "pipe", "p1");
            var api = new FakePage(s.Page); var done = new AgentCommitter(api).Commit(s, CancellationToken.None);
            Equal("Verified", done.Status); Equal("😀|𠮷👍🏽", Texts(api.Page));
            True(api.Page.ToString().Contains("https://example.com")); True(api.Page.Descendants(One + "T").Any(x => x.Value.Contains("<b>")));
        });
        Test("table conversion undo skips a later link edit", () =>
        {
            ConversionUndoConflict(false, page =>
            {
                var text = page.Descendants(One + "T").Last();
                text.Value = text.Value.Replace("/old", "/new");
            });
        });
        Test("code conversion undo skips later formatting", () =>
        {
            ConversionUndoConflict(true, page => page.Descendants(One + "Cell").First().Descendants(One + "OE").First().SetAttributeValue("alignment", "right"));
        });
        Test("table conversion undo preserves an image added afterwards", () =>
        {
            ConversionUndoConflict(false, page => page.Descendants(One + "Cell").First().Element(One + "OEChildren").Add(Image("user-image", "cb-user")));
        });
        Test("selection outdent rejects reparenting an unselected sibling atomically", () =>
        {
            var parent = Paragraph("a", "父"); parent.Add(new XElement(One + "OEChildren", Paragraph("b", "乙"), Paragraph("c", "丙")));
            var page = Page(parent);
            var s = new AgentPageSnapshot(page.ToString(), new HashSet<string> { "b" }, new AgentOptions());
            var t = Tools(s); Read(t, s); var before = s.Layout.ToString();
            Rejects("选区", () => Indent(t, s, "out", "p1"));
            Equal(before, s.Layout.ToString()); Equal(0, s.Revision); Equal(0, s.LayoutChanges.Count);
            // 把连带调整的后续兄弟段落也纳入选区后，原来的 out 语义仍然可用。
            var all = new AgentPageSnapshot(page.ToString(), new HashSet<string> { "b", "c" }, new AgentOptions());
            var tools = Tools(all); Read(tools, all); Indent(tools, all, "out", "p1");
            var api = new FakePage(all.Page); Equal("Verified", new AgentCommitter(api).Commit(all, CancellationToken.None).Status);
            Equal("b", (string)AgentCommitter.Find(api.Page, "c").Ancestors(One + "OE").First().Attribute("objectID"));
        });
        Test("parent heading falls back to appearance, preserves child and can be undone", () =>
        {
            var parent = Paragraph("a", "父标题"); parent.Add(new XElement(One + "OEChildren", Paragraph("b", "子段正文")));
            var s = Snapshot(Page(parent, Paragraph("c", "其他段"))); var t = Tools(s); Read(t, s);
            var result = Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1", "p3" }, preset_id = "heading1" });
            True(Json(result).Contains("\"appearance_only\":[\"p1\"]"));
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var child = AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "b"), api.Page);
            var done = c.Commit(s, CancellationToken.None);
            Equal("Verified", done.Status); Equal(2, done.Applied); True(done.Message.Contains("仅设置外观"));
            Equal(child, AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "b"), api.Page));
            Equal("Verified", c.Undo(s.PageId, done, s.Options, CancellationToken.None).Status);
            Equal(AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(s.Page, "a"), s.Page), AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "a"), api.Page));
        });
        Test("native heading remains available when a child has its own style", () =>
        {
            foreach (var native in new[] { true, false })
            {
                var child = Paragraph("b", "子段"); child.SetAttributeValue("quickStyleIndex", "0");
                var parent = Paragraph("a", "父标题"); parent.Add(new XElement(One + "OEChildren", child));
                var page = Page(parent); var definition = ParagraphStyles.Definition("body", new AgentOptions()); definition.SetAttributeValue("index", "0"); page.AddFirst(definition);
                var s = Snapshot(page); s.Options.EnableNativeHeadings = native; var t = Tools(s); Read(t, s); Style(t, s);
                True(!s.Blocks[0].AppearanceOnly);
                var api = new FakePage(s.Page); var c = new AgentCommitter(api); var done = c.Commit(s, CancellationToken.None);
                Equal("Verified", done.Status); Equal(0, done.AppearanceOnly.Count);
                Equal(native, AgentCommitter.Find(api.Page, "a").Attribute("quickStyleIndex") != null);
                Equal(AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(s.Page, "b"), s.Page), AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "b"), api.Page));
                Equal("Verified", c.Undo(s.PageId, done, s.Options, CancellationToken.None).Status);
            }
        });
        Test("nested heading batch isolates only the affected parent and preserves drafts on rejection", () =>
        {
            var inner = Paragraph("b", "内层标题"); inner.SetAttributeValue("quickStyleIndex", "0"); inner.Add(new XElement(One + "OEChildren", Paragraph("c", "保留的正文")));
            var outer = Paragraph("a", "外层标题"); outer.Add(new XElement(One + "OEChildren", inner));
            var page = Page(outer); var definition = ParagraphStyles.Definition("body", new AgentOptions()); definition.SetAttributeValue("index", "0"); page.AddFirst(definition);
            var s = Snapshot(page); var t = Tools(s); Read(t, s);
            var result = Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1", "p2" }, preset_id = "heading1" });
            True(Json(result).Contains("\"appearance_only\":[\"p2\"]"));
            var draft = s.CreateDraftPage().ToString(); var revision = s.Revision;
            Rejects("页面标题", () => Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1", "p2" }, preset_id = "page_title" }));
            Equal(revision, s.Revision); Equal(draft, s.CreateDraftPage().ToString()); True(s.Blocks[1].AppearanceOnly);
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var done = c.Commit(s, CancellationToken.None);
            Equal("Verified", done.Status); Equal("p2", string.Join(",", done.AppearanceOnly)); Equal(2, done.Applied);
            Equal(AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(s.Page, "c"), s.Page), AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "c"), api.Page));
            Equal("Verified", c.Undo(s.PageId, done, s.Options, CancellationToken.None).Status);
        });
        Test("removing inherited bold, italic and underline is effective and idempotent", () =>
        {
            var page = Page(Paragraph("a", "文字"));
            page.Element(One + "Outline").SetAttributeValue("style", "font-family:Arial;font-size:11pt;color:#222222;font-weight:bold;font-style:italic;text-decoration:underline");
            var s = Snapshot(page); var t = Tools(s); Read(t, s);
            var args = new { snapshot_id = s.SnapshotId, targets = new[] { new { block_id = "p1", quote = "文字", occurrence = 1, style = new { bold = false, italic = false, underline = false } } } };
            Invoke(t, "set_text_style", args); Equal(1, s.Revision);
            Invoke(t, "set_text_style", args); Equal(1, s.Revision);
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var done = c.Commit(s, CancellationToken.None);
            Equal("Verified", done.Status); Equal(1, done.Applied);
            var sig = AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "a"), api.Page);
            True(sig.Contains("font-weight:normal")); True(sig.Contains("font-style:normal")); True(sig.Contains("text-decoration:none"));
            Equal("Verified", c.Undo(s.PageId, done, s.Options, CancellationToken.None).Status);
        });
        Test("conversion undo detects marks, locked widths, whitespace and relocation", () =>
        {
            ConversionUndoConflict(false, page => page.Descendants(One + "OE").First(e => e.Elements(One + "T").Any()).AddFirst(new XElement(One + "List", new XElement(One + "Bullet", new XAttribute("bullet", "2")))));
            ConversionUndoConflict(false, page => page.Descendants(One + "Table").Single().SetAttributeValue("bordersVisible", "false"));
            ConversionUndoConflict(false, page => page.Descendants(One + "Column").First().SetAttributeValue("isLocked", "true"));
            ConversionUndoConflict(false, page => page.Descendants(One + "T").First().Value += "&nbsp;");
            ConversionUndoConflict(false, page =>
            {
                var wrapper = page.Descendants(One + "Table").Single().Parent; wrapper.Remove();
                page.Add(Box("elsewhere", 500, wrapper));
            });
        });
        Test("conversion undo ignores transient selection, wrapper ID and automatic widths", () =>
        {
            var s = CodePrepared(); var api = new FakePage(s.Page); var c = new AgentCommitter(api);
            var done = c.Commit(s, CancellationToken.None);
            var table = api.Page.Descendants(One + "Table").Single();
            table.Parent.SetAttributeValue("objectID", "regenerated-wrapper");
            foreach (var e in table.DescendantsAndSelf()) { e.SetAttributeValue("selected", "all"); e.SetAttributeValue("lastModifiedTime", "2026-09-29T00:00:00Z"); }
            foreach (var col in table.Descendants(One + "Column")) { col.SetAttributeValue("isLocked", "false"); col.SetAttributeValue("width", "888"); }
            var undo = c.Undo(s.PageId, done, s.Options, CancellationToken.None);
            Equal("Verified", undo.Status); Equal(1, undo.CodeBlocks); Equal(0, undo.Conflicts);
        });
        Test("conversion undo protects referenced definitions and ancestor styles", () =>
        {
            foreach (var definition in new[] { false, true })
            {
                var p = Page(Paragraph("a", "甲|乙"));
                p.AddFirst(new XElement(One + "QuickStyleDef", new XAttribute("index", "0"), new XAttribute("name", "p"), new XAttribute("font", "Arial"), new XAttribute("fontSize", "11")));
                p.Element(One + "Outline").SetAttributeValue("quickStyleIndex", "0");
                var s = Snapshot(p); var t = Tools(s); Read(t, s); Table(t, s, "pipe", "p1");
                var api = new FakePage(s.Page); var c = new AgentCommitter(api); var done = c.Commit(s, CancellationToken.None);
                Equal("Verified", done.Status);
                if (definition) api.Page.Element(One + "QuickStyleDef").SetAttributeValue("fontSize", "18");
                else api.Page.Element(One + "Outline").SetAttributeValue("style", "font-style:italic");
                var changed = api.Page.ToString(); var undo = c.Undo(s.PageId, done, s.Options, CancellationToken.None);
                Equal(1, undo.Conflicts); Equal(1, api.Writes); Equal(changed, api.Page.ToString());
            }
        });
        Test("several conversions undo against the same pre-write positions", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "甲|乙"), Paragraph("b", "中间正文"), Paragraph("c", "丙|丁")));
            var t = Tools(s); Read(t, s); Table(t, s, "pipe", "p1"); Table(t, s, "pipe", "p3");
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var done = c.Commit(s, CancellationToken.None);
            Equal(2, done.TextTables); var undo = c.Undo(s.PageId, done, s.Options, CancellationToken.None);
            Equal("Verified", undo.Status); Equal(2, undo.TextTables); Equal("甲|乙|中间正文|丙|丁", Texts(api.Page));
        });
        Test("conversion undo ignores paragraphs and text boxes added around it", () =>
        {
            var edits = new (Action<XElement> Edit, string Text)[]
            {
                (page => page.Element(One + "Outline").AddBeforeSelf(Box("new-box", 20, Paragraph("n1", "别处新文本框"))), "别处新文本框"),
                (page => page.Descendants(One + "Table").Single().Parent.AddBeforeSelf(Paragraph("n2", "紧挨着新打的一行")), "紧挨着新打的一行"),
            };
            foreach (var code in new[] { true, false })
                foreach (var (edit, text) in edits)
                {
                    var s = code ? CodePrepared() : Snapshot(Page(Paragraph("a", "甲|乙"), Paragraph("b", "结尾")));
                    if (!code) { var t = Tools(s); Read(t, s); Table(t, s, "pipe", "p1"); }
                    var api = new FakePage(s.Page); var c = new AgentCommitter(api); var done = c.Commit(s, CancellationToken.None);
                    Equal("Verified", done.Status); edit(api.Page);
                    var undo = c.Undo(s.PageId, done, s.Options, CancellationToken.None);
                    Equal("Verified", undo.Status); Equal(0, undo.Conflicts); Equal(code ? 1 : 0, undo.CodeBlocks); Equal(code ? 0 : 1, undo.TextTables);
                    True(!api.Page.Descendants(One + "Table").Any()); True(Texts(api.Page).Contains(text));
                }
        });
        Test("Unicode edits preserve runs, formatting, links, spaces and breaks", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "&nbsp;😀<b>按装</b>", "<a href='https://example.com'>𠮷</a><br>👍🏽尾")));
            var t = Tools(s); Read(t, s); Fix(t, s, "p1", "按装", "安装");
            Equal("\u00a0😀安装𠮷\n👍🏽尾", s.Blocks[0].CurrentText);
            var draft = s.Blocks[0].Draft;
            Equal(2, draft.Elements(One + "T").Count()); True(draft.Element(One + "T").Value.Contains("<b>安装</b>"));
            True(draft.Elements(One + "T").Last().Value.Contains("href="));
            var revision = s.Revision;
            Throws(() => Fix(t, s, "p1", "👍", "好")); Equal(revision, s.Revision);
        });
        Test("selection structural tools protect unselected descendants and images", () =>
        {
            var parent = Paragraph("a", "父"); parent.Add(new XElement(One + "OEChildren", Paragraph("b", "没选中的子段")));
            var s = new AgentPageSnapshot(Page(Paragraph("c", "目标"), parent).ToString(), new HashSet<string> { "a", "c" }, new AgentOptions());
            var t = Tools(s); Read(t, s);
            // 挂到另一段下会改变未选子段的祖先链。
            Rejects("选区", () => Indent(t, s, "in", "p2")); Equal(0, s.Revision);
            var imagePage = Boxes(Box("A", 100, Paragraph("a", "选中的目标")), Box("B", 300, Image("picture", "cb")));
            var imageSelection = new AgentPageSnapshot(imagePage.ToString(), new HashSet<string> { "a" }, new AgentOptions());
            var imageTools = Tools(imageSelection); Read(imageTools, imageSelection);
            Rejects("选区", () => Merge(imageTools, imageSelection, "B", "p1", "after")); Equal(0, imageSelection.Revision);
        });
        Test("selection permits inserting or moving selected siblings around unselected text", () =>
        {
            var s = new AgentPageSnapshot(Page(Paragraph("a", "甲"), Paragraph("b", "未选"), Paragraph("c", "丙")).ToString(), new HashSet<string> { "a", "c" }, new AgentOptions());
            var t = Tools(s); Read(t, s); Insert(t, s, "p1", "新增"); Move(t, s, new[] { "p2" }, "p1", "before");
            var api = new FakePage(s.Page); Equal("Verified", new AgentCommitter(api).Commit(s, CancellationToken.None).Status);
            Equal("丙|甲|新增|未选", Texts(api.Page));
        });
        Test("selection structural tools carry fully selected tables and images", () =>
        {
            XElement Grid() => GridPage().Element(One + "Outline").Element(One + "OEChildren").Element(One + "OE");
            var cells = new[] { "h1", "h2", "d1", "d2" };
            // 文字和图片都选中：图片跟着合并过去，带着数据写入。
            var image = Image("img", "cb"); image.Element(One + "Image").SetAttributeValue("selected", "all");
            var imageSelection = new AgentPageSnapshot(Boxes(Box("A", 100, Paragraph("a", "甲")), Box("B", 300, Paragraph("b", "乙"), image)).ToString(),
                new HashSet<string> { "a", "b" }, new AgentOptions());
            var imageTools = Tools(imageSelection); Read(imageTools, imageSelection);
            Merge(imageTools, imageSelection, "B", "p1", "after"); Equal(1, imageSelection.Revision);
            var imageApi = new FakePage(imageSelection.Page); imageApi.Binary["cb"] = "IMAGEDATA";
            Equal("Verified", new AgentCommitter(imageApi).Commit(imageSelection, CancellationToken.None).Status);
            Equal(1, imageApi.Page.Elements(One + "Outline").Count()); Equal("IMAGEDATA", ImageData(imageApi, imageApi.Page));
            // 文字和表格全部选中：整框合并，提交和撤销都通过。
            var tableSelection = new AgentPageSnapshot(Boxes(Box("A", 100, Paragraph("a", "甲")), Box("B", 300, Paragraph("b", "乙"), Grid())).ToString(),
                new HashSet<string>(cells) { "a", "b" }, new AgentOptions());
            var tableTools = Tools(tableSelection); Read(tableTools, tableSelection);
            Merge(tableTools, tableSelection, "B", "p1", "after"); Equal(1, tableSelection.Revision);
            var tableApi = new FakePage(tableSelection.Page); var committer = new AgentCommitter(tableApi);
            var merged = committer.Commit(tableSelection, CancellationToken.None);
            Equal("Verified", merged.Status); Equal("甲|乙|名称|说明|甲|乙", Texts(tableApi.Page)); Equal(1, tableApi.Page.Elements(One + "Outline").Count());
            Equal("Verified", committer.Undo(tableSelection.PageId, merged, tableSelection.Options, CancellationToken.None).Status);
            Equal(2, tableApi.Page.Elements(One + "Outline").Count());
            // 下面挂着全选表格的段落可以缩进，表格跟着走。
            var parent = Paragraph("p", "父段"); parent.Add(new XElement(One + "OEChildren", Grid()));
            var indentPage = Page(Paragraph("q", "前一段"), parent);
            var indentSelection = new AgentPageSnapshot(indentPage.ToString(), new HashSet<string>(cells) { "q", "p" }, new AgentOptions());
            var indentTools = Tools(indentSelection); Read(indentTools, indentSelection);
            Indent(indentTools, indentSelection, "in", "p2"); Equal(1, indentSelection.Revision);
            var indentApi = new FakePage(indentSelection.Page);
            Equal("Verified", new AgentCommitter(indentApi).Commit(indentSelection, CancellationToken.None).Status);
            var moved = AgentCommitter.Find(indentApi.Page, "p");
            Equal("q", (string)moved.Ancestors(One + "OE").First().Attribute("objectID")); True(moved.Descendants(One + "Table").Any());
            // 表格有一格没选中就不在选区里，仍然整次拒绝。
            var partial = new AgentPageSnapshot(indentPage.ToString(), new HashSet<string> { "q", "p", "h1", "h2", "d1" }, new AgentOptions());
            var partialTools = Tools(partial); Read(partialTools, partial); var layout = partial.Layout.ToString();
            Rejects("选区", () => Indent(partialTools, partial, "in", "p2")); Equal(0, partial.Revision); Equal(layout, partial.Layout.ToString());
        });
        Test("heading after indentation uses current children and reports verified appearance only", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "父"), Paragraph("b", "子"), Paragraph("c", "末尾"))); var t = Tools(s); Read(t, s);
            Indent(t, s, "in", "p2");
            var outcome = Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" }, preset_id = "heading2" });
            True(Json(outcome).Contains("\"appearance_only\":[\"p1\"]"));
            var revision = s.Revision; Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" }, preset_id = "heading2" }); Equal(revision, s.Revision);
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var done = c.Commit(s, CancellationToken.None);
            Equal("Verified", done.Status); Equal(1, done.AppearanceOnly.Count); Equal(1, done.OutlineUndo.Count);
            Equal(AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(s.Page, "b"), s.Page), AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "b"), api.Page));
            Equal("Verified", c.Undo(s.PageId, done, s.Options, CancellationToken.None).Status);
        });
        Test("appearance fallback does not count conflicted or unverified paragraphs", () =>
        {
            foreach (var conflict in new[] { false, true })
            {
                var parent = Paragraph("a", "父"); parent.Add(new XElement(One + "OEChildren", Paragraph("b", "子")));
                var s = Snapshot(Page(parent)); var t = Tools(s); Read(t, s); Style(t, s);
                var api = new FakePage(s.Page);
                if (conflict) AgentCommitter.Find(api.Page, "b").Element(One + "T").Value = "后来编辑";
                else api.AfterSave = () => AgentCommitter.Find(api.Page, "a").SetAttributeValue("alignment", "right");
                var report = new AgentCommitter(api).Commit(s, CancellationToken.None);
                Equal(0, report.AppearanceOnly.Count); True(!report.Message.Contains("仅设置外观"));
                True(conflict ? report.Conflicts > 0 : report.Unverified > 0);
            }
        });
        Test("appearance-only fallback matches the native preset appearance", () =>
        {
            XElement Bold() { var d = ParagraphStyles.Definition("heading1", new AgentOptions()); d.SetAttributeValue("index", "0"); return d; }
            XElement Italic() => new XElement(One + "QuickStyleDef", new XAttribute("index", "0"), new XAttribute("name", "quote"), new XAttribute("fontColor", "#595959"),
                new XAttribute("highlightColor", "automatic"), new XAttribute("font", "Calibri"), new XAttribute("fontSize", "11.0"), new XAttribute("italic", "true"),
                new XAttribute("spaceBefore", "0.0"), new XAttribute("spaceAfter", "0.0"));
            // 原有定义带来的加粗、斜体不能留在父段上；文字上显式的样式和行内格式照旧保留。
            var cases = new (Func<XElement> Definition, string Preset, string Style)[] { (Bold, "body", null), (Italic, "heading1", null), (Italic, "heading1", "font-style:italic") };
            foreach (var (definition, preset, style) in cases)
            {
                XElement Build(bool child)
                {
                    var parent = Paragraph("a", "父<u>段</u>"); parent.SetAttributeValue("quickStyleIndex", "0");
                    if (style != null) parent.Element(One + "T").SetAttributeValue("style", style);
                    if (child) parent.Add(new XElement(One + "OEChildren", Paragraph("b", "子段")));
                    var built = Page(parent); built.AddFirst(definition()); return built;
                }
                string Look(AgentPageSnapshot snapshot) { var draft = snapshot.CreateDraftPage(); return new AgentRichText(AgentLayout.Find(draft, "p1")).Signature(draft, true); }
                var control = Snapshot(Build(false)); var ct = Tools(control); Read(ct, control);
                Invoke(ct, "set_paragraph_style", new { snapshot_id = control.SnapshotId, block_ids = new[] { "p1" }, preset_id = preset });
                var s = Snapshot(Build(true)); var t = Tools(s); Read(t, s);
                var result = Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" }, preset_id = preset });
                True(Json(result).Contains("\"appearance_only\":[\"p1\"]"));
                Equal(Look(control), Look(s));
                var api = new FakePage(s.Page); var c = new AgentCommitter(api); var child = AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "b"), api.Page);
                var done = c.Commit(s, CancellationToken.None);
                Equal("Verified", done.Status); Equal("p1", string.Join(",", done.AppearanceOnly));
                Equal(Look(control), new AgentRichText(AgentCommitter.Find(api.Page, "a")).Signature(api.Page, true));
                Equal(child, AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "b"), api.Page));
                Equal("Verified", c.Undo(s.PageId, done, s.Options, CancellationToken.None).Status);
            }
        });
        Test("duplicate paragraph IDs neither abort styling nor commit", () =>
        {
            string Duplicates(XElement page) => string.Join("|", page.Descendants(One + "OE").Where(e => (string)e.Attribute("objectID") == "dup").Select(e => AgentPageSnapshot.SemanticFormat(e, page)));
            // 目标父段下有重复 ID 的子段：按位置核对继承格式，提交和撤销照常。
            var parent = Paragraph("a", "父段"); parent.Add(new XElement(One + "OEChildren", Paragraph("dup", "子一"), Paragraph("dup", "子二")));
            var s = Snapshot(Page(parent, Paragraph("c", "其他段"))); var t = Tools(s); Read(t, s);
            var result = Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1", "p4" }, preset_id = "heading1" });
            True(Json(result).Contains("\"appearance_only\":[\"p1\"]"));
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var children = Duplicates(api.Page);
            var done = c.Commit(s, CancellationToken.None);
            Equal("Verified", done.Status); Equal(2, done.Applied); Equal(children, Duplicates(api.Page));
            Equal("Verified", c.Undo(s.PageId, done, s.Options, CancellationToken.None).Status);
            // 重复 ID 的段落不在目标下，只改别的段落。
            var other = Snapshot(Page(Paragraph("dup", "甲"), Paragraph("dup", "乙"), Paragraph("c", "丙")));
            var ot = Tools(other); Read(ot, other);
            Invoke(ot, "set_paragraph_style", new { snapshot_id = other.SnapshotId, block_ids = new[] { "p3" }, preset_id = "heading2" });
            var otherApi = new FakePage(other.Page); var oc = new AgentCommitter(otherApi); var od = oc.Commit(other, CancellationToken.None);
            Equal("Verified", od.Status); Equal(1, od.Applied);
            Equal("Verified", oc.Undo(other.PageId, od, other.Options, CancellationToken.None).Status);
        });
        Test("appearance-only fallback keeps emphasis set on the paragraph itself", () =>
        {
            // 父段 style 上的加粗、斜体、下划线在原生样式下覆盖定义，回退时不能被定义的默认值盖掉；先设父段、再设子段照常完成。
            foreach (var (style, preset) in new[] { ("font-weight:bold", "body"), ("font-style:italic", "heading1"), ("text-decoration:underline", "heading2") })
            {
                XElement Build(bool child)
                {
                    var parent = Paragraph("a", "父段"); parent.SetAttributeValue("style", style);
                    if (child) parent.Add(new XElement(One + "OEChildren", Paragraph("b", "子段")));
                    return Page(parent);
                }
                string Look(AgentPageSnapshot snapshot) { var draft = snapshot.CreateDraftPage(); return new AgentRichText(AgentLayout.Find(draft, "p1")).Signature(draft, true); }
                var control = Snapshot(Build(false)); var ct = Tools(control); Read(ct, control);
                Invoke(ct, "set_paragraph_style", new { snapshot_id = control.SnapshotId, block_ids = new[] { "p1" }, preset_id = preset });
                var s = Snapshot(Build(true)); var t = Tools(s); Read(t, s);
                var result = Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" }, preset_id = preset });
                True(Json(result).Contains("\"appearance_only\":[\"p1\"]"));
                Equal(Look(control), Look(s));
                Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p2" }, preset_id = "body" });
                var api = new FakePage(s.Page); var c = new AgentCommitter(api);
                var done = c.Commit(s, CancellationToken.None);
                Equal("Verified", done.Status); Equal(2, done.Applied); Equal("p1", string.Join(",", done.AppearanceOnly));
                Equal(Look(control), new AgentRichText(AgentCommitter.Find(api.Page, "a")).Signature(api.Page, true));
                Equal("Verified", c.Undo(s.PageId, done, s.Options, CancellationToken.None).Status);
            }
        });
        Test("conversion undo ignores quick style renumbering but not definition changes", () =>
        {
            // OneNote 可能给样式定义重新编号：内容不变时代码框和转换的表格照常换回，定义内容变了仍算冲突。
            XElement Definition(string index, string name, string size) => new XElement(One + "QuickStyleDef", new XAttribute("index", index), new XAttribute("name", name),
                new XAttribute("fontColor", "automatic"), new XAttribute("highlightColor", "automatic"), new XAttribute("font", "Calibri"), new XAttribute("fontSize", size),
                new XAttribute("spaceBefore", "0.0"), new XAttribute("spaceAfter", "0.0"));
            foreach (var code in new[] { true, false })
            foreach (var redefine in new[] { false, true })
            {
                var parent = Paragraph("a", "上级段落"); parent.SetAttributeValue("quickStyleIndex", "1");
                parent.Add(new XElement(One + "OEChildren", Paragraph("c1", code ? "def f(x):" : "甲|乙")));
                var page = Page(parent); page.AddFirst(Definition("0", "h1", "16.0"), Definition("1", "p", "11.0"));
                var s = Snapshot(page); var t = Tools(s); Read(t, s);
                if (code) Code(t, s, "python", "p2"); else Table(t, s, "pipe", "p2");
                var api = new FakePage(s.Page); var c = new AgentCommitter(api); var done = c.Commit(s, CancellationToken.None);
                Equal("Verified", done.Status);
                var swap = new Dictionary<string, string> { ["0"] = "1", ["1"] = "0" };
                foreach (var d in api.Page.Elements(One + "QuickStyleDef")) d.SetAttributeValue("index", swap[(string)d.Attribute("index")]);
                foreach (var a in api.Page.Descendants().Attributes("quickStyleIndex")) a.Value = swap[a.Value];
                if (redefine) api.Page.Elements(One + "QuickStyleDef").Single(d => (string)d.Attribute("name") == "p").SetAttributeValue("fontSize", "12.0");
                var undo = c.Undo(s.PageId, done, s.Options, CancellationToken.None);
                Equal(redefine ? 1 : 0, undo.Conflicts);
                Equal(redefine, api.Page.Descendants(One + "Table").Any());
                if (!redefine)
                {
                    Equal("Verified", undo.Status);
                    True(api.Page.Descendants(One + "OE").Any(e => e.Elements(One + "T").Any() && AgentCode.PlainText(e) == (code ? "def f(x):" : "甲|乙")));
                }
            }
        });
    }

    private static void ConversionUndoConflict(bool code, Action<XElement> edit)
    {
        var s = code ? CodePrepared() : Snapshot(Page(Paragraph("a", "甲|<a href='https://example.com/old'>乙</a>")));
        if (!code) { var t = Tools(s); Read(t, s); Table(t, s, "pipe", "p1"); }
        var api = new FakePage(s.Page); var c = new AgentCommitter(api); var done = c.Commit(s, CancellationToken.None);
        Equal("Verified", done.Status); edit(api.Page); var changed = api.Page.ToString();
        var undo = c.Undo(s.PageId, done, s.Options, CancellationToken.None);
        Equal(1, undo.Conflicts); Equal(1, api.Writes); Equal(changed, api.Page.ToString());
    }

    internal static XElement Paragraph(string id, params string[] runs) => new XElement(One + "OE", new XAttribute("objectID", id), runs.Select(t => new XElement(One + "T", new XCData(t))));
    internal static XElement Page(params XElement[] paragraphs) => new XElement(One + "Page", new XAttribute("ID", "page"), new XAttribute("lastModifiedTime", "2026-09-26T00:00:00Z"),
        new XElement(One + "Outline", new XAttribute("objectID", "outline"), new XAttribute("style", "font-family:Arial;font-size:11pt;color:#222222"), new XElement(One + "OEChildren", paragraphs)));
    private static AgentPageSnapshot Snapshot(XElement page = null) => new AgentPageSnapshot((page ?? Page(Paragraph("a", "第一段"), Paragraph("b", "第二段"))).ToString(), null, new AgentOptions());
    // 849 段，其中首 100 段有 60 段正文，后续有 396 段；不保存用户文档作为测试数据。
    private static XElement LongCoveragePage() => Page(Enumerable.Range(0, 849).Select(i => Paragraph("long" + i,
        (i < 100 ? i % 5 < 3 : (i - 99) * 396 / 749 != (i - 100) * 396 / 749) ? "测试正文 " + i : "")).ToArray());
    private static object ToolField(object value, string name) => value.GetType().GetProperty(name)?.GetValue(value, null);
    private static XElement TablePage(bool locked) => Page(new XElement(One + "OE", new XAttribute("objectID", "wrapper"),
        new XElement(One + "Table", new XAttribute("objectID", "table"),
            new XElement(One + "Columns", new XElement(One + "Column", new XAttribute("index", "0"), new XAttribute("width", "200"), new XAttribute("isLocked", locked))),
            new XElement(One + "Row", new XAttribute("objectID", "row"), new XElement(One + "Cell", new XAttribute("objectID", "cell"), new XElement(One + "OEChildren", Paragraph("a", "第一段")))))));
    /// <summary>两行两列、显示边框的表格：表头「名称 | 说明」，数据「甲 | 乙」。</summary>
    private static XElement GridPage() => Page(new XElement(One + "OE", new XAttribute("objectID", "wrapper"),
        new XElement(One + "Table", new XAttribute("objectID", "table"), new XAttribute("bordersVisible", "true"),
            new XElement(One + "Columns", new XElement(One + "Column", new XAttribute("index", "0"), new XAttribute("width", "100")),
                new XElement(One + "Column", new XAttribute("index", "1"), new XAttribute("width", "100"))),
            new XElement(One + "Row", new XAttribute("objectID", "r1"), GridCell("h1", "名称"), GridCell("h2", "说明")),
            new XElement(One + "Row", new XAttribute("objectID", "r2"), GridCell("d1", "甲"), GridCell("d2", "乙")))));
    private static XElement GridCell(string id, string text) => new XElement(One + "Cell", new XAttribute("objectID", "cell-" + id), new XElement(One + "OEChildren", Paragraph(id, text)));
    private static XElement TagDef(string index, int symbol, string name) => new XElement(One + "TagDef", new XAttribute("index", index), new XAttribute("type", "0"),
        new XAttribute("symbol", symbol), new XAttribute("fontColor", "automatic"), new XAttribute("highlightColor", "none"), new XAttribute("name", name));
    private static XElement Tag(string index) => new XElement(One + "Tag", new XAttribute("index", index), new XAttribute("completed", "false"), new XAttribute("disabled", "false"));
    private static XElement Listed(string id, string text, string bullet)
    { var oe = Paragraph(id, text); oe.AddFirst(new XElement(One + "List", new XElement(One + "Bullet", new XAttribute("bullet", bullet)))); return oe; }
    private static string Json(object value) => AgentChatClient.Serializer().Serialize(value);
    /// <summary>页面上各文字段落的纯文字，按页面顺序用 | 连接。</summary>
    private static string Texts(XElement page) => string.Join("|", page.Descendants(One + "OE").Where(e => e.Elements(One + "T").Any()).Select(AgentCode.PlainText));
    private static string Ids(XElement page) => string.Join(",", page.Descendants(One + "OE").Select(e => (string)e.Attribute("objectID")));
    /// <summary>几个文本框组成的页面。</summary>
    private static XElement Boxes(params XElement[] outlines) => new XElement(One + "Page", new XAttribute("ID", "page"), new XAttribute("lastModifiedTime", "2026-09-26T00:00:00Z"), outlines);
    private static XElement Box(string id, int y, params XElement[] paragraphs) => new XElement(One + "Outline", new XAttribute("objectID", id),
        new XElement(One + "Position", new XAttribute("x", "36"), new XAttribute("y", y), new XAttribute("z", "0")),
        new XElement(One + "Size", new XAttribute("width", "400"), new XAttribute("height", "100")), new XElement(One + "OEChildren", paragraphs));
    /// <summary>上面的文本框 A：甲一、甲二（p1、p2）；下面的 B：乙一带下级乙一细节、乙二（p3–p5），再接 extra。</summary>
    private static XElement TwoBoxes(params XElement[] extra)
    {
        var b1 = Paragraph("b1", "乙一"); b1.Add(new XElement(One + "OEChildren", Paragraph("b1c", "乙一细节")));
        return Boxes(Box("A", 100, Paragraph("a1", "甲一"), Paragraph("a2", "甲二")), Box("B", 300, new[] { b1, Paragraph("b2", "乙二") }.Concat(extra).ToArray()));
    }
    private static string BoxTexts(XElement page, string outlineId) => Texts(page.Elements(One + "Outline").Single(o => (string)o.Attribute("objectID") == outlineId));
    private static XElement Image(string id, string callback) => new XElement(One + "OE", new XAttribute("objectID", id), new XElement(One + "Image",
        new XElement(One + "Size", new XAttribute("width", "16"), new XAttribute("height", "16")), new XElement(One + "CallbackID", new XAttribute("callbackID", callback))));
    /// <summary>页面上唯一一张图片的数据；坏图返回 null。</summary>
    private static string ImageData(FakePage api, XElement page) =>
        api.Binary.TryGetValue((string)page.Descendants(One + "Image").Single().Element(One + "CallbackID").Attribute("callbackID"), out var data) ? data : null;
    private static object Merge(AgentTools tools, AgentPageSnapshot s, string source, string target, string position) =>
        Invoke(tools, "merge_outlines", new { snapshot_id = s.SnapshotId, source_id = source, target_id = target, position });
    private static void RemainingMerges(object result, string expected)
    {
        var ids = (string[])ToolField(result, "mergeable_outline_ids");
        Equal(expected, string.Join(",", ids)); Equal(ids.Length, (int)ToolField(result, "mergeable_left"));
    }
    private static XElement NestedMergePage(int parents)
    {
        var source = Box("B", 300, Paragraph("b", "不兼容")); source.SetAttributeValue("style", "font-size:30pt");
        return Boxes(Box("A", 100, Enumerable.Range(0, parents).Select(i =>
        {
            var p = Paragraph("parent" + i, "父段 " + i);
            p.Add(new XElement(One + "OEChildren", Paragraph("child" + i, "子段"))); return p;
        }).ToArray()), source, Box("C", 500, Paragraph("c", "本次合并")));
    }
    /// <summary>独立于缓存逐一尝试所有段落的前后；只调用完整准备，不发布草稿。</summary>
    private static string FullRemainingMerges(AgentTools tools, AgentPageSnapshot snapshot, string into)
    {
        var prepare = typeof(AgentTools).GetMethod("PrepareMerge", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        var targets = snapshot.Layout.Elements(One + "Outline").Single(o => (string)o.Attribute("objectID") == into)
            .Descendants(One + "OE").Select(AgentLayout.KeyOf).Where(id => id != null).ToArray();
        return string.Join(",", snapshot.Layout.Elements(One + "Outline").Select(o => (string)o.Attribute("objectID"))
            .Where(id => id != null && id != into && CanMerge(id)));
        bool CanMerge(string source)
        {
            foreach (var target in targets) foreach (var before in new[] { true, false })
            {
                try { prepare.Invoke(tools, new object[] { snapshot.Layout, source, target, before, null }); return true; }
                catch (System.Reflection.TargetInvocationException ex) when (ex.InnerException is AiException) { }
            }
            return false;
        }
    }
    private static string MergeDraftState(AgentPageSnapshot snapshot) => Json(new
    {
        page = snapshot.Page.ToString(), layout = snapshot.Layout.ToString(), revision = snapshot.Revision,
        styles = snapshot.DraftStyles.ToString(), tags = snapshot.DraftTags.ToString(), frozen = snapshot.Frozen, spacing = snapshot.CodeSpacingRequested,
        blocks = snapshot.Blocks.Select(b => new { id = b.Id, draft = b.Draft.ToString(), fixes = b.TextFixes.ToArray(), marks = b.MarkdownMarks,
            appearance = b.AppearanceOnly, read = b.Read }).ToArray(),
        inserted = snapshot.Inserted.Select(i => new { id = i.Id, outline = i.OutlineId, text = i.Text }).ToArray(),
        changes = snapshot.LayoutChanges.Select(c => new { kind = c.Kind, outline = c.OutlineId, from = c.From, ids = c.Ids }).ToArray(),
        conversions = snapshot.CodeConversions.Select(c => new { ids = c.Blocks.Select(b => b.Id).ToArray(), code = c.Code, table = c.Table.ToString(),
            textTable = c.TextTable }).ToArray()
    });
    private static object Move(AgentTools tools, AgentPageSnapshot s, string[] ids, string target, string position) =>
        Invoke(tools, "move_blocks", new { snapshot_id = s.SnapshotId, block_ids = ids, target_id = target, position });
    private static object Indent(AgentTools tools, AgentPageSnapshot s, string direction, params string[] ids) =>
        Invoke(tools, "set_indent", new { snapshot_id = s.SnapshotId, block_ids = ids, direction });
    private static object Insert(AgentTools tools, AgentPageSnapshot s, string target, string text) =>
        Invoke(tools, "insert_blocks", new { snapshot_id = s.SnapshotId, target_id = target, position = "after", paragraphs = new[] { new { text, preset_id = "body" } } });
    private static object InsertBlanks(AgentTools tools, AgentPageSnapshot s, string target, int count = 1, string position = "after") =>
        Invoke(tools, "insert_blocks", new { snapshot_id = s.SnapshotId, target_id = target, position,
            paragraphs = Enumerable.Range(0, count).Select(_ => new { blank = true }).ToArray() });
    private static object RemoveBlanks(AgentTools tools, AgentPageSnapshot s, string mode = "all") =>
        Invoke(tools, "remove_blank_lines", new { snapshot_id = s.SnapshotId, mode });
    private static object Table(AgentTools tools, AgentPageSnapshot s, string delimiter, params string[] ids) =>
        Invoke(tools, "text_to_table", new { snapshot_id = s.SnapshotId, block_ids = ids, delimiter, header_shading = "#DEEAF6" });
    private static AgentTools Tools(AgentPageSnapshot s) => new AgentTools(s, new AgentCommitter(new FakePage(s.Page)), CancellationToken.None);
    private static object Invoke(AgentTools tools, string name, object args) => tools.Execute(new AgentToolCall { Id = "test", Name = name, Arguments = AgentChatClient.Serializer().Serialize(args) });
    private static void Read(AgentTools tools, AgentPageSnapshot s) => Invoke(tools, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = s.Blocks.Where(b => b.Editable).Select(b => b.Id).ToArray() });
    private static void Style(AgentTools tools, AgentPageSnapshot s) => Invoke(tools, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" }, preset_id = "heading1" });
    private static AgentPageSnapshot Prepared(XElement page = null) { var s = Snapshot(page); var t = Tools(s); Read(t, s); Style(t, s); return s; }
    /// <summary>把几段文字当作一个文本框里连续的段落去掉 Markdown 标记，结果用 | 连接。</summary>
    private static string Strip(params string[] paragraphs)
    {
        var fences = new AgentMarkdown.Fences(); var kinds = new HashSet<string>(AgentMarkdown.Kinds);
        return string.Join("|", paragraphs.Select(p => AgentMarkdown.Remove(p, AgentMarkdown.Analyze(p, p.Split('\n').Select(fences.Next).ToList(), kinds, false).Marks)));
    }
    private static object Fix(AgentTools tools, AgentPageSnapshot s, string id, string quote, string replacement) =>
        Invoke(tools, "fix_text", new { snapshot_id = s.SnapshotId, fixes = new[] { new { block_id = id, quote, occurrence = 1, replacement } } });
    private static object Cleanup(AgentTools tools, AgentPageSnapshot s, params string[] ids) =>
        Invoke(tools, "strip_markdown", new { snapshot_id = s.SnapshotId, block_ids = ids });

    private static object NumberList(AgentTools tools, AgentPageSnapshot s, params string[] ids) =>
        Invoke(tools, "set_list", new { snapshot_id = s.SnapshotId, block_ids = ids, list = "number" });

    private static string HttpFailure(int status, string body, string media = "application/json")
    {
        var handler = new StubHttp(body, media) { Status = (HttpStatusCode)status };
        var config = AiConfigStore.Parse(XElement.Parse("<AiConfig><ApiUrl>https://test.invalid</ApiUrl><ApiKey>test</ApiKey></AiConfig>"));
        using (var http = new HttpClient(handler))
        {
            try { new AgentChatClient(config, "test", "none", http).CompleteAsync(new List<object>(), new object[0], null, CancellationToken.None).GetAwaiter().GetResult(); }
            catch (AiException ex) { return ex.Message; }
        }
        throw new Exception("Expected HTTP failure");
    }

    /// <summary>模拟 OneNote 跨普通正文继续编号的已复现行为；只有显式起点重启，同一父节点内递增。</summary>
    private static void RenderNumbering(XElement page)
    {
        foreach (var children in page.Descendants(One + "OEChildren"))
        {
            var counter = 0;
            foreach (var oe in children.Elements(One + "OE"))
            {
                var number = oe.Element(One + "List")?.Element(One + "Number");
                if (number == null) continue;
                counter = (int?)number.Attribute("restartNumberingAt") ?? counter + 1;
                number.SetAttributeValue("text", ((string)number.Attribute("numberFormat")).Replace("##", counter.ToString(CultureInfo.InvariantCulture)));
                number.SetAttributeValue("fontSize", "11.0");
            }
        }
    }
    /// <summary>正文、三行代码（中间一个空行）、正文。代码段落 p2–p4。</summary>
    private static XElement CodePage()
    {
        var indented = Paragraph("c3", "&nbsp;&nbsp;&nbsp;&nbsp;return x + 1"); indented.SetAttributeValue("style", "font-size:9pt");
        var first = Paragraph("c1", "def f(x):"); first.SetAttributeValue("style", "font-size:9pt");
        return Page(Paragraph("a", "示例："), first, Paragraph("c2", ""), indented, Paragraph("b", "结尾"));
    }
    private static object Code(AgentTools tools, AgentPageSnapshot s, string language, params string[] ids) =>
        Invoke(tools, "highlight_code", new { snapshot_id = s.SnapshotId, block_ids = ids, language });
    private static object Normalize(AgentTools tools, AgentPageSnapshot s) => AgentChatClient.Parse(Json(Invoke(tools, "normalize_code_spacing", new { snapshot_id = s.SnapshotId })));
    private static XElement SpacingBox(string id)
    {
        var lines = new[] { "x = 1", "&nbsp;", "y = 2" }.Select((text, i) => new XElement(One + "OE",
            new XAttribute("objectID", id + (i + 1)), new XAttribute("style", "font-family:Consolas;font-size:10pt"), new XElement(One + "T", new XCData(text))));
        return new XElement(One + "OE", new XAttribute("objectID", id + "Wrapper"), new XElement(One + "Table", new XAttribute("objectID", id),
            new XAttribute("bordersVisible", "true"), new XAttribute("hasHeaderRow", "false"),
            new XElement(One + "Columns", new XElement(One + "Column", new XAttribute("index", 0), new XAttribute("width", 300))),
            new XElement(One + "Row", new XElement(One + "Cell", new XElement(One + "OEChildren", lines)))));
    }
    private static string SpacingTexts(XElement page) => string.Join("|", page.Element(One + "Outline").Element(One + "OEChildren").Elements(One + "OE")
        .Select(oe => oe.Element(One + "Table") != null ? "#" : AgentCode.PlainText(oe)));
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
        /// <summary>最近一次提交的 XML。</summary>
        internal string LastXml;
        /// <summary>DeletePageContent 的次数；DeleteConflicts 大于 0 时先按时间戳冲突失败这么多次。</summary>
        internal int Deletes, DeleteConflicts;
        /// <summary>本机实测：写入只剩一行空白的文本框时 OneNote 直接删掉它。设为 true 模拟留着不删的情况。</summary>
        internal bool KeepEmptyOutlines;
        /// <summary>图片数据：CallbackID → one:Data。piBinaryData 读取时换成数据，找不到数据的坏图整个不出现（本机实测）。</summary>
        internal readonly Dictionary<string, string> Binary = new Dictionary<string, string>();
        private int _ids;
        internal FakePage(XElement page) { Page = new XElement(page); }
        /// <summary>直接提交快照里的草稿，返回写入后的页面。</summary>
        internal XElement Commit(AgentPageSnapshot snapshot) { new AgentCommitter(this).Commit(snapshot, CancellationToken.None); return Page; }
        public string GetPageContent(string id, PageInfo info)
        {
            if (FailReadAfterSave && Writes > 0) throw new Exception("read failed");
            if (info != PageInfo.piBinaryData) return Page.ToString();
            var copy = new XElement(Page);
            foreach (var image in copy.Descendants(One + "Image").ToList())
            {
                var callback = (string)image.Element(One + "CallbackID")?.Attribute("callbackID");
                if (callback == null || !Binary.TryGetValue(callback, out var data)) { image.Remove(); continue; }
                image.Element(One + "CallbackID").ReplaceWith(new XElement(One + "Data", data));
            }
            return copy.ToString();
        }
        public void DeletePageContent(string pageId, string objectId, DateTime expected)
        {
            True(expected == AgentPageSnapshot.Modified(Page));
            if (DeleteConflicts-- > 0) throw new COMException("conflict", unchecked((int)0x80042010));
            Page.Elements().Single(e => (string)e.Attribute("objectID") == objectId).Remove();
            Deletes++; Page.SetAttributeValue("lastModifiedTime", DateTime.UtcNow.AddSeconds(Deletes).ToString("o", CultureInfo.InvariantCulture));
        }
        public void UpdatePageContent(string xml, DateTime expected)
        {
            Attempts++; True(expected != DateTime.MinValue); LastXml = xml;
            if (ConflictsRemaining-- > 0) { OnConflict?.Invoke(); throw new COMException("conflict", unchecked((int)0x80042010)); }
            var homes = AgentLayout.Homes(Page);
            foreach (var c in XElement.Parse(xml).Elements())
            {
                // 本机实测：移到另一个文本框的对象一律按新对象建立，带着原 ID 也一样。
                var container = (string)c.Attribute("objectID") ?? c.Name.LocalName;
                foreach (var e in c.Descendants().Where(e => e.Attribute("objectID") != null && homes.TryGetValue((string)e.Attribute("objectID"), out var home) && home != container))
                    e.Attribute("objectID").Remove();
                // 图片带数据的存下数据；新建段落里只带 CallbackID 的图片没有数据，成了坏图。
                foreach (var image in c.Descendants(One + "Image"))
                {
                    var data = (string)image.Element(One + "Data");
                    if (data != null) { var callback = "cb-" + ++_ids; Binary[callback] = data; image.Element(One + "Data").ReplaceWith(new XElement(One + "CallbackID", new XAttribute("callbackID", callback))); }
                    else if (image.Parent?.Attribute("objectID") == null) image.Element(One + "CallbackID")?.SetAttributeValue("callbackID", "broken-" + ++_ids);
                }
                // 本机实测：已有段落省略 List 时 OneNote 保留原来的列表，只有重建的段落才没有列表。
                foreach (var oe in c.DescendantsAndSelf(One + "OE").Where(e => e.Attribute("objectID") != null && e.Element(One + "List") == null))
                {
                    var list = Page.Descendants(One + "OE").FirstOrDefault(e => (string)e.Attribute("objectID") == (string)oe.Attribute("objectID"))?.Element(One + "List");
                    if (list != null) oe.AddFirst(new XElement(list));
                }
                // OneNote 给新建的段落、表格分配 ID。
                var fresh = c.Name == One + "Outline" && c.Attribute("objectID") == null;
                foreach (var e in c.DescendantsAndSelf().Where(e => new[] { "Outline", "OE", "Table", "Row", "Cell" }.Contains(e.Name.LocalName) && e.Attribute("objectID") == null))
                    e.SetAttributeValue("objectID", "new-" + ++_ids);
                var identity = c.Name == One + "QuickStyleDef" || c.Name == One + "TagDef" ? "index" : "objectID";
                var current = Page.Elements(c.Name).FirstOrDefault(e => (string)e.Attribute(identity) == (string)c.Attribute(identity));
                // 本机实测：新建的文本框按位置排进页面 XML。
                var below = fresh ? Page.Elements(One + "Outline").FirstOrDefault(o => Top(o) > Top(c)) : null;
                if (below != null) below.AddBeforeSelf(new XElement(c));
                else if (current == null) Page.Add(new XElement(c)); else current.ReplaceWith(new XElement(c));
                var written = Page.Elements(c.Name).FirstOrDefault(e => (string)e.Attribute("objectID") == (string)c.Attribute("objectID"));
                if (!KeepEmptyOutlines && c.Name == One + "Outline" && written != null && written.Descendants(One + "OE").Count() == 1 && BlankLines.IsBlankLine(written.Descendants(One + "OE").Single()))
                    written.Remove();
            }
            Writes++; Page.SetAttributeValue("lastModifiedTime", DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture));
            AfterSave?.Invoke();
            if (ThrowAfterSave) throw new COMException("uncertain");
        }
        private static double Top(XElement outline) => (double?)outline.Element(One + "Position")?.Attribute("y") ?? 0;
    }
    private sealed class StubHttp : HttpMessageHandler
    {
        private readonly string _response, _media;
        internal string Body;
        internal int Requests;
        internal HttpContent ResponseContent;
        internal HttpStatusCode Status = HttpStatusCode.OK;
        internal StubHttp(string response, string media) { _response = response; _media = media; }
        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellation)
        {
            Requests++;
            Body = await request.Content.ReadAsStringAsync();
            return new HttpResponseMessage(Status) { Content = ResponseContent ?? new StringContent(_response, System.Text.Encoding.UTF8, _media) };
        }
    }
    /// <summary>已返回错误状态，但响应正文一直未就绪；释放响应时中断读取。</summary>
    private sealed class WaitingErrorContent : HttpContent
    {
        private readonly TaskCompletionSource<System.IO.Stream> _stream = new TaskCompletionSource<System.IO.Stream>();
        internal bool Disposed;
        protected override Task<System.IO.Stream> CreateContentReadStreamAsync() => _stream.Task;
        protected override Task SerializeToStreamAsync(System.IO.Stream stream, TransportContext context) => throw new NotSupportedException();
        protected override bool TryComputeLength(out long length) { length = 0; return false; }
        protected override void Dispose(bool disposing)
        {
            Disposed = true;
            _stream.TrySetException(new ObjectDisposedException(nameof(WaitingErrorContent)));
            base.Dispose(disposing);
        }
    }
    /// <summary>同步记下每条进度；Progress&lt;T&gt; 会投递到同步上下文，测试里看不到顺序。</summary>
    private sealed class ProgressLog : IProgress<AgentProgress>
    {
        internal readonly List<AgentProgress> Items = new List<AgentProgress>();
        public void Report(AgentProgress value) { lock (Items) Items.Add(value); }
    }
    /// <summary>在指定的几次请求上返回被 max_tokens 截断的输出（带半截工具调用），其余交给 inner；记下每次发出的消息。</summary>
    private sealed class TruncatingClient : IAgentChatClient
    {
        private readonly IAgentChatClient _inner; private readonly HashSet<int> _cut; private int _call;
        internal readonly List<string> Sent = new List<string>();
        internal TruncatingClient(IAgentChatClient inner, params int[] cut) { _inner = inner; _cut = new HashSet<int>(cut); }
        public Task<AgentReply> CompleteAsync(List<object> messages, object[] tools, IProgress<AgentProgress> progress, CancellationToken cancellation)
        {
            Sent.Add(AgentJson.Serialize(messages));
            if (!_cut.Contains(_call++)) return _inner.CompleteAsync(messages, tools, progress, cancellation);
            var reply = new AgentReply { FinishReason = "length", Reasoning = "TRUNCATED_THOUGHT" };
            reply.Calls.Add(0, new AgentToolCall { Id = "cut" + _call, Name = "fix_text", Arguments = "{\"snapshot_id\":" });
            return Task.FromResult(reply);
        }
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
    /// <summary>一直不收尾的模型：设完样式后反复看概况，只剩 finish_edit 可用时才提交。</summary>
    private sealed class WrapUpClient : IAgentChatClient
    {
        private readonly AgentPageSnapshot _s; private int _turn;
        private readonly bool _readAll, _change;
        internal readonly List<int> ToolCounts = new List<int>();
        internal string LastJson = "";
        internal WrapUpClient(AgentPageSnapshot s, bool readAll = false, bool change = true) { _s = s; _readAll = readAll; _change = change; }
        public Task<AgentReply> CompleteAsync(List<object> messages, object[] tools, IProgress<AgentProgress> progress, CancellationToken cancellation)
        {
            ToolCounts.Add(tools.Length);
            LastJson = AgentJson.Serialize(messages);
            string name; object args;
            if (tools.Length == 1) { name = "finish_edit"; args = new { snapshot_id = _s.SnapshotId, draft_revision = _s.Revision }; }
            else if (_turn == 0) { name = "read_blocks"; args = new { snapshot_id = _s.SnapshotId, block_ids = _readAll ? _s.Blocks.Where(b => b.Editable).Select(b => b.Id).ToArray() : new[] { "p1" } }; }
            else if (_turn == 1 && _change) { name = "set_paragraph_style"; args = new { snapshot_id = _s.SnapshotId, block_ids = new[] { "p1" }, preset_id = "heading1" }; }
            else { name = "get_page_overview"; args = new { }; }
            _turn++;
            var reply = new AgentReply { FinishReason = "tool_calls" };
            reply.Calls.Add(0, new AgentToolCall { Id = "w" + _turn, Name = name, Arguments = AgentChatClient.Serializer().Serialize(args) });
            return Task.FromResult(reply);
        }
    }
    private sealed class ScriptedClient : IAgentChatClient
    {
        private readonly AgentPageSnapshot _s; private int _turn;
        internal bool SawToolResult, SawReasoning;
        internal string LastJson = "";
        internal ScriptedClient(AgentPageSnapshot s) { _s = s; }
        public Task<AgentReply> CompleteAsync(List<object> messages, object[] tools, IProgress<AgentProgress> progress, CancellationToken cancellation)
        {
            var json = AgentChatClient.Serializer().Serialize(messages);
            LastJson = AgentJson.Serialize(messages);
            SawToolResult |= json.Contains("tool_call_id"); SawReasoning |= json.Contains("reasoning_content");
            string name; object args;
            switch (_turn++)
            {
                case 0: name = "get_page_overview"; args = new { }; break;
                case 1: name = "read_blocks"; args = new { snapshot_id = _s.SnapshotId, block_ids = _s.Blocks.Where(b => b.Editable || b.CodeCandidate).Select(b => b.Id).ToArray() }; break;
                case 2: name = "set_paragraph_style"; args = new { snapshot_id = _s.SnapshotId, block_ids = new[] { "p1" }, preset_id = "heading1" }; break;
                case 3: name = "get_pending_changes"; args = new { snapshot_id = _s.SnapshotId }; break;
                default: name = "finish_edit"; args = new { snapshot_id = _s.SnapshotId, draft_revision = _s.Revision }; break;
            }
            var reply = new AgentReply { FinishReason = "tool_calls", Reasoning = "opaque provider history" };
            reply.Calls.Add(0, new AgentToolCall { Id = "c" + _turn, Name = name, Arguments = AgentChatClient.Serializer().Serialize(args) });
            return Task.FromResult(reply);
        }
    }
    /// <summary>先漏读并尝试提交，再按工具返回的下一批 ID 读完；所有写入仅发生在 FakePage。</summary>
    private sealed class CoverageClient : IAgentChatClient
    {
        private readonly AgentPageSnapshot _s;
        private readonly Action _onRejection;
        private int _turn;
        internal bool SawRejection;
        internal CoverageClient(AgentPageSnapshot s, Action onRejection = null) { _s = s; _onRejection = onRejection; }
        public Task<AgentReply> CompleteAsync(List<object> messages, object[] tools, IProgress<AgentProgress> progress, CancellationToken cancellation)
        {
            string name; object args;
            if (_turn == 0) { name = "read_blocks"; args = new { snapshot_id = _s.SnapshotId, block_ids = new[] { "p1" } }; }
            else if (_turn == 1) { name = "set_paragraph_style"; args = new { snapshot_id = _s.SnapshotId, block_ids = new[] { "p1" }, preset_id = "heading1" }; }
            else if (_turn > 2 && _s.Blocks.Any(b => (b.Editable || b.CodeCandidate) && !b.Read))
            {
                var result = (string)ToolField(messages.Last(), "content");
                var outcome = AgentChatClient.Parse(result);
                if (!SawRejection)
                {
                    Equal(false, (bool)AiClient.Get(outcome, "ok")); SawRejection = true; _onRejection?.Invoke();
                    cancellation.ThrowIfCancellationRequested();
                }
                // 读取结果不重复返回进度；之后用 get_pending_changes 获取下一批。
                if (AiClient.Get(outcome, "next_read_block_ids") is object[] ids)
                { name = "read_blocks"; args = new { snapshot_id = _s.SnapshotId, block_ids = ids.Cast<string>().ToArray() }; }
                else { name = "get_pending_changes"; args = new { snapshot_id = _s.SnapshotId }; }
            }
            else { name = "finish_edit"; args = new { snapshot_id = _s.SnapshotId, draft_revision = _s.Revision }; }
            _turn++;
            var reply = new AgentReply { FinishReason = "tool_calls" };
            reply.Calls.Add(0, new AgentToolCall { Id = "coverage" + _turn, Name = name, Arguments = Json(args) });
            return Task.FromResult(reply);
        }
    }
}
