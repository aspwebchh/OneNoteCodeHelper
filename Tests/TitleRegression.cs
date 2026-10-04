using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using OneNoteCodeHelper.Services;
using OneNoteCodeHelper.Services.Agent;

internal static partial class Program
{
    private static XElement TitlePage(string text = "")
    {
        var page = Page(Paragraph("body", "正文"), Paragraph("blank", ""));
        page.AddFirst(new XElement(One + "Title", new XAttribute("style", "font-family:Arial;font-size:20pt"), Paragraph("title", text)));
        page.SetAttributeValue("name", text.Length == 0 ? "默认页面名称" : text);
        return page;
    }

    private static object SetTitle(AgentTools tools, AgentPageSnapshot s, string title, bool replace = false) =>
        Invoke(tools, "set_page_title", new { snapshot_id = s.SnapshotId, title, replace_existing = replace });

    private static string TitleText(XElement page) => AgentPageTitle.Text(page.Element(One + "Title"));

    private static void TestPageTitles()
    {
        Test("page title tool is page-only", () =>
        {
            var s = new AgentPageSnapshot(TitlePage().ToString(), null, new AgentOptions());
            var t = Tools(s); True(t.Has("set_page_title"));
            True(AgentRunner.SystemPrompt(t).Contains(AgentRunner.PageTitlePrompt));
            var selected = new AgentPageSnapshot(TitlePage().ToString(), new HashSet<string> { "body" }, new AgentOptions());
            var st = Tools(selected); True(!st.Has("set_page_title")); True(!AgentRunner.SystemPrompt(st).Contains(AgentRunner.PageTitlePrompt));
            Throws(() => SetTitle(st, selected, "越界标题"));
            var overview = Json(Invoke(t, "get_page_overview", new { }));
            True(overview.Contains("\"title_is_empty\":true")); True(overview.Contains("\"native_page_title\":\"\""));
        });
        Test("page title requires reading the page and invalid input leaves all drafts unchanged", () =>
        {
            var s = Snapshot(TitlePage()); var t = Tools(s);
            Rejects("完整读取", () => SetTitle(t, s, "标题")); Read(t, s);
            foreach (var invalid in new[] { "", "   ", "甲\n乙", "甲\r乙", "甲\t乙", "甲\u2028乙", "甲\u2029乙", new string('甲', 501) })
                Throws(() => SetTitle(t, s, invalid));
            foreach (var invalid in new[] { "甲￿乙", "甲￾乙" })
                Rejects("XML 不允许的字符", () => SetTitle(t, s, invalid));
            Equal(0, s.Revision); Equal(null, s.TitleEdit); Equal("", s.NativeTitle);
        });
        Test("blank native title writes escaped text and title-only edits undo without touching body blanks", () =>
        {
            var s = Snapshot(TitlePage()); var t = Tools(s); Read(t, s);
            var body = new XElement(s.Page.Element(One + "Outline"));
            const string title = "C# <b>泛型</b> & 配置 😀";
            SetTitle(t, s, title);
            Equal("empty", s.Blocks.Single(b => b.ObjectId == "blank").ProtectedReason);
            True(Json(Invoke(t, "get_pending_changes", new { snapshot_id = s.SnapshotId })).Contains("\"page_title_changed\":true"));
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); True(r.TitleChanged && r.CanUndo); Equal(0, r.Applied); Equal(0, r.Undo.Count);
            Equal(title, TitleText(api.Page)); Equal(title, (string)api.Page.Attribute("name"));
            True(api.LastXml.Contains("&lt;b&gt;泛型&lt;/b&gt;")); True(!api.LastXml.Contains("urn:onenote-code-helper"));
            True(XNode.DeepEquals(body, api.Page.Element(One + "Outline")));
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None);
            Equal("Verified", undo.Status); Equal("", TitleText(api.Page)); True(XNode.DeepEquals(body, api.Page.Element(One + "Outline")));
            undo.ClearUndo(); True(!undo.CanUndo);
        });
        Test("missing native title ignores fallback page name, supports styles and undo restores blank title", () =>
        {
            var page = Page(Paragraph("body", "自动作为页面名称的正文")); page.SetAttributeValue("name", "自动作为页面名称的正文");
            var s = Snapshot(page); var t = Tools(s); Read(t, s); SetTitle(t, s, "真正的标题");
            var b = s.Blocks.Single(x => x.IsPageTitle);
            True(Json(Invoke(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = new[] { b.Id } })).Contains("真正的标题"));
            Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { b.Id }, preset_id = "page_title", overrides = new { alignment = "center" } });
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal("真正的标题", TitleText(api.Page));
            Equal("center", (string)api.Page.Element(One + "Title").Element(One + "OE").Attribute("alignment"));
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status); Equal("", TitleText(api.Page));
            Equal("自动作为页面名称的正文", new AgentRichText(api.Page.Element(One + "Outline").Descendants(One + "OE").Single()).Text);
        });
        Test("a completely empty page can receive an explicit title without body text or read calls", () =>
        {
            var s = Snapshot(Page()); var api = new FakePage(s.Page); var c = new AgentCommitter(api);
            var t = new AgentTools(s, c, CancellationToken.None); Equal(0, t.UnreadCount);
            SetTitle(t, s, "指定的页面名称");
            Invoke(t, "finish_edit", new { snapshot_id = s.SnapshotId, draft_revision = s.Revision });
            Equal("Verified", t.Report.Status); Equal("指定的页面名称", TitleText(api.Page)); True(t.Report.CanUndo);
            Equal(0, api.Page.Element(One + "Outline").Descendants(One + "OE").Count());
            Equal("Verified", c.Undo(s.PageId, t.Report, s.Options, CancellationToken.None).Status); Equal("", TitleText(api.Page));
        });
        Test("title containers without OE, T or objectID can be filled and undone", () =>
        {
            foreach (var native in new[] { new XElement(One + "Title"), new XElement(One + "Title", new XElement(One + "OE")),
                new XElement(One + "Title", new XElement(One + "OE", new XElement(One + "T", ""))) })
            {
                var page = Page(Paragraph("body", "正文")); page.AddFirst(native);
                var s = Snapshot(page); var t = Tools(s); Read(t, s); SetTitle(t, s, "标题");
                var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
                Equal("Verified", r.Status); Equal("标题", TitleText(api.Page));
                Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status); Equal("", TitleText(api.Page));
            }
        });
        Test("existing title is preserved by default and explicit rename uses final valid draft", () =>
        {
            var s = Snapshot(TitlePage("原标题")); var t = Tools(s); Read(t, s);
            True(Json(SetTitle(t, s, "意外覆盖")).Contains("title_exists")); Equal(0, s.Revision); Equal(null, s.TitleEdit);
            SetTitle(t, s, "第一次改名", true); SetTitle(t, s, "最终标题", true);
            True(Json(SetTitle(t, s, "意外覆盖")).Contains("title_exists")); Equal("最终标题", s.NativeTitle);
            var revision = s.Revision; SetTitle(t, s, "最终标题", true); Equal(revision, s.Revision);
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(1, api.Writes); Equal("最终标题", TitleText(api.Page));
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status); Equal("原标题", TitleText(api.Page));
        });
        Test("repeated filling of an originally blank title refines one draft and finish reports verified title", () =>
        {
            var s = Snapshot(TitlePage()); var api = new FakePage(s.Page);
            var t = new AgentTools(s, new AgentCommitter(api), CancellationToken.None); Read(t, s);
            SetTitle(t, s, "初稿"); SetTitle(t, s, "终稿");
            Equal("终稿", s.NativeTitle); Equal(1, s.Blocks.Count(b => b.IsPageTitle));
            var result = Json(Invoke(t, "finish_edit", new { snapshot_id = s.SnapshotId, draft_revision = s.Revision }));
            True(result.Contains("\"page_title_changed\":true")); True(result.Contains("终稿")); True(t.Report.CanUndo); Equal(1, api.Writes);
        });
        Test("title text, title styles, body format and inserted paragraphs share one submission and undo", () =>
        {
            var s = Snapshot(TitlePage("原标题")); var t = Tools(s); Read(t, s);
            var b = s.Blocks.Single(x => x.IsPageTitle);
            Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { b.Id }, preset_id = "page_title", overrides = new { alignment = "center" } });
            SetTitle(t, s, "新标题", true);
            Invoke(t, "set_text_style", new { snapshot_id = s.SnapshotId, targets = new[] { new { block_id = b.Id, quote = "新标题", occurrence = 1, style = new { underline = true } } } });
            var body = s.Blocks.Single(x => x.ObjectId == "body");
            Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { body.Id }, preset_id = "heading2" });
            Insert(t, s, body.Id, "摘要");
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal(1, api.Writes); Equal(1, r.Applied); Equal(1, r.Inserted); True(r.TitleChanged);
            var title = api.Page.Element(One + "Title").Element(One + "OE");
            Equal("center", (string)title.Attribute("alignment")); True(title.Element(One + "T").Value.Contains("underline"));
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status);
            Equal("原标题", TitleText(api.Page)); Equal("原标题|正文|", Texts(api.Page));
            Equal(AgentPageSnapshot.SemanticFormat(s.Page.Element(One + "Title").Element(One + "OE"), s.Page),
                AgentPageSnapshot.SemanticFormat(api.Page.Element(One + "Title").Element(One + "OE"), api.Page));
        });
        Test("independent title draft survives undoing code-spacing drafts before further structural edits", () =>
        {
            var s = Snapshot(Page(Paragraph("a", "正文"), SpacingBox("code"), Paragraph("b", "结尾"))); var t = Tools(s); Read(t, s);
            Normalize(t, s); SetTitle(t, s, "标题");
            var id = s.Blocks.Single(b => b.ObjectId == "b").Id;
            Insert(t, s, id, "新增正文");
            var title = s.Blocks.Single(b => b.IsPageTitle);
            Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { title.Id }, preset_id = "page_title" });
            Normalize(t, s);
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal("标题", TitleText(api.Page)); True(r.Inserted > 0 && r.InsertedBlankLines > 0);
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status); Equal("", TitleText(api.Page));
        });
        Test("new title read, text fixes and clear-format share the same draft even without an original objectID", () =>
        {
            var s = Snapshot(Page(Paragraph("body", "正文"))); var t = Tools(s); Read(t, s); SetTitle(t, s, "按装说明");
            var b = s.Blocks.Single(x => x.IsPageTitle);
            Fix(t, s, b.Id, "按装", "安装"); Equal("安装说明", s.NativeTitle);
            Invoke(t, "set_text_style", new { snapshot_id = s.SnapshotId, targets = new[] { new { block_id = b.Id, quote = "安装", occurrence = 1, style = new { bold = true } } } });
            Invoke(t, "clear_format", new { snapshot_id = s.SnapshotId, block_ids = new[] { b.Id } });
            var api = new FakePage(s.Page); var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal("安装说明", TitleText(api.Page)); Equal(1, r.TextFixes.Count);
            True(!api.Page.Element(One + "Title").Element(One + "OE").Element(One + "T").Value.Contains("font-weight:bold"));
        });
        Test("whitespace-only original titles count as blank and undo preserves their empty-title format", () =>
        {
            var s = Snapshot(TitlePage("&nbsp;")); var t = Tools(s); Read(t, s); SetTitle(t, s, "标题");
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            Equal("Verified", r.Status); Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status); Equal("", TitleText(api.Page));
        });
        Test("title edit conflicts skip the entire title while independent body edits still commit", () =>
        {
            var s = Snapshot(TitlePage()); var t = Tools(s); Read(t, s); SetTitle(t, s, "Agent 标题");
            var body = s.Blocks.Single(b => b.ObjectId == "body");
            Invoke(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { body.Id }, preset_id = "heading1" });
            var api = new FakePage(s.Page); api.Page.Element(One + "Title").Element(One + "OE").Element(One + "T").Value = "用户标题";
            var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
            Equal("PartiallyApplied", r.Status); Equal(1, r.Conflicts); Equal(1, r.Applied); True(!r.TitleChanged);
            Equal("用户标题", TitleText(api.Page)); True(XElement.Parse(api.LastXml).Element(One + "Title") == null);
        });
        Test("title fingerprint includes referenced style definitions but ignores unrelated body edits", () =>
        {
            var page = TitlePage(); page.AddFirst(new XElement(One + "QuickStyleDef", new XAttribute("index", 0), new XAttribute("name", "p"), new XAttribute("fontSize", 20)));
            page.Element(One + "Title").Element(One + "OE").SetAttributeValue("quickStyleIndex", 0);
            var s = Snapshot(page); var t = Tools(s); Read(t, s); SetTitle(t, s, "标题");
            var changed = new FakePage(s.Page); changed.Page.Element(One + "QuickStyleDef").SetAttributeValue("fontSize", 24);
            var r = new AgentCommitter(changed).Commit(s, CancellationToken.None); Equal(1, r.Conflicts); Equal(0, changed.Writes);
            var api = new FakePage(s.Page); api.Page.Element(One + "Outline").Descendants(One + "T").First().Value = "用户修改的正文";
            r = new AgentCommitter(api).Commit(s, CancellationToken.None); Equal("Verified", r.Status); Equal("用户修改的正文", api.Page.Element(One + "Outline").Descendants(One + "T").First().Value);
        });
        Test("title timestamp conflicts reread before retry and skip a title edited during retry", () =>
        {
            foreach (var changeTitle in new[] { false, true })
            {
                var s = Snapshot(TitlePage()); var t = Tools(s); Read(t, s); SetTitle(t, s, "Agent 标题");
                var api = new FakePage(s.Page) { ConflictsRemaining = 1 };
                api.OnConflict = () => api.Page.Descendants(One + "T").First(e => changeTitle ? e.Ancestors(One + "Title").Any() : e.Ancestors(One + "Outline").Any()).Value = "用户修改";
                var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
                Equal(changeTitle ? "NoChange" : "Verified", r.Status); Equal(changeTitle ? 0 : 1, api.Writes);
                Equal(changeTitle ? "用户修改" : "Agent 标题", TitleText(api.Page));
            }
        });
        Test("title-only uncertain writes are verified by readback and read failures remain unknown", () =>
        {
            foreach (var failRead in new[] { false, true })
            {
                var s = Snapshot(TitlePage()); var t = Tools(s); Read(t, s); SetTitle(t, s, "标题");
                var api = new FakePage(s.Page) { ThrowAfterSave = true, FailReadAfterSave = failRead };
                var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
                Equal(failRead ? "CommitOutcomeUnknown" : "Verified", r.Status); Equal(!failRead, r.TitleChanged); Equal(!failRead, r.CanUndo);
                if (failRead) Equal(1, r.Unverified);
            }
        });
        Test("title verification rejects wrong native text or page name and catches body damage", () =>
        {
            foreach (var fault in new[] { "text", "name", "body" })
            {
                var s = Snapshot(TitlePage()); var t = Tools(s); Read(t, s); SetTitle(t, s, "标题");
                var api = new FakePage(s.Page);
                api.AfterSave = () => { if (fault == "name") api.Page.SetAttributeValue("name", "错误名称");
                    else api.Page.Element(One + (fault == "text" ? "Title" : "Outline")).Descendants(One + "T").First().Value = "意外改动"; };
                var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
                Equal(fault == "body" ? "CommitOutcomeUnknown" : "PartiallyApplied", r.Status);
                if (fault != "body") { True(!r.TitleChanged && !r.CanUndo); Equal(1, r.Unverified); }
            }
        });
        Test("native title emoji font fallback is accepted while ordinary text font changes are rejected", () =>
        {
            foreach (var font in new[] { "Arial", "Webdings" })
            {
                var s = Snapshot(TitlePage()); var t = Tools(s); Read(t, s); SetTitle(t, s, "说明😀");
                var api = new FakePage(s.Page);
                api.AfterSave = () => api.Page.Element(One + "Title").Element(One + "OE").Element(One + "T").Value =
                    "<span style='font-family:" + font + "'>说明</span><span style='font-family:Segoe UI Emoji'>😀</span>";
                var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
                Equal(font == "Arial" ? "Verified" : "PartiallyApplied", r.Status); Equal(font == "Arial", r.TitleChanged);
            }
        });
        Test("title undo does not overwrite user title edits and blank native placeholders are accepted", () =>
        {
            var s = Snapshot(TitlePage("原标题")); var t = Tools(s); Read(t, s); SetTitle(t, s, "新标题", true);
            var api = new FakePage(s.Page); var c = new AgentCommitter(api); var r = c.Commit(s, CancellationToken.None);
            api.Page.Element(One + "Title").Element(One + "OE").Element(One + "T").Value = "用户再次修改";
            var undo = c.Undo(s.PageId, r, s.Options, CancellationToken.None); Equal(1, undo.Conflicts); Equal("用户再次修改", TitleText(api.Page)); Equal(1, api.Writes);
            s = Snapshot(TitlePage()); t = Tools(s); Read(t, s); SetTitle(t, s, "标题");
            api = new FakePage(s.Page); c = new AgentCommitter(api); r = c.Commit(s, CancellationToken.None);
            api.AfterSave = () => api.Page.Element(One + "Title").Remove();
            Equal("Verified", c.Undo(s.PageId, r, s.Options, CancellationToken.None).Status); Equal("", TitleText(api.Page));
        });
        Test("unsupported title content is rejected without relaxing paragraph protection", () =>
        {
            var page = TitlePage(); page.Element(One + "Title").Element(One + "OE").Add(new XElement(One + "Image"));
            var s = Snapshot(page); var t = Tools(s); Read(t, s);
            Rejects("不支持", () => SetTitle(t, s, "标题")); Equal(null, s.TitleEdit); Equal(0, s.Revision);
            page = TitlePage(); page.Element(One + "Title").Add(Paragraph("second-title", "多余段落"));
            s = Snapshot(page); t = Tools(s); Read(t, s); Throws(() => SetTitle(t, s, "标题")); Equal(null, s.TitleEdit);
        });
        Test("title steps explain preserved originals without exposing title text", () =>
        {
            Equal(("设置页面标题 · 保留原标题", AgentStepState.Done), AgentTools.DescribeStep("set_page_title", "{\"title\":\"PRIVATE\"}", "{\"ok\":true,\"changed\":false,\"reason\":\"title_exists\"}"));
            True(!AgentTools.DescribeStep("set_page_title", "{\"title\":\"PRIVATE\"}", "{\"ok\":true,\"changed\":true}").Text.Contains("PRIVATE"));
        });
    }
}
