using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using OneNoteCodeHelper.Services.Agent;

internal static partial class Program
{
    private static void ClearFormatRegressionTests()
    {
        Test("format projection includes native highlight, strike and scripts and normalizes absent backgrounds", () =>
        {
            foreach (var script in new[] { "superscript", "subscript" })
            {
                var oe = Paragraph("a", "文字"); oe.SetAttributeValue("quickStyleIndex", "0");
                var page = Page(oe);
                var definition = ParagraphStyles.Definition("body", new AgentOptions()); definition.SetAttributeValue("index", "0");
                definition.SetAttributeValue("highlightColor", "#FFFF00"); definition.SetAttributeValue("strikethrough", "true");
                definition.SetAttributeValue("underline", "true"); definition.SetAttributeValue(script, "true"); page.AddFirst(definition);
                var style = Css.Effective(oe.Element(One + "T"), page);
                Equal("#ffff00", style["background-color"]); Equal("underline line-through", style["text-decoration"]);
                Equal(script == "superscript" ? "super" : "sub", style["vertical-align"]);
            }
            string Look(string style)
            {
                var oe = Paragraph("a", "文字"); oe.SetAttributeValue("style", style);
                return new AgentRichText(oe).Signature(Page(oe), true);
            }
            foreach (var style in new[] { "background:none", "background-color:transparent", "background:automatic", "background:rgba(0, 0, 0, 0)" })
                Equal(Look(null), Look(style));
            Equal("transparent", Css.Read("background:yellow;background-color:none")["background-color"]);
            Equal("yellow", Css.Read("background-color:none;background:yellow")["background-color"]);
        });
        Test("clearing only a styled parent resets its own text and preserves the child with either native heading setting", () =>
        {
            foreach (var native in new[] { true, false })
            foreach (var style in new[] { "font-weight:bold", "font-style:italic", "text-decoration:underline line-through", "background:yellow", "vertical-align:super", "vertical-align:sub" })
            foreach (var quick in new[] { false, true })
            {
                var parent = Paragraph("parent", "父段"); parent.SetAttributeValue("style", "font-size:20pt;color:red;" + style);
                parent.Add(new XElement(One + "OEChildren", Paragraph("child", "子段")));
                var page = Page(parent);
                if (quick)
                {
                    var definition = ParagraphStyles.Definition("heading1", new AgentOptions()); definition.SetAttributeValue("index", "0");
                    definition.SetAttributeValue("italic", "true"); definition.SetAttributeValue("highlightColor", "#FFFF00");
                    definition.SetAttributeValue("strikethrough", "true"); definition.SetAttributeValue("superscript", "true");
                    page.AddFirst(definition); parent.SetAttributeValue("quickStyleIndex", "0");
                }
                var s = new AgentPageSnapshot(page.ToString(), new HashSet<string> { "parent" }, new AgentOptions { EnableNativeHeadings = native });
                var t = Tools(s); ReadAll(t, s);
                var childBefore = AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(s.Page, "child"), s.Page);
                var parentBefore = AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(s.Page, "parent"), s.Page);
                Clear(t, s, "p1");
                var revision = s.Revision; Clear(t, s, "p1"); Equal(revision, s.Revision);
                var api = new FakePage(s.Page); var c = new AgentCommitter(api); var done = c.Commit(s, CancellationToken.None);
                Equal("Verified", done.Status); Equal(1, done.Applied); Equal("p1", string.Join(",", done.AppearanceOnly));
                ClearedBody(AgentCommitter.Find(api.Page, "parent"), api.Page, s.Options);
                Equal(childBefore, AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "child"), api.Page));
                Equal("Verified", c.Undo(s.PageId, done, s.Options, CancellationToken.None).Status);
                Equal(parentBefore, AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "parent"), api.Page));
                Equal(childBefore, AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "child"), api.Page));
            }
        });
        Test("clearing child or parent-child batch overrides paragraph and outline inheritance including empty and multiple T runs", () =>
        {
            foreach (var native in new[] { true, false })
            foreach (var outline in new[] { true, false })
            foreach (var all in new[] { true, false })
            {
                var parent = Paragraph("parent", "父段");
                parent.Add(new XElement(One + "OEChildren", Paragraph("child", "甲", "<b>乙</b><sup>2</sup>", ""), Paragraph("empty", "")));
                var page = Page(parent);
                (outline ? page.Element(One + "Outline") : parent).SetAttributeValue("style",
                    "font-family:Calibri;font-size:20pt;color:red;font-weight:bold;font-style:italic;text-decoration:underline line-through;vertical-align:sub;background-color:yellow");
                var s = new AgentPageSnapshot(page.ToString(), null, new AgentOptions { EnableNativeHeadings = native, FontFamily = "Arial" });
                var t = Tools(s); ReadAll(t, s);
                var parentBefore = AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(s.Page, "parent"), s.Page);
                Clear(t, s, all ? new[] { "p1", "p2", "p3" } : new[] { "p2", "p3" });
                var revision = s.Revision; Clear(t, s, all ? new[] { "p1", "p2", "p3" } : new[] { "p2", "p3" }); Equal(revision, s.Revision);
                var api = new FakePage(s.Page); var c = new AgentCommitter(api); var done = c.Commit(s, CancellationToken.None);
                Equal("Verified", done.Status); Equal(Texts(s.Page), Texts(api.Page));
                foreach (var id in all ? new[] { "parent", "child", "empty" } : new[] { "child", "empty" }) ClearedBody(AgentCommitter.Find(api.Page, id), api.Page, s.Options);
                if (!all) Equal(parentBefore, AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "parent"), api.Page));
                Equal("Verified", c.Undo(s.PageId, done, s.Options, CancellationToken.None).Status);
                foreach (var id in new[] { "parent", "child", "empty" })
                    Equal(AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(s.Page, id), s.Page), AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, id), api.Page));
            }
        });
        Test("clearing preserves emoji and symbol fonts from inline T paragraph and outline styles", () =>
        {
            foreach (var font in new[] { "Segoe UI Emoji", "Segoe UI Symbol" })
            foreach (var level in new[] { "span", "T", "OE", "Outline" })
            {
                var oe = Paragraph("a", level == "span" ? "<span style='font-family:" + font + ";color:red'>😀</span>" : "😀");
                var page = Page(oe);
                if (level != "span") (level == "T" ? oe.Element(One + "T") : level == "OE" ? oe : page.Element(One + "Outline"))
                    .SetAttributeValue("style", "font-family:" + font + ";font-size:20pt;color:red");
                var s = Snapshot(page); var t = Tools(s); ReadAll(t, s); Clear(t, s, "p1");
                var api = new FakePage(s.Page); var c = new AgentCommitter(api); var done = c.Commit(s, CancellationToken.None);
                Equal("Verified", done.Status);
                True(new AgentRichText(AgentCommitter.Find(api.Page, "a")).Signature(api.Page, true).Contains("font-family:" + font.ToLowerInvariant()));
                Equal("Verified", c.Undo(s.PageId, done, s.Options, CancellationToken.None).Status);
                Equal(AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(s.Page, "a"), s.Page), AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(api.Page, "a"), api.Page));
            }
        });
        Test("cleared text and empty T reject restored highlight scripts and native decorations in both commit paths", () =>
        {
            foreach (var empty in new[] { true, false })
            foreach (var structural in new[] { true, false })
            foreach (var corruption in new[] { "background-color:yellow", "vertical-align:super", "native" })
            {
                var s = Snapshot(Page(Paragraph("a", empty ? "" : "正文"), Paragraph("b", "参照"))); var t = Tools(s); ReadAll(t, s); Clear(t, s, "p1");
                if (structural) Insert(t, s, "p2", "插入");
                var api = new FakePage(s.Page);
                api.AfterSave = () =>
                {
                    var run = AgentCommitter.Find(api.Page, "a").Element(One + "T");
                    if (corruption == "native")
                    {
                        var definition = ParagraphStyles.Definition("body", s.Options); definition.SetAttributeValue("highlightColor", "#FFFF00");
                        definition.SetAttributeValue("superscript", "true"); definition.SetAttributeValue("strikethrough", "true");
                        run.SetAttributeValue("quickStyleIndex", ParagraphStyles.EnsureDefinition(api.Page, definition)); run.SetAttributeValue("style", null);
                    }
                    else run.SetAttributeValue("style", (string)run.Attribute("style") + ";" + corruption);
                };
                var done = new AgentCommitter(api).Commit(s, CancellationToken.None);
                Equal("PartiallyApplied", done.Status); True(done.Unverified > 0);
                True(done.Undo.All(u => u.ObjectId != "a"));
            }
        });
        Test("cleared text tolerates equivalent transparent background readback without overlooking inherited highlight", () =>
        {
            foreach (var style in new[] { "", "background:none", "background-color:rgba(0, 0, 0, 0)" })
            {
                var s = Snapshot(Page(Paragraph("a", "正文"))); var t = Tools(s); ReadAll(t, s); Clear(t, s, "p1");
                var api = new FakePage(s.Page);
                api.AfterSave = () =>
                {
                    var run = AgentCommitter.Find(api.Page, "a").Element(One + "T");
                    run.SetAttributeValue("style", ((string)run.Attribute("style")).Replace("background-color:transparent", style));
                };
                Equal("Verified", new AgentCommitter(api).Commit(s, CancellationToken.None).Status);
            }
        });
        Test("unwrap links option keeps or removes links counts only actual changes and undo restores the original box", () =>
        {
            foreach (var remove in new bool?[] { null, true, false })
            {
                var s = Snapshot(Page(Paragraph("body", "<a href='https://example.com/body'>正文链接</a>"),
                    CodeBox("box", "<a href='https://example.com/code'>x = 1</a>", "&nbsp;")));
                var t = Tools(s); ReadAll(t, s);
                Invoke(t, "clear_format", new { snapshot_id = s.SnapshotId, block_ids = new[] { "p1" }, links = false });
                object args = remove.HasValue ? (object)new { snapshot_id = s.SnapshotId, table_ids = new[] { "t1" }, links = remove.Value }
                    : new { snapshot_id = s.SnapshotId, table_ids = new[] { "t1" } };
                var result = Invoke(t, "unwrap_code", args); var count = remove == false ? 0 : 1;
                Equal(count, (int)ToolField(result, "links_removed"));
                var revision = s.Revision; var repeated = Invoke(t, "unwrap_code", args); Equal(revision, s.Revision); Equal(0, (int)ToolField(repeated, "links_removed"));
                Equal(count, (int)ToolField(Invoke(t, "get_pending_changes", new { snapshot_id = s.SnapshotId }), "links_removed"));
                var api = new FakePage(s.Page); var c = new AgentCommitter(api); var done = c.Commit(s, CancellationToken.None);
                Equal("Verified", done.Status); Equal(1, done.Unwrapped); Equal(count, done.LinksRemoved);
                Equal(1, new AgentRichText(ByText(api.Page, "正文链接")).LinkCount); Equal(remove == false ? 1 : 0, new AgentRichText(ByText(api.Page, "x = 1")).LinkCount);
                var undo = c.Undo(s.PageId, done, s.Options, CancellationToken.None); Equal("Verified", undo.Status); Equal(count, undo.LinksRemoved);
                Equal(1, new AgentRichText(ByText(api.Page, "x = 1")).LinkCount); Equal(Texts(s.Page), Texts(api.Page));
            }
        });
        Test("unwrap links counts skip conflicts and failed boxes including whitespace-only code lines", () =>
        {
            foreach (var mode in new[] { "conflict", "format", "empty" })
            {
                var s = Snapshot(Page(CodeBox("one", "<a href='https://example.com/one'>x = 1</a>"), CodeBox("two", "<a href='https://example.com/two'>y = 2</a>", "&nbsp;")));
                var t = Tools(s); Invoke(t, "unwrap_code", new { snapshot_id = s.SnapshotId, table_ids = new[] { "t1", "t2" } });
                var api = new FakePage(s.Page);
                if (mode == "conflict") ByText(api.Page, "x = 1").Element(One + "T").Value = "用户改动";
                else api.AfterSave = () =>
                {
                    var line = mode == "empty" ? api.Page.Descendants(One + "OE").Last(e => e.Elements(One + "T").Any()) : ByText(api.Page, "y = 2");
                    var run = line.Element(One + "T");
                    if (mode == "empty") run.Value = "";
                    run.SetAttributeValue("style", (string)run.Attribute("style") + ";background:yellow");
                };
                var c = new AgentCommitter(api); var done = c.Commit(s, CancellationToken.None);
                Equal(mode == "conflict" ? "NoChange" : "PartiallyApplied", done.Status); Equal(mode == "conflict" ? 0 : 1, done.Unwrapped); Equal(mode == "conflict" ? 0 : 1, done.LinksRemoved);
                if (mode == "conflict") Equal(0, api.Writes); else True(done.Unverified > 0);
            }
        });
        Test("unwrap link preservation and removal survive a cross-outline merge and grouped undo", () =>
        {
            foreach (var remove in new[] { true, false })
            {
                var page = Page(Paragraph("target", "参照"));
                page.Add(new XElement(One + "Outline", new XAttribute("objectID", "source"), new XAttribute("style", "font-family:Arial;font-size:11pt;color:#222222"),
                    new XElement(One + "OEChildren", CodeBox("box", "<a href='https://example.com/code'>x = 1</a>"))));
                var s = Snapshot(page); var t = Tools(s); ReadAll(t, s);
                Invoke(t, "unwrap_code", new { snapshot_id = s.SnapshotId, table_ids = new[] { "t1" }, links = remove });
                Merge(t, s, "source", "p1", "after");
                var api = new FakePage(s.Page); var c = new AgentCommitter(api); var done = c.Commit(s, CancellationToken.None);
                Equal("Verified", done.Status); Equal(1, done.Unwrapped); Equal(remove ? 1 : 0, done.LinksRemoved);
                Equal(remove ? 0 : 1, new AgentRichText(ByText(api.Page, "x = 1")).LinkCount);
                Equal("Verified", c.Undo(s.PageId, done, s.Options, CancellationToken.None).Status);
                Equal(1, new AgentRichText(ByText(api.Page, "x = 1")).LinkCount); Equal(2, api.Page.Elements(One + "Outline").Count());
                // 任一组员后来被改，整组撤销都要跳过。
                var other = new FakePage(s.Page); var otherCommitter = new AgentCommitter(other); var applied = otherCommitter.Commit(s, CancellationToken.None);
                ByText(other.Page, "参照").Element(One + "T").Value = "后来修改";
                var undo = otherCommitter.Undo(s.PageId, applied, s.Options, CancellationToken.None); True(undo.Conflicts > 0); Equal(1, other.Writes);
            }
        });
        Test("runner preserves ordinary and code links when clearing and prompts both tools to receive links false", () =>
        {
            var page = ClearPage(); AgentCommitter.Find(page, "box1").Element(One + "T").Value = "<a href='https://example.com/code'>if x:</a>";
            var s = Snapshot(page); var api = new FakePage(s.Page); var c = new AgentCommitter(api);
            True(AgentRunner.SystemPrompt(Tools(s)).Contains("clear_format 和 unwrap_code 都传 links=false"));
            var report = new AgentRunner(new ClearScriptClient(s, true), c).RunAsync(s, "清除所有格式，包括代码框，保留链接", null, CancellationToken.None).GetAwaiter().GetResult();
            Equal("Verified", report.Status); Equal(1, report.Unwrapped); Equal(0, report.LinksRemoved);
            Equal(1, new AgentRichText(ByText(api.Page, "重点")).LinkCount); Equal(1, new AgentRichText(ByText(api.Page, "if x:")).LinkCount);
            Equal("Verified", c.Undo(s.PageId, report, s.Options, CancellationToken.None).Status);
        });
    }

    private static void ClearedBody(XElement oe, XElement page, AgentOptions options)
    {
        foreach (var run in oe.Elements(One + "T"))
        {
            var style = Css.Effective(run, page);
            Equal(Css.Normalize(options.FontFamily), style["font-family"]); Equal("11pt", style["font-size"]); Equal("#222222", style["color"]);
            Equal("normal", style["font-weight"]); Equal("normal", style["font-style"]); Equal("none", style["text-decoration"]);
            Equal("baseline", style["vertical-align"]); Equal("transparent", style["background-color"]);
        }
        var signature = new AgentRichText(oe).Signature(page, true);
        True(!signature.Contains("line-through") && !signature.Contains("vertical-align:super") && !signature.Contains("vertical-align:sub") && !signature.Contains("background-color:yellow"));
    }
}
