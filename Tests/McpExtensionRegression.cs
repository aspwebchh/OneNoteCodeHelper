using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using OneNoteCodeHelper.Services;
using OneNoteCodeHelper.Services.Agent;
using OneNoteCodeHelper.Services.Mcp;

internal static partial class Program
{
    private sealed class FakeWorkspace : IOneNoteWorkspaceAccess
    {
        internal XElement Hierarchy = new XElement(One + "Notebooks", new XElement(One + "Notebook", new XAttribute("ID", "book"), new XAttribute("name", "笔记本"),
            new XElement(One + "Section", new XAttribute("ID", "s1"), new XAttribute("name", "工作")),
            new XElement(One + "SectionGroup", new XAttribute("ID", "group"), new XAttribute("name", "资料"), new XElement(One + "Section", new XAttribute("ID", "s2"), new XAttribute("name", "归档")))));
        internal readonly Dictionary<string, FakePage> Pages = new Dictionary<string, FakePage>();
        internal int Creates, Recycles, Navigations, Searches, Reads;
        internal bool FailSearch, FailRecycle, FailCreate, FailExport, RecycleThrowsAfterApplied, FailHierarchy;
        internal Action AfterCreate, AfterUpdate, AfterRecycle;
        internal string Current = "page";
        internal readonly HashSet<string> Unreadable = new HashSet<string>();
        private int _ids;
        internal void Add(string id, XElement page, string section = "s1")
        {
            page.SetAttributeValue("ID", id); Pages[id] = new FakePage(page); Stamp(Pages[id].Page);
            Node(section).Add(new XElement(One + "Page", new XAttribute("ID", id), new XAttribute("name", (string)page.Attribute("name") ?? "测试"), new XAttribute("pageLevel", 1)));
        }
        /// <summary>
        /// 按本机实测补上 OneNote 回存时自行维护的内容：页面日期、层级和语言，页面设置，作者和编辑记录，默认左对齐，
        /// 表格开关，列表字号和标记日期，自动计算的文本框宽度，标题表情的符号字体，以及 36.0 这样的数值写法。写入只带草稿内容，回读仍会出现这些。
        /// </summary>
        private static void Stamp(XElement page)
        {
            void Default(XElement e, string name, string value) { if (e.Attribute(name) == null) e.SetAttributeValue(name, value); }
            const string Time = "2026-09-26T00:00:00.000Z";
            Default(page, "dateTime", Time); Default(page, "pageLevel", "1"); Default(page, "lang", "zh-CN");
            if (page.Element(One + "PageSettings") == null)
            {
                var settings = new XElement(One + "PageSettings", new XAttribute("RTL", "false"), new XAttribute("color", "automatic"),
                    new XElement(One + "PageSize", new XElement(One + "Automatic")), new XElement(One + "RuleLines", new XAttribute("visible", "false")));
                var anchor = page.Elements().FirstOrDefault(e => e.Name == One + "Title" || e.Name == One + "Outline" || e.Name == One + "Image");
                if (anchor != null) anchor.AddBeforeSelf(settings); else page.Add(settings);
            }
            foreach (var e in page.Elements(One + "Title").Concat(page.Descendants(One + "Cell"))) Default(e, "lang", "zh-CN");
            // 标题里的表情换成符号字体。
            foreach (var t in page.Elements(One + "Title").Descendants(One + "T").Where(t => !t.Value.Contains("Segoe UI Emoji")))
                t.ReplaceNodes(new XCData(System.Text.RegularExpressions.Regex.Replace(t.Value, @"[\uD800-\uDBFF][\uDC00-\uDFFF]", m => "<span style='font-family:\"Segoe UI Emoji\"'>" + m.Value + "</span>")));
            foreach (var e in page.Descendants().Where(e => e.Name == One + "OE" || e.Name == One + "Outline" || e.Name == One + "Cell"))
            {
                Default(e, "author", "测试者"); Default(e, "authorInitials", "测"); Default(e, "lastModifiedBy", "测试者"); Default(e, "lastModifiedByInitials", "测");
                Default(e, "lastModifiedTime", Time);
                if (e.Name != One + "OE") continue;
                Default(e, "creationTime", Time); Default(e, "alignment", "left");
                Default(e, "authorResolutionID", "<resolutionId provider=\"Windows Live\" hash=\"test\"/>"); Default(e, "lastModifiedByResolutionID", "<resolutionId provider=\"Windows Live\" hash=\"test\"/>");
            }
            foreach (var table in page.Descendants(One + "Table")) { Default(table, "bordersVisible", "false"); Default(table, "hasHeaderRow", "false"); Default(table, "lastModifiedTime", Time); }
            foreach (var mark in page.Descendants(One + "Bullet").Concat(page.Descendants(One + "Number"))) Default(mark, "fontSize", "11.0");
            foreach (var tag in page.Descendants(One + "Tag")) { Default(tag, "disabled", "false"); Default(tag, "creationDate", Time); }
            // 带标记的段落把段落间距再写进 style 的 margin；一组编号的第一项省略默认起点 1。
            foreach (var oe in page.Descendants(One + "OE").Where(o => o.Element(One + "Tag") != null && (o.Attribute("spaceBefore") != null || o.Attribute("spaceAfter") != null) && !((string)o.Attribute("style") ?? "").Contains("margin")))
                oe.SetAttributeValue("style", ((string)oe.Attribute("style") is string css && css.Length > 0 ? css + ";" : "") +
                    "margin-top:" + ((string)oe.Attribute("spaceBefore") ?? "0") + "pt;margin-bottom:" + ((string)oe.Attribute("spaceAfter") ?? "0") + "pt");
            foreach (var number in page.Descendants(One + "Number").Where(n => (string)n.Attribute("restartNumberingAt") == "1").ToList())
                if (number.Parent?.Parent?.ElementsBeforeSelf(One + "OE").LastOrDefault()?.Element(One + "List")?.Element(One + "Number") == null)
                    number.SetAttributeValue("restartNumberingAt", null);
            foreach (var outline in page.Elements(One + "Outline"))
            {
                var size = outline.Element(One + "Size");
                if (size == null) { size = new XElement(One + "Size"); if (outline.Element(One + "Position") is XElement position) position.AddAfterSelf(size); else outline.AddFirst(size); }
                // 未由用户设定的宽度由 OneNote 按内容计算，和写入时给的不同。
                if ((string)size.Attribute("isSetByUser") != "true") size.SetAttributeValue("width", "480.25");
                Default(size, "height", "120.5");
            }
            foreach (var a in page.Descendants().Where(e => e.Name == One + "Position" || e.Name == One + "Size" || e.Name == One + "Column").Attributes()
                .Where(a => a.Name.LocalName != "z" && a.Name.LocalName != "index" && !a.Value.Contains(".") && double.TryParse(a.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out _)).ToList())
                a.Value += ".0";
        }
        private XElement Node(string id) => Hierarchy.Descendants().First(e => (string)e.Attribute("ID") == id);
        public string GetHierarchy(string id, HierarchyScope scope)
        {
            if (FailHierarchy) throw new Exception("hierarchy unavailable");
            foreach (var p in Hierarchy.Descendants(One + "Page")) if (Pages.TryGetValue((string)p.Attribute("ID"), out var saved)) p.SetAttributeValue("lastModifiedTime", (string)saved.Page.Attribute("lastModifiedTime"));
            var notebook = Hierarchy.Element(One + "Notebook");
            if (!notebook.Elements(One + "SectionGroup").Any(e => (string)e.Attribute("isRecycleBin") == "true"))
                notebook.Add(new XElement(One + "SectionGroup", new XAttribute("ID", "recycle"), new XAttribute("isRecycleBin", "true")));
            var tree = id == "" || id == null ? new XElement(Hierarchy) : new XElement(Node(id));
            if (scope == HierarchyScope.hsNotebooks) foreach (var e in tree.Descendants(One + "Notebook").ToList()) e.RemoveNodes();
            else if (scope == HierarchyScope.hsSections) tree.Descendants(One + "Page").Remove();
            return tree.ToString(SaveOptions.DisableFormatting);
        }
        public string FindPages(string root, string query)
        {
            Searches++; if (FailSearch) throw new Exception("search unavailable");
            var tree = McpReadService.Xml(GetHierarchy(root, HierarchyScope.hsPages));
            foreach (var page in tree.Descendants(One + "Page").Where(p => !Pages[(string)p.Attribute("ID")].Page.Value.Contains(query) && !((string)p.Attribute("name")).Contains(query)).ToList()) page.Remove();
            return tree.ToString();
        }
        public string GetPageContent(string id, PageInfo info) { Reads++; if (Unreadable.Contains(id)) throw new Exception("page locked"); return Pages[id].GetPageContent(id, info); }
        public void UpdatePageContent(string xml, DateTime expected)
        {
            var changes = XElement.Parse(xml); var id = (string)changes.Attribute("ID"); var page = Pages[id];
            page.UpdatePageContent(xml, expected);
            if (changes.Attribute("name") != null) page.Page.SetAttributeValue("name", (string)changes.Attribute("name"));
            Stamp(page.Page); Node(id).SetAttributeValue("name", (string)page.Page.Attribute("name")); AfterUpdate?.Invoke();
        }
        public void DeletePageContent(string page, string objectId, DateTime expected) => Pages[page].DeletePageContent(page, objectId, expected);
        public string CreatePage(string section)
        {
            if (FailCreate) throw new Exception("create failed");
            // 本机实测的新建空白页：PageTitle 和正文样式定义，标题 OE 引用 PageTitle，正文为空。
            var id = "created-" + ++_ids;
            XElement Style(string index, string name, string font, string size) => new XElement(One + "QuickStyleDef", new XAttribute("index", index), new XAttribute("name", name),
                new XAttribute("fontColor", "automatic"), new XAttribute("highlightColor", "automatic"), new XAttribute("font", font), new XAttribute("fontSize", size),
                new XAttribute("spaceBefore", "0.0"), new XAttribute("spaceAfter", "0.0"));
            var page = new XElement(One + "Page", new XAttribute("name", ""), new XAttribute("lastModifiedTime", "2026-09-26T00:00:00Z"),
                Style("0", "PageTitle", "Microsoft YaHei Light", "20.0"), Style("1", "p", "Microsoft YaHei", "11.0"),
                new XElement(One + "Title", new XElement(One + "OE", new XAttribute("objectID", "title-" + id), new XAttribute("quickStyleIndex", "0"), new XElement(One + "T", new XCData("")))));
            Add(id, page, section); Creates++; AfterCreate?.Invoke(); return id;
        }
        public string CreateSection(string parent, string name)
        { var id = "section-" + ++_ids; Node(parent).Add(new XElement(One + "Section", new XAttribute("ID", id), new XAttribute("name", name))); Creates++; return id; }
        public void Recycle(string id, DateTime expected)
        { if (FailRecycle) throw new Exception("recycle failed"); Node(id).Remove(); Pages.Remove(id); Recycles++; AfterRecycle?.Invoke(); if (RecycleThrowsAfterApplied) throw new Exception("recycle response lost"); }
        public string GetLink(string id, string obj) => "onenote:test#" + id + (obj == null ? "" : "&object=" + obj);
        public void Navigate(string id, string obj) { Node(id); Current = id; Navigations++; }
        public void ExportPdf(string id, string path) { if (FailExport) throw new Exception("export failed"); File.WriteAllText(path, "%PDF-1.4\nsynthetic offline output"); }
    }
    private static McpEditService WorkspaceService(FakeWorkspace api, Func<DateTime> now = null, Func<AgentOptions> options = null)
    { var service = new McpEditService(api, () => api.Current, options ?? (() => new AgentOptions()), () => new AddInSettings(), now); service.BindClient(McpClientA, McpSecretA); return service; }
    private static string WorkspaceBegin(McpEditService service) => (string)McpCall(service, "begin_workspace_edit", new { })["snapshot_id"];
    private static void McpReadAll(McpEditService service, string id)
    {
        var overview = McpCall(service, "get_page_overview", new { snapshot_id = id });
        foreach (var block in ((IList)overview["blocks"]).Cast<object>().Select(Map).Where(b => b.TryGetValue("editable", out var value) && (bool)value)) McpRead(service, id, (string)block["id"]);
    }
    private static void TestMcpExtensions(string executable)
    {
        TestMcpObjectMarks();
        TestMcpWriteVerification();
        Test("MCP readonly tree search stable pagination selection code and cursor isolation", () =>
        {
            var api = new FakeWorkspace(); api.Add("page", Page(Paragraph("a", "中文 <b>重点</b> <a href='https://example.com'>来源</a>"), Paragraph("b", "第二段")));
            var table = CodeBlockBuilder.BuildTable("  int x = 1;\n\n  x++;", OneNoteCodeHelper.Highlighting.LanguageRegistry.Find("csharp"), OneNoteCodeHelper.Highlighting.Themes.CodeThemes.Light, new AddInSettings());
            api.Pages["page"].Page.Elements(One + "Outline").First().Element(One + "OEChildren").Add(new XElement(One + "OE", table));
            var now = DateTime.UtcNow;
            using (var s = WorkspaceService(api, () => now))
            {
                Equal(1, ((IList)McpCall(s, "list_nodes", new { })["items"]).Count);
                Equal(2, ((IList)McpCall(s, "list_nodes", new { parent_id = "book" })["items"]).Count);
                var group = Map(((IList)McpCall(s, "list_nodes", new { parent_id = "group" })["items"])[0]); Equal("笔记本 / 资料 / 归档", group["path"]);
                var read = McpCall(s, "read_page", new { page_id = "page", limit = 1 }); var cursor = (string)read["next_cursor"];
                AgentCommitter.Find(api.Pages["page"].Page, "b").Element(One + "T").Value = "后来编辑";
                var next = McpCall(s, "read_page", new { page_id = "page", limit = 100, cursor });
                Equal("第二段", Map(((IList)next["items"])[0])["text"]); True(((IList)next["items"]).Cast<object>().Select(Map).Any(b => (string)b["type"] == "code"));
                s.BindClient("bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb", "22222222222222222222222222222222");
                Throws<McpFault>(() => McpCall(s, "read_page", new { page_id = "page", cursor }, "bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb"));
                now = now.AddMinutes(6); Throws<McpFault>(() => McpCall(s, "read_page", new { page_id = "page", cursor }));
                Throws<McpFault>(() => McpCall(s, "read_selection", new { }));
                AgentCommitter.Find(api.Pages["page"].Page, "a").SetAttributeValue("selected", "all");
                Equal(1, ((IList)McpCall(s, "read_selection", new { })["items"]).Count);
                Equal(1, ((IList)McpCall(s, "search_pages", new { query = "中文" })["items"]).Count); Equal(0, api.Navigations);
                api.FailSearch = true; Throws<McpFault>(() => McpCall(s, "search_pages", new { query = "中文" }));
            }
        });
        Test("MCP append empty page idempotent draft commits and whole outline undo", () =>
        {
            var api = new FakeWorkspace(); var blank = Page(); blank.Elements(One + "Outline").Remove(); api.Add("page", blank);
            using (var s = WorkspaceService(api))
            {
                var id = McpBegin(s); var args = new { snapshot_id = id, content = "  中文😀\n\n末行", format = "plain" };
                McpCall(s, "append_content", args); Equal(0, api.Pages["page"].Writes); Equal(1, McpRevision(s, id));
                McpCall(s, "append_content", args); Equal(1, McpRevision(s, id));
                Equal("Verified", McpFinish(s, id, 1)["status"]); True(Texts(api.Pages["page"].Page).Contains("  中文😀"));
                Equal("Verified", McpCall(s, "undo_edit", new { snapshot_id = id })["status"]); Equal(0, api.Pages["page"].Page.Elements(One + "Outline").Count());
            }
        });
        Test("MCP read quota is independent and table objects cannot silently disappear", () =>
        {
            var api = new FakeWorkspace(); var rows = Enumerable.Range(0, 1005).Select(i => Paragraph("p" + i, "中文段落")); api.Add("large", Page(rows.ToArray()));
            using (var s = WorkspaceService(api, options: () => new AgentOptions { MaxPageChars = 1000 }))
            {
                Equal(1005, McpCall(s, "read_page", new { page_id = "large" })["total"]); Throws<AiException>(() => McpCall(s, "begin_edit", new { page_id = "large", scope = "page" }));
                var table = new XElement(One + "Table", new XElement(One + "Row", new XElement(One + "Cell", new XElement(One + "OEChildren", new XElement(One + "OE", new XAttribute("objectID", "image-oe"), new XElement(One + "Image", new XElement(One + "OCRData", new XElement(One + "OCRText", "识别文字"))))))));
                api.Add("table", Page(new XElement(One + "OE", table), new XElement(One + "OE", new XElement(One + "InsertedFile"))));
                var result = McpCall(s, "read_page", new { page_id = "table" }); True(!(bool)result["complete"]); True(((IList)result["issues"]).Cast<string>().Contains("image_placeholder")); True(((IList)result["issues"]).Cast<string>().Contains("unsupported_object:InsertedFile"));
                True(((string)McpCall(s, "read_page", new { page_id = "table", format = "markdown" })["markdown"]).Contains("识别文字"));
            }
        });
        Test("MCP search reads one batch and never imports later candidate edits", () =>
        {
            var api = new FakeWorkspace(); api.Add("first", Page(Paragraph("a", "同名中文一"))); api.Add("second", Page(Paragraph("b", "同名中文二")));
            using (var s = WorkspaceService(api))
            {
                var found = McpCall(s, "search_pages", new { query = "中文", limit = 1 }); Equal(1, api.Reads); Equal(2, found["total"]);
                api.Pages["second"].Page.SetAttributeValue("lastModifiedTime", DateTime.UtcNow.AddMinutes(1).ToString("o"));
                var next = McpCall(s, "search_pages", new { query = "中文", limit = 1, cursor = (string)found["next_cursor"] }); Equal("page_changed", Map(((IList)next["items"])[0])["reason"]); True(!(bool)next["complete"]); Equal(2, api.Reads);
                McpCall(s, "search_pages", new { query = "中文", limit = 1, cursor = (string)found["next_cursor"] }); Equal(2, api.Reads);
            }
        });
        Test("MCP imported Markdown styles lists todo links code blanks tables and conflict undo", () =>
        {
            var api = new FakeWorkspace(); api.Add("page", Page(Paragraph("a", "原文")));
            using (var s = WorkspaceService(api))
            {
                var id = McpBegin(s); McpReadAll(s, id);
                McpCall(s, "insert_content", new { snapshot_id = id, target_id = "p1", position = "after", format = "markdown", content = "# 标题\n- [ ] **任务**\n  - 子项\n[来源](https://example.com)\n```csharp\n  int x = 1;\n\n```\n| 名称 | 数量 |\n| --- | --- |\n| A | 2 |" });
                Equal("Verified", McpFinish(s, id, McpRevision(s, id))["status"]);
                var page = api.Pages["page"].Page; Equal(2, page.Descendants(One + "Table").Count()); True(page.Descendants(One + "Tag").Any()); True(Texts(page).Contains("任务"));
                Equal("Verified", McpCall(s, "undo_edit", new { snapshot_id = id })["status"]); Equal("原文", Texts(api.Pages["page"].Page));
                var append = McpBegin(s); McpReadAll(s, append); McpCall(s, "append_content", new { snapshot_id = append, content = "新增", format = "plain" });
                Equal("Verified", McpFinish(s, append, McpRevision(s, append))["status"]);
                api.Pages["page"].Page.Elements(One + "Outline").Last().Descendants(One + "T").First().Value = "用户后续修改";
                Equal("NoChange", McpCall(s, "undo_edit", new { snapshot_id = append })["status"]); True(Texts(api.Pages["page"].Page).Contains("用户后续修改"));
            }
        });
        Test("MCP content and built-in formatting rewrite text boxes that already have to-dos", () =>
        {
            var tagged = Paragraph("b", "已有待办"); tagged.AddFirst(Tag("0"));
            var page = Page(Paragraph("a", "原文"), tagged); page.AddFirst(TagDef("0", 3, "待办事项"));
            var api = new FakeWorkspace(); api.Add("page", page);
            using (var s = WorkspaceService(api))
            {
                var style = McpBegin(s); McpReadAll(s, style); McpStyle(s, style);
                Equal("Verified", McpFinish(s, style, McpRevision(s, style))["status"]);
                var insert = McpBegin(s); McpReadAll(s, insert);
                McpCall(s, "insert_content", new { snapshot_id = insert, target_id = "p1", position = "after", format = "markdown", content = "- [ ] 新待办\n1. 第一\n2. 第二" });
                Equal("Verified", McpFinish(s, insert, McpRevision(s, insert))["status"]); True(Texts(api.Pages["page"].Page).Contains("第二"));
                Equal("Verified", McpCall(s, "undo_edit", new { snapshot_id = insert })["status"]);
                var append = McpBegin(s); McpReadAll(s, append);
                McpCall(s, "append_content", new { snapshot_id = append, format = "markdown", content = "- [ ] 追加待办\n1. 第一" });
                Equal("Verified", McpFinish(s, append, McpRevision(s, append))["status"]);
                Equal("Verified", McpCall(s, "undo_edit", new { snapshot_id = append })["status"]); Equal(1, api.Pages["page"].Page.Elements(One + "Outline").Count());
            }
        });
        Test("MCP replace text preserves rich runs and rejects concurrent text or unsupported targets", () =>
        {
            var api = new FakeWorkspace(); api.Add("page", Page(Paragraph("a", "<b>旧结论</b> <a href='https://example.com'>链接</a>")));
            using (var s = WorkspaceService(api))
            {
                var id = McpBegin(s); McpRead(s, id, "p1");
                Equal("a", Map(((IList)McpCall(s, "get_page_overview", new { snapshot_id = id })["blocks"])[0])["object_id"]);
                var args = new { snapshot_id = id, block_id = "p1", quote = "旧结论", occurrence = 1, replacement = "新结论" };
                McpCall(s, "replace_text", args); McpCall(s, "replace_text", args); Equal(1, McpRevision(s, id));
                Equal("Verified", McpFinish(s, id, 1)["status"]); True(AgentCommitter.Find(api.Pages["page"].Page, "a").Element(One + "T").Value.Contains("<b>新结论</b>"));
                Equal("Verified", McpCall(s, "undo_edit", new { snapshot_id = id })["status"]);
                var conflict = McpBegin(s); McpRead(s, conflict, "p1"); McpCall(s, "replace_text", new { snapshot_id = conflict, block_id = "p1", quote = "旧结论", occurrence = 1, replacement = "待写入" });
                AgentCommitter.Find(api.Pages["page"].Page, "a").Element(One + "T").Value = "用户编辑";
                Equal("NoChange", McpFinish(s, conflict, 1)["status"]); Equal("用户编辑", Texts(api.Pages["page"].Page));
            }
        });
        Test("MCP explicit page edits stay bound and selection on another page rejected", () =>
        {
            var api = new FakeWorkspace(); api.Add("page", Page(Paragraph("a", "当前"))); api.Add("other", Page(Paragraph("b", "目标")));
            using (var s = WorkspaceService(api))
            {
                Throws<McpFault>(() => McpCall(s, "begin_edit", new { scope = "selection", page_id = "other" }));
                var id = (string)McpCall(s, "begin_edit", new { scope = "page", page_id = "other" })["snapshot_id"]; McpRead(s, id, "p1"); McpStyle(s, id);
                Equal("Verified", McpFinish(s, id, McpRevision(s, id))["status"]); Equal(0, api.Pages["page"].Writes); Equal(1, api.Pages["other"].Writes);
            }
        });
        Test("MCP exact occurrence deletion keeps other identical text styles and links", () =>
        {
            var api = new FakeWorkspace(); api.Add("page", Page(Paragraph("a", "<b>相同</b><a href='https://example.com'><i>相同</i></a>😀👩‍💻")));
            using (var s = WorkspaceService(api))
            {
                var id = McpBegin(s); McpRead(s, id, "p1");
                Throws<AiException>(() => McpCall(s, "replace_text", new { snapshot_id = id, block_id = "p1", quote = "👩", occurrence = 1, replacement = "字符" })); Equal(0, McpRevision(s, id));
                McpCall(s, "replace_text", new { snapshot_id = id, block_id = "p1", quote = "相同", occurrence = 1, replacement = "" });
                Equal("Verified", McpFinish(s, id, 1)["status"]); var html = AgentCommitter.Find(api.Pages["page"].Page, "a").Element(One + "T").Value;
                True(html.Contains("<i>相同</i>")); True(html.Contains("href=")); True(!html.Contains("<b>")); True(html.Contains("😀"));
                Equal("Verified", McpCall(s, "undo_edit", new { snapshot_id = id })["status"]);
            }
        });
        Test("MCP workspace creates page and section only on commit and repeats results", () =>
        {
            var api = new FakeWorkspace();
            using (var s = WorkspaceService(api))
            {
                var id = WorkspaceBegin(s); var args = new { snapshot_id = id, section_id = "s1", title = "会议😀", format = "markdown",
                    content = "# 结论\n普通 **粗体** [链接](https://example.com)\n- 项目\n  - 子项\n- [ ] 行动\n1. 第一\n```csharp\n  int x = 1;\n\n```\n| 名称 | 数量 |\n| --- | --- |\n| A | 2 |" };
                McpCall(s, "create_page", args); McpCall(s, "create_page", args); Equal(0, api.Creates);
                Throws<McpFault>(() => McpFinish(s, id, 0));
                var result = McpFinish(s, id, 1);
                Equal("Verified", result["status"]); Equal(1, api.Creates); True(result["link"] != null);
                // 标题沿用新建页的原生 PageTitle 样式，正文按内容对应样式定义。
                var created = api.Pages[(string)result["page_id"]].Page; var titleLine = created.Element(One + "Title").Element(One + "OE");
                Equal("会议😀", AgentCode.PlainText(titleLine)); Equal("会议😀", (string)created.Attribute("name"));
                Equal("PageTitle", (string)created.Elements(One + "QuickStyleDef").Single(d => (string)d.Attribute("index") == (string)titleLine.Attribute("quickStyleIndex")).Attribute("name"));
                True(Texts(created).Contains("子项")); Equal(2, created.Descendants(One + "Table").Count());
                Equal("Verified", McpFinish(s, id, 1)["status"]); Equal(1, api.Creates);
                Equal("Verified", McpCall(s, "undo_edit", new { snapshot_id = id })["status"]); Equal(1, api.Recycles);
                var section = WorkspaceBegin(s); McpCall(s, "create_section", new { snapshot_id = section, parent_id = "group", name = "新分区" });
                Equal("Verified", McpFinish(s, section, 1)["status"]); Equal("Verified", McpCall(s, "undo_edit", new { snapshot_id = section })["status"]);
                var abort = WorkspaceBegin(s); McpCall(s, "create_page", new { snapshot_id = abort, section_id = "s1", title = "取消" }); McpCall(s, "abort_edit", new { snapshot_id = abort });
                Throws<McpFault>(() => McpFinish(s, abort, 1)); Equal(2, api.Creates);
            }
        });
        Test("MCP workspace copies images and moves by verified copy then recycle with compensation", () =>
        {
            var api = new FakeWorkspace(); api.Add("page", Page(Paragraph("a", "<b>正文</b>"), new XElement(One + "OE", new XAttribute("objectID", "pic"), new XElement(One + "Image", new XElement(One + "CallbackID", new XAttribute("callbackID", "image")))))); api.Pages["page"].Binary["image"] = Convert.ToBase64String(new byte[] { 1, 2, 3 });
            using (var s = WorkspaceService(api))
            {
                var copy = WorkspaceBegin(s); McpCall(s, "copy_page", new { snapshot_id = copy, page_id = "page", section_id = "s2" });
                var copied = McpFinish(s, copy, 1); Equal("Verified", copied["status"]); True(api.Pages.ContainsKey("page")); True(((IList)copied["object_id_map"]).Count > 0);
                Equal("Verified", McpCall(s, "undo_edit", new { snapshot_id = copy })["status"]);
                var move = WorkspaceBegin(s); McpCall(s, "move_page", new { snapshot_id = move, page_id = "page", section_id = "s2" });
                var moved = McpFinish(s, move, 1); Equal("Verified", moved["status"]); True(!api.Pages.ContainsKey("page")); Equal(true, moved["source_recycled"]);
                var undo = McpCall(s, "undo_edit", new { snapshot_id = move }); Equal("Verified", undo["status"]); True(api.Pages.ContainsKey((string)undo["restored_page_id"])); True(!api.Pages.ContainsKey((string)moved["page_id"]));
            }
        });
        Test("MCP workspace copies real page settings and verifies against OneNote-managed metadata and layout", () =>
        {
            var api = new FakeWorkspace(); var page = Page(Paragraph("a", "<b>正文</b>"), Listed("b", "列表", "2"));
            page.AddFirst(TagDef("0", 3, "待办事项")); AgentCommitter.Find(page, "a").AddFirst(Tag("0"));
            var code = CodeBlockBuilder.BuildTable("  int x = 1;", OneNoteCodeHelper.Highlighting.LanguageRegistry.Find("csharp"), OneNoteCodeHelper.Highlighting.Themes.CodeThemes.Light, new AddInSettings());
            page.Element(One + "Outline").Element(One + "OEChildren").Add(new XElement(One + "OE", code));
            page.Element(One + "Outline").AddBeforeSelf(new XElement(One + "PageSettings", new XAttribute("RTL", "false"), new XAttribute("color", "#FFFFE0"),
                new XElement(One + "PageSize", new XElement(One + "Automatic")), new XElement(One + "RuleLines", new XAttribute("visible", "true"))),
                new XElement(One + "Title", new XElement(One + "OE", new XAttribute("objectID", "title"), new XElement(One + "T", new XCData("源标题")))));
            api.Add("page", page);
            using (var s = WorkspaceService(api))
            {
                var copy = WorkspaceBegin(s); McpCall(s, "copy_page", new { snapshot_id = copy, page_id = "page", section_id = "s2" });
                var copied = McpFinish(s, copy, 1); Equal("Verified", copied["status"]);
                var target = api.Pages[(string)copied["page_id"]].Page;
                Equal("true", (string)target.Element(One + "PageSettings").Element(One + "RuleLines").Attribute("visible")); Equal(Texts(api.Pages["page"].Page), Texts(target));
                True(target.Descendants(One + "Tag").Any()); Equal(1, target.Descendants(One + "Table").Count()); Equal("源标题", AgentPageTitle.Text(target.Element(One + "Title")));
                Equal("Verified", McpCall(s, "undo_edit", new { snapshot_id = copy })["status"]);
                var move = WorkspaceBegin(s); McpCall(s, "move_page", new { snapshot_id = move, page_id = "page", section_id = "s2" });
                // 用户查看源页只改变阅读状态，不算源页变化。
                api.Pages["page"].Page.SetAttributeValue("isUnread", "true");
                var moved = McpFinish(s, move, 1); Equal("Verified", moved["status"]); Equal(true, moved["source_recycled"]);
                Equal("Verified", McpCall(s, "undo_edit", new { snapshot_id = move })["status"]);
            }
        });
        Test("MCP workspace refuses complex pages children changes and preserves source on failures", () =>
        {
            var api = new FakeWorkspace(); api.Add("page", Page(Paragraph("a", "源页")));
            using (var s = WorkspaceService(api))
            {
                var changed = WorkspaceBegin(s); McpCall(s, "move_page", new { snapshot_id = changed, page_id = "page", section_id = "s2" });
                AgentCommitter.Find(api.Pages["page"].Page, "a").Element(One + "T").Value = "用户编辑";
                Equal("NoChange", McpFinish(s, changed, 1)["status"]); Equal(0, api.Creates);
                api.Pages["page"].Page.Add(new XElement(One + "InkDrawing"));
                var complex = WorkspaceBegin(s); Throws<McpFault>(() => McpCall(s, "copy_page", new { snapshot_id = complex, page_id = "page", section_id = "s2" })); Equal(0, McpRevision(s, complex));
                api.Pages["page"].Page.Elements(One + "InkDrawing").Remove();
                api.Hierarchy.Descendants(One + "Page").First().AddAfterSelf(new XElement(One + "Page", new XAttribute("ID", "child"), new XAttribute("name", "子页"), new XAttribute("pageLevel", 2)));
                Throws<McpFault>(() => McpCall(s, "move_page", new { snapshot_id = complex, page_id = "page", section_id = "s2" })); api.Hierarchy.Descendants(One + "Page").Single(e => (string)e.Attribute("ID") == "child").Remove();
                var move = WorkspaceBegin(s); McpCall(s, "move_page", new { snapshot_id = move, page_id = "page", section_id = "s2" });
                api.AfterUpdate = () => AgentCommitter.Find(api.Pages["page"].Page, "a").Element(One + "T").Value = "提交期间编辑";
                Equal("PartiallyApplied", McpFinish(s, move, 1)["status"]); True(api.Pages.ContainsKey("page")); Equal(0, api.Recycles);
                api.AfterUpdate = null; var failure = WorkspaceBegin(s); McpCall(s, "move_page", new { snapshot_id = failure, page_id = "page", section_id = "s2" }); api.FailRecycle = true;
                Equal("PartiallyApplied", McpFinish(s, failure, 1)["status"]); var count = api.Creates; McpFinish(s, failure, 1); Equal(count, api.Creates); True(api.Pages.ContainsKey("page"));
            }
        });
        Test("MCP workspace lost create or recycle responses never replay or delete the sole copy", () =>
        {
            var api = new FakeWorkspace(); api.Add("page", Page(Paragraph("a", "源正文")));
            using (var s = WorkspaceService(api))
            {
                var fail = WorkspaceBegin(s); McpCall(s, "create_page", new { snapshot_id = fail, section_id = "s1", title = "失败" }); api.FailCreate = true;
                Equal("CommitOutcomeUnknown", McpFinish(s, fail, 1)["status"]); api.FailCreate = false; McpFinish(s, fail, 1); Equal(0, api.Creates);
                var cancelled = WorkspaceBegin(s); McpCall(s, "create_page", new { snapshot_id = cancelled, section_id = "s1", title = "在途取消" });
                api.AfterCreate = () => McpCall(s, "abort_edit", new { snapshot_id = cancelled });
                Equal("Verified", McpFinish(s, cancelled, 1)["status"]); api.AfterCreate = null;
                var move = WorkspaceBegin(s); McpCall(s, "move_page", new { snapshot_id = move, page_id = "page", section_id = "s2" }); api.RecycleThrowsAfterApplied = true;
                Equal("Verified", McpFinish(s, move, 1)["status"]); True(!api.Pages.ContainsKey("page")); api.RecycleThrowsAfterApplied = false;
                Equal("Verified", McpCall(s, "undo_edit", new { snapshot_id = move })["status"]); True(api.Pages.Values.Any(p => Texts(p.Page).Contains("源正文")));
                var source = api.Pages.First(p => Texts(p.Value.Page).Contains("源正文")).Key;
                var unknown = WorkspaceBegin(s); McpCall(s, "move_page", new { snapshot_id = unknown, page_id = source, section_id = "s2" });
                api.AfterRecycle = () => api.FailHierarchy = true;
                Equal("CommitOutcomeUnknown", McpFinish(s, unknown, 1)["status"]); True(!(bool)McpCall(s, "get_edit_status", new { snapshot_id = unknown })["can_undo"]);
                api.FailHierarchy = false; api.AfterRecycle = null; Throws<McpFault>(() => McpCall(s, "undo_edit", new { snapshot_id = unknown })); True(!api.Pages.ContainsKey(source));
            }
        });
        Test("MCP hierarchy changes protect children before move and before copy undo", () =>
        {
            var api = new FakeWorkspace(); api.Add("page", Page(Paragraph("a", "有子页的父页")));
            using (var s = WorkspaceService(api))
            {
                var move = WorkspaceBegin(s); McpCall(s, "move_page", new { snapshot_id = move, page_id = "page", section_id = "s2" });
                var source = api.Hierarchy.Descendants(One + "Page").Single(e => (string)e.Attribute("ID") == "page");
                source.AddAfterSelf(new XElement(One + "Page", new XAttribute("ID", "new-child"), new XAttribute("pageLevel", 2)));
                Equal("NoChange", McpFinish(s, move, 1)["status"]); Equal(0, api.Creates);
                var copy = WorkspaceBegin(s); McpCall(s, "copy_page", new { snapshot_id = copy, page_id = "page", section_id = "s2" });
                var copied = McpFinish(s, copy, 1); Equal("Verified", copied["status"]);
                api.Hierarchy.Descendants(One + "Page").Single(e => (string)e.Attribute("ID") == (string)copied["page_id"]).AddAfterSelf(new XElement(One + "Page", new XAttribute("ID", "target-child"), new XAttribute("pageLevel", 2)));
                Equal("NoChange", McpCall(s, "undo_edit", new { snapshot_id = copy })["status"]); Equal(0, api.Recycles);
            }
        });
        Test("MCP tasks links Markdown and PDF export refuse overwrite and preserve real files", () =>
        {
            var api = new FakeWorkspace(); var page = Page(Paragraph("a", "待办"), Paragraph("b", "完成")); page.AddFirst(TagDef("0", 3, "待办事项"));
            AgentCommitter.Find(page, "a").AddFirst(Tag("0")); AgentCommitter.Find(page, "b").AddFirst(Tag("0")); AgentCommitter.Find(page, "b").Element(One + "Tag").SetAttributeValue("completed", "true"); api.Add("page", page);
            var folder = Path.Combine(Path.GetTempPath(), "onenote-mcp-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
            try
            {
                using (var s = WorkspaceService(api))
                {
                    Equal(1, ((IList)McpCall(s, "find_tasks", new { })["items"]).Count); Equal(1, ((IList)McpCall(s, "find_tasks", new { completed = true })["items"]).Count);
                    api.Add("locked", Page(Paragraph("c", "无法读"))); api.Unreadable.Add("locked"); Equal(1, ((IList)McpCall(s, "find_tasks", new { })["skipped"]).Count);
                    using (var cancelled = new CancellationTokenSource()) { cancelled.Cancel(); Throws<OperationCanceledException>(() => s.Call(McpClientA, "find_tasks", "{}", cancelled.Token)); }
                    True(((string)McpCall(s, "get_link", new { page_id = "page", object_id = "a" })["link"]).Contains("object=a"));
                    Throws<McpFault>(() => McpCall(s, "get_link", new { page_id = "page", object_id = "p1" }));
                    McpCall(s, "navigate_to", new { page_id = "page" }); Equal(1, api.Navigations);
                    var path = Path.Combine(folder, "中文.md"); McpCall(s, "export_page", new { page_id = "page", format = "markdown", output_path = path }); True(File.ReadAllText(path).Contains("- [ ] 待办"));
                    Throws<McpFault>(() => McpCall(s, "export_page", new { page_id = "page", format = "markdown", output_path = path }));
                    Throws<McpFault>(() => McpCall(s, "export_page", new { page_id = "page", format = "markdown", output_path = "C:relative.md" }));
                    var pdf = Path.Combine(folder, "中文.pdf"); McpCall(s, "export_page", new { page_id = "page", format = "pdf", output_path = pdf }); True(File.ReadAllText(pdf).StartsWith("%PDF"));
                    api.FailExport = true; Throws<Exception>(() => McpCall(s, "export_page", new { page_id = "page", format = "pdf", output_path = Path.Combine(folder, "失败.pdf") })); Equal(2, Directory.GetFiles(folder).Length);
                }
            }
            finally { foreach (var f in Directory.GetFiles(folder)) File.Delete(f); Directory.Delete(folder); }
        });
        Test("real MCP workspace tools expose output schemas and create repeat undo through stdio", () =>
        {
            var api = new FakeWorkspace(); var pipe = "OneNoteCodeHelper.ExtensionTest." + Guid.NewGuid().ToString("N");
            using (var s = WorkspaceService(api)) using (var host = new McpPipeHost(s, pipe))
            {
                host.Start(); True(host.Ready.Wait(5000)); using (var child = new McpChild(executable, pipe))
                {
                    child.Initialize(); var catalog = child.Rpc("tools/list", new { }); True(((IList)catalog["tools"]).Cast<object>().Select(Map).Single(t => (string)t["name"] == "read_page").ContainsKey("outputSchema"));
                    child.Rpc("tools/call", new { name = "list_nodes", arguments = new { }, _meta = new { progressToken = "offline-progress" } }); True(child.Notifications.Any(n => (string)n["method"] == "notifications/progress"));
                    var id = (string)child.Call("begin_workspace_edit", new { })["snapshot_id"];
                    child.Call("create_page", new { snapshot_id = id, section_id = "s1", title = "协议中文" });
                    Equal("Verified", child.Call("finish_edit", new { snapshot_id = id, draft_revision = 1 })["status"]); Equal(1, api.Creates);
                    Equal("Verified", child.Call("finish_edit", new { snapshot_id = id, draft_revision = 1 })["status"]); Equal(1, api.Creates);
                    Equal("Verified", child.Call("undo_edit", new { snapshot_id = id })["status"]);
                }
            }
        });
    }

    private static XElement McpMarkedObjectPage(string kind)
    {
        var oe = new XElement(One + "OE", new XAttribute("objectID", "marked-object"), Tag("0"),
            new XElement(One + "List", new XElement(One + "Number", new XAttribute("numberSequence", "0"),
                new XAttribute("numberFormat", "##."), new XAttribute("restartNumberingAt", "1"))));
        if (kind == "image") oe.Add(new XElement(One + "Image", new XElement(One + "CallbackID", new XAttribute("callbackID", "marked-image"))));
        else if (kind == "table") oe.Add(new XElement(One + "Table", new XAttribute("objectID", "marked-table"),
            new XElement(One + "Columns", new XElement(One + "Column", new XAttribute("index", "0"), new XAttribute("width", "120"))),
            new XElement(One + "Row", new XElement(One + "Cell", new XElement(One + "OEChildren", Paragraph("cell", "单元格"))))));
        var page = Page(oe, Paragraph("body", "正文"));
        page.AddFirst(TagDef("0", 3, "待办事项"), TagDef("1", 13, "重要"));
        return page;
    }
    private static XElement McpNonTextObject(XElement page) => page.Descendants(One + "OE").Single(e => !e.Elements(One + "T").Any());
    private static FakeWorkspace McpMarkedObjectWorkspace(string kind)
    {
        var api = new FakeWorkspace(); api.Add("page", McpMarkedObjectPage(kind));
        if (kind == "image") api.Pages["page"].Binary["marked-image"] = "AQID";
        return api;
    }
    private static void TestMcpObjectMarks()
    {
        var kinds = new[] { "image", "table", "empty" };
        Test("MCP write verification detects marks and lists changing on every no-text object", () =>
        {
            var changes = new Dictionary<string, Action<XElement>>
            {
                ["removed tag"] = oe => oe.Elements(One + "Tag").Remove(),
                ["added tag"] = oe => oe.AddFirst(Tag("1")),
                ["completed tag"] = oe => oe.Element(One + "Tag").SetAttributeValue("completed", "true"),
                ["tag icon"] = oe => oe.Element(One + "Tag").SetAttributeValue("index", "1"),
                ["removed list"] = oe => oe.Elements(One + "List").Remove(),
                ["list kind"] = oe => oe.Element(One + "List").ReplaceWith(new XElement(One + "List", new XElement(One + "Bullet", new XAttribute("bullet", "2")))),
                ["number sequence"] = oe => oe.Element(One + "List").Element(One + "Number").SetAttributeValue("numberSequence", "2"),
                ["number format"] = oe => oe.Element(One + "List").Element(One + "Number").SetAttributeValue("numberFormat", "##)"),
                ["number start"] = oe => oe.Element(One + "List").Element(One + "Number").SetAttributeValue("restartNumberingAt", "4")
            };
            foreach (var kind in kinds)
            {
                var page = McpMarkedObjectPage(kind); var expected = McpPageModel.WriteSignature(page);
                foreach (var change in changes)
                {
                    var actual = new XElement(page); change.Value(McpNonTextObject(actual));
                    if (expected == McpPageModel.WriteSignature(actual)) throw new Exception(kind + ": failed to detect " + change.Key);
                }
                var unmarked = new XElement(page); var oe = McpNonTextObject(unmarked);
                oe.Elements(One + "Tag").Remove(); oe.Elements(One + "List").Remove();
                var tagged = new XElement(unmarked); McpNonTextObject(tagged).AddFirst(Tag("0"));
                True(McpPageModel.WriteSignature(unmarked) != McpPageModel.WriteSignature(tagged));
                var listed = new XElement(unmarked); McpNonTextObject(listed).AddFirst(new XElement(One + "List", new XElement(One + "Bullet", new XAttribute("bullet", "2"))));
                True(McpPageModel.WriteSignature(unmarked) != McpPageModel.WriteSignature(listed));
            }
        });
        Test("MCP no-text marks tolerate definition renumbering timestamps and default list starts", () =>
        {
            foreach (var kind in kinds)
            {
                var expected = McpMarkedObjectPage(kind); var actual = new XElement(expected); var oe = McpNonTextObject(actual);
                actual.Elements(One + "TagDef").Single(d => (string)d.Attribute("index") == "0").SetAttributeValue("index", "40");
                oe.Element(One + "Tag").SetAttributeValue("index", "40");
                oe.Element(One + "Tag").SetAttributeValue("creationDate", "2026-10-03T01:00:00Z");
                var number = oe.Element(One + "List").Element(One + "Number");
                number.SetAttributeValue("restartNumberingAt", null); number.SetAttributeValue("numberSequence", "00"); number.SetAttributeValue("fontSize", "11.0");
                actual.SetAttributeValue("dateTime", "2026-10-03T01:00:00Z"); actual.SetAttributeValue("pageLevel", "1");
                oe.SetAttributeValue("author", "测试者"); oe.SetAttributeValue("alignment", "left");
                Equal(McpPageModel.WriteSignature(expected), McpPageModel.WriteSignature(actual));
            }
        });
        Test("MCP copy move and compensation preserve image table and empty-object marks", () =>
        {
            foreach (var kind in kinds)
            {
                var api = McpMarkedObjectWorkspace(kind);
                using (var s = WorkspaceService(api))
                {
                    var copy = WorkspaceBegin(s); McpCall(s, "copy_page", new { snapshot_id = copy, page_id = "page", section_id = "s2" });
                    var copied = McpFinish(s, copy, 1); Equal("Verified", copied["status"]);
                    CheckMarks((string)copied["page_id"]); True(api.Pages.ContainsKey("page"));
                    Equal("Verified", McpCall(s, "undo_edit", new { snapshot_id = copy })["status"]);
                    True(!api.Pages.ContainsKey((string)copied["page_id"]));
                    var move = WorkspaceBegin(s); McpCall(s, "move_page", new { snapshot_id = move, page_id = "page", section_id = "s2" });
                    var moved = McpFinish(s, move, 1); Equal("Verified", moved["status"]); Equal(true, moved["source_recycled"]);
                    CheckMarks((string)moved["page_id"]); True(!api.Pages.ContainsKey("page"));
                    var undo = McpCall(s, "undo_edit", new { snapshot_id = move }); Equal("Verified", undo["status"]);
                    CheckMarks((string)undo["restored_page_id"]); True(!api.Pages.ContainsKey((string)moved["page_id"]));
                }
                void CheckMarks(string id)
                {
                    var page = XElement.Parse(api.GetPageContent(id, PageInfo.piBinaryData)); var oe = McpNonTextObject(page);
                    Equal(1, oe.Elements(One + "Tag").Count()); Equal("false", (string)oe.Element(One + "Tag").Attribute("completed"));
                    Equal("3", (string)AgentMarks.Definition(page, oe.Element(One + "Tag")).Attribute("symbol"));
                    Equal("0", (string)oe.Element(One + "List").Element(One + "Number").Attribute("numberSequence"));
                    Equal("##.", (string)oe.Element(One + "List").Element(One + "Number").Attribute("numberFormat"));
                    if (kind == "image") Equal("AQID", page.Descendants(One + "Image").Single().Element(One + "Data").Value);
                }
            }
        });
        Test("MCP lost no-text tags freeze copy and move without recycling or replaying", () =>
        {
            foreach (var kind in kinds) foreach (var operation in new[] { "copy_page", "move_page" })
            {
                var api = McpMarkedObjectWorkspace(kind);
                using (var s = WorkspaceService(api))
                {
                    var id = WorkspaceBegin(s); McpCall(s, operation, new { snapshot_id = id, page_id = "page", section_id = "s2" });
                    api.AfterUpdate = () => McpNonTextObject(api.Pages.Single(p => p.Key != "page").Value.Page).Elements(One + "Tag").Remove();
                    var result = McpFinish(s, id, 1); Equal("CommitOutcomeUnknown", result["status"]); Equal(false, result["source_recycled"]);
                    Equal(0, api.Recycles); True(api.Pages.ContainsKey("page")); True(api.Pages.ContainsKey((string)result["page_id"]));
                    Equal(1, McpNonTextObject(api.Pages["page"].Page).Elements(One + "Tag").Count());
                    Equal(0, McpNonTextObject(api.Pages[(string)result["page_id"]].Page).Elements(One + "Tag").Count());
                    var status = McpCall(s, "get_edit_status", new { snapshot_id = id }); Equal("CommitOutcomeUnknown", status["state"]); Equal(false, status["can_undo"]);
                    var writes = api.Pages.Values.Sum(p => p.Writes);
                    Equal(McpJson.Serialize(result), McpJson.Serialize(McpFinish(s, id, 1))); Equal(1, api.Creates); Equal(writes, api.Pages.Values.Sum(p => p.Writes));
                    Throws<McpFault>(() => McpCall(s, operation, new { snapshot_id = id, page_id = "page", section_id = "s2" }));
                    Throws<McpFault>(() => McpCall(s, "undo_edit", new { snapshot_id = id }));
                }
            }
        });
        Test("MCP move compensation keeps the target when restored no-text marks are missing", () =>
        {
            foreach (var kind in kinds)
            {
                var api = McpMarkedObjectWorkspace(kind);
                using (var s = WorkspaceService(api))
                {
                    var id = WorkspaceBegin(s); McpCall(s, "move_page", new { snapshot_id = id, page_id = "page", section_id = "s2" });
                    var moved = McpFinish(s, id, 1); Equal("Verified", moved["status"]); var target = (string)moved["page_id"];
                    api.AfterUpdate = () => McpNonTextObject(api.Pages.Single(p => p.Key != target).Value.Page).Elements(One + "Tag").Remove();
                    var undo = McpCall(s, "undo_edit", new { snapshot_id = id }); Equal("CommitOutcomeUnknown", undo["status"]);
                    True(api.Pages.ContainsKey(target)); True(api.Pages.ContainsKey((string)undo["restored_page_id"])); Equal(1, api.Recycles);
                    Equal(1, McpNonTextObject(api.Pages[target].Page).Elements(One + "Tag").Count());
                    Equal(0, McpNonTextObject(api.Pages[(string)undo["restored_page_id"]].Page).Elements(One + "Tag").Count());
                    Equal(false, McpCall(s, "get_edit_status", new { snapshot_id = id })["can_undo"]);
                    var writes = api.Pages.Values.Sum(p => p.Writes);
                    Equal(McpJson.Serialize(undo), McpJson.Serialize(McpCall(s, "undo_edit", new { snapshot_id = id })));
                    Equal(2, api.Creates); Equal(1, api.Recycles); Equal(writes, api.Pages.Values.Sum(p => p.Writes));
                }
            }
        });
        Test("MCP move undo preserves later edits to no-text target marks", () =>
        {
            foreach (var kind in kinds)
            {
                var api = McpMarkedObjectWorkspace(kind);
                using (var s = WorkspaceService(api))
                {
                    var id = WorkspaceBegin(s); McpCall(s, "move_page", new { snapshot_id = id, page_id = "page", section_id = "s2" });
                    var moved = McpFinish(s, id, 1); Equal("Verified", moved["status"]); var target = (string)moved["page_id"];
                    McpNonTextObject(api.Pages[target].Page).Element(One + "Tag").SetAttributeValue("completed", "true");
                    Equal("NoChange", McpCall(s, "undo_edit", new { snapshot_id = id })["status"]);
                    True(api.Pages.ContainsKey(target)); Equal(1, api.Creates); Equal(1, api.Recycles);
                    Equal("true", (string)McpNonTextObject(api.Pages[target].Page).Element(One + "Tag").Attribute("completed"));
                }
            }
        });
    }
}
