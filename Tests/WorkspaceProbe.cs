using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using OneNoteCodeHelper.Services;
using OneNoteCodeHelper.Services.Agent;
using OneNoteCodeHelper.Services.Mcp;

/// <summary>
/// 显式运行才打开 OneNote：在指定目录新建独立的测试笔记本，核对工作区新建分区、新建页、复制、移动及撤销在真实 OneNote 上的回存核验。
/// 只读写这个测试笔记本，结束后关闭它，文件留在目录里备查。
/// </summary>
internal static class WorkspaceProbe
{
    private static XNamespace One => OneNoteApi.One;
    private const string Markdown = "# 结论\n普通 **粗体** [链接](https://example.com)\n- 项目\n  - 子项\n- [ ] 行动\n1. 第一\n```csharp\n  int x = 1;\n\n```\n| 名称 | 数量 |\n| --- | --- |\n| A | 2 |";
    private static int _failures;

    internal static int Run(string directory)
    {
        IApplication app = null; string notebook = null;
        try
        {
            directory = Path.GetFullPath(directory);
            Directory.CreateDirectory(directory);
            var folder = Path.Combine(directory, "MCP-workspace-probe-" + Guid.NewGuid().ToString("N"));
            // 新建笔记本时目录必须已存在，否则 OpenHierarchy 返回 0x80042006。
            Directory.CreateDirectory(folder);
            Console.WriteLine("Workspace probe notebook: " + folder);
            app = new ApplicationClass();
            app.OpenHierarchy(folder, "", out notebook, CreateFileType.cftNotebook);
            var api = new OneNoteApi(app);
            var read = new McpReadService(api, () => null, () => DateTime.UtcNow, () => new AgentOptions());
            var options = new AgentOptions { EnableInsert = true, EnableMoves = true, EnableLists = true, EnableTags = true, EnableTextTables = true, EnableCodeHighlight = true };
            AgentTools.UseInstalledFont(options);
            var settings = new AddInSettings();
            EnsureRecycleBin(app, api, notebook);

            var a = Commit(read, options, settings, "create_section", new { parent_id = notebook, name = "探针分区A" });
            var b = Commit(read, options, settings, "create_section", new { parent_id = notebook, name = "探针分区B" });
            var created = Commit(read, options, settings, "create_page", new { section_id = (string)a.Result["section_id"], title = "探针页面😀", format = "markdown", content = Markdown });
            var pageId = (string)created.Result["page_id"];
            Expect("created title", "探针页面😀", AgentPageTitle.Text(created.Draft.Written?.Element(One + "Title")));
            AddImage(api, pageId);

            var copy = Commit(read, options, settings, "copy_page", new { page_id = pageId, section_id = (string)b.Result["section_id"] });
            var source = read.ReadPage(pageId);
            Console.WriteLine("copy keeps page date: " + ((string)source.Attribute("dateTime") == (string)copy.Draft.Written?.Attribute("dateTime")));
            Expect("copy undo", "Verified", Status(copy.Draft.Undo(CancellationToken.None)));

            var move = Commit(read, options, settings, "move_page", new { page_id = pageId, section_id = (string)b.Result["section_id"] });
            Expect("move recycles source", "True", Convert.ToString(move.Result["source_recycled"]));
            var undo = McpJson.Arguments(McpJson.Serialize(move.Draft.Undo(CancellationToken.None)));
            var restored = (string)undo["restored_page_id"];
            if (!Expect("move undo", "Verified", (string)undo["status"]) && restored != null)
                Diff("move undo", move.Draft.Expected, read.ReadPage(restored, PageInfo.piBinaryData));

            if (restored != null)
            {
                // 页面编辑草稿里的导入工具走 Agent 的整框提交和核验。
                using (var service = new McpEditService(api, () => null, () => options, () => settings))
                {
                    const string client = "11111111111111111111111111111111";
                    service.BindClient(client, "22222222222222222222222222222222");
                    IDictionary<string, object> Call(string tool, object args) => McpJson.Arguments(service.Call(client, tool, McpJson.Serialize(args), CancellationToken.None));
                    foreach (var tool in new[] { "append_content", "insert_content" })
                    {
                        var id = (string)Call("begin_edit", new { scope = "page", page_id = restored })["snapshot_id"];
                        var blocks = ((System.Collections.IList)Call("get_page_overview", new { snapshot_id = id })["blocks"]).Cast<IDictionary<string, object>>().ToList();
                        var readBlocks = Call("read_blocks", new { snapshot_id = id, block_ids = blocks.Where(b => b.TryGetValue("editable", out var e) && (bool)e).Select(b => (string)b["id"]).ToArray() });
                        var target = ((System.Collections.IList)readBlocks["blocks"]).Cast<IDictionary<string, object>>().First(b => ((string)b["text"]).StartsWith("普通", StringComparison.Ordinal));
                        if (tool == "append_content") Call(tool, new { snapshot_id = id, content = Markdown, format = "markdown" });
                        else Call(tool, new { snapshot_id = id, target_id = (string)target["id"], position = "after", content = "插入段落\n- 项目\n  - 子项", format = "markdown" });
                        var revision = (int)Call("get_pending_changes", new { snapshot_id = id })["draft_revision"];
                        var finished = Call("finish_edit", new { snapshot_id = id, draft_revision = revision });
                        Expect(tool, "Verified", (string)finished["status"]);
                        if ((string)finished["status"] != "Verified") Console.WriteLine("     " + McpJson.Serialize(finished));
                        string undone;
                        try { undone = (string)Call("undo_edit", new { snapshot_id = id })["status"]; } catch (Exception ex) { undone = ex.Message; }
                        Expect(tool + " undo", "Verified", undone);
                    }
                    // 内置 Agent 的格式工具：文本框里已有待办，整框写回时要连同 TagDef 提交。
                    var sid = (string)Call("begin_edit", new { scope = "page", page_id = restored })["snapshot_id"];
                    var all = ((System.Collections.IList)Call("get_page_overview", new { snapshot_id = sid })["blocks"]).Cast<IDictionary<string, object>>().ToList();
                    var read2 = Call("read_blocks", new { snapshot_id = sid, block_ids = all.Where(b => b.TryGetValue("editable", out var e) && (bool)e).Select(b => (string)b["id"]).ToArray() });
                    var paragraph = ((System.Collections.IList)read2["blocks"]).Cast<IDictionary<string, object>>().First(b => ((string)b["text"]).StartsWith("普通", StringComparison.Ordinal));
                    Call("set_paragraph_style", new { snapshot_id = sid, block_ids = new[] { (string)paragraph["id"] }, preset_id = "heading2" });
                    Expect("built-in set_paragraph_style beside a to-do", "Verified", (string)Call("finish_edit", new { snapshot_id = sid, draft_revision = (int)Call("get_pending_changes", new { snapshot_id = sid })["draft_revision"] })["status"]);
                    Expect("built-in set_paragraph_style undo", "Verified", (string)Call("undo_edit", new { snapshot_id = sid })["status"]);
                }
                // AI 文字功能的写回：同样整框回传带待办的文本框。
                var aiPage = AgentPageSnapshot.ParsePage(api.GetPageContent(restored, PageInfo.piBasic));
                PageEditor.ReadAiTargets(aiPage, restored, null, null, out var targets);
                var original = targets.Paragraphs.First(x => x.Text.StartsWith("普通", StringComparison.Ordinal));
                var written = PageEditor.ApplyParagraphEdits(api, targets, new[] { new AiParagraphEdit(original, original.Text + "（AI）", new string[0]) }, false, out var applied, out _, out _);
                Expect("AI text write-back beside a to-do", "True", Convert.ToString(written.Success && applied.Count == 1 &&
                    AgentPageSnapshot.ParsePage(api.GetPageContent(restored, PageInfo.piBasic)).Descendants(One + "OE").Any(o => PageEditor.ExtractPlainText(o).EndsWith("（AI）", StringComparison.Ordinal))));
                var page = McpJson.Arguments(McpJson.Serialize(read.Call("probe", "read_page", Args(new { page_id = restored }), CancellationToken.None)));
                Console.WriteLine($"read_page restored: blocks={page["total"]} complete={page["complete"]}");
                var output = Path.Combine(directory, "MCP-workspace-probe-" + Guid.NewGuid().ToString("N") + ".md");
                read.Call("probe", "export_page", Args(new { page_id = restored, format = "markdown", output_path = output }), CancellationToken.None);
                Console.WriteLine("export_page markdown: " + output);
            }
            Expect("section undo", "Verified", Status(b.Draft.Undo(CancellationToken.None)));
            Console.WriteLine(_failures == 0 ? "Workspace probe: all checks passed." : $"Workspace probe: {_failures} check(s) failed.");
            return _failures == 0 ? 0 : 1;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally
        {
            if (app != null && notebook != null) try { app.CloseNotebook(notebook); } catch (Exception ex) { Console.Error.WriteLine("Close notebook failed: " + ex.Message); }
            if (app != null) Marshal.ReleaseComObject(app);
        }
    }

    private static (McpWorkspaceDraft Draft, IDictionary<string, object> Result) Commit(McpReadService read, AgentOptions options, AddInSettings settings, string kind, object args)
    {
        var draft = new McpWorkspaceDraft(read, options, settings);
        draft.Stage(kind, Args(args));
        var result = McpJson.Arguments(McpJson.Serialize(draft.Commit(CancellationToken.None)));
        if (!Expect(kind, "Verified", (string)result["status"])) Diff(kind, draft.Expected, draft.Written);
        return (draft, result);
    }

    private static IDictionary<string, object> Args(object value) => McpJson.Arguments(McpJson.Serialize(value));
    private static string Status(object outcome) => (string)McpJson.Arguments(McpJson.Serialize(outcome))["status"];

    private static bool Expect(string name, string expected, string actual)
    {
        var ok = expected == actual;
        Console.WriteLine($"{(ok ? "ok  " : "FAIL")} {name}: {actual}");
        if (!ok) _failures++;
        return ok;
    }

    /// <summary>打印写入签名的第一处差异。签名只含本探针写入的测试文字。</summary>
    private static void Diff(string name, XElement expected, XElement written)
    {
        if (expected == null || written == null) { Console.WriteLine($"     {name}: no expected/written page to compare"); return; }
        var left = McpPageModel.SignatureContent(expected, false, true); var right = McpPageModel.SignatureContent(written, false, true);
        var at = 0; while (at < left.Length && at < right.Length && left[at] == right[at]) at++;
        string Around(string s) => s.Substring(Math.Max(0, at - 120), Math.Min(s.Length - Math.Max(0, at - 120), 360));
        Console.WriteLine($"     first difference at {at}\n     expected: {Around(left)}\n     written:  {Around(right)}");
        // 签名第一段每行对应一个对象，按同样顺序找出两边的对象。
        var line = left.Substring(0, at).Count(c => c == '\n');
        XElement At(XElement page) => page.Descendants().Where(e => new[] { "Title", "Outline", "OE", "Table", "Row", "Cell", "Image" }.Contains(e.Name.LocalName)).ElementAtOrDefault(line);
        string Xml(XElement e) { if (e == null) return "(none)"; var copy = new XElement(e); copy.Elements(One + "OEChildren").Remove(); var s = copy.ToString(SaveOptions.DisableFormatting); return s.Length > 900 ? s.Substring(0, 900) : s; }
        Console.WriteLine($"     expected object: {Xml(At(expected))}\n     written object:  {Xml(At(written))}");
    }

    /// <summary>新建笔记本可能还没有回收站分区组：在测试笔记本里删一页一分区（进入回收站）让它出现。</summary>
    private static void EnsureRecycleBin(IApplication app, OneNoteApi api, string notebook)
    {
        bool Has() => McpReadService.Xml(api.GetHierarchy(notebook, HierarchyScope.hsPages)).DescendantsAndSelf(One + "SectionGroup").Any(e => (string)e.Attribute("isRecycleBin") == "true");
        Console.WriteLine("recycle bin present in new notebook: " + Has());
        if (Has()) return;
        app.OpenHierarchy("回收站准备.one", notebook, out var section, CreateFileType.cftSection);
        app.CreateNewPage(section, out var page, NewPageStyle.npsBlankPageWithTitle);
        app.DeleteHierarchy(page, default(DateTime), false);
        app.DeleteHierarchy(section, default(DateTime), false);
        Console.WriteLine("recycle bin present after deleting a probe page: " + Has());
        if (!Has()) throw new InvalidOperationException("测试笔记本没有可确认的回收站，工作区工具会拒绝写入。");
    }

    /// <summary>给页面追加一个带图片数据的文本框，复制和移动要连同图片一起核验。</summary>
    private static void AddImage(OneNoteApi api, string pageId)
    {
        string png;
        using (var bitmap = new System.Drawing.Bitmap(8, 8))
        using (var bytes = new MemoryStream())
        { bitmap.SetPixel(1, 1, System.Drawing.Color.Red); bitmap.Save(bytes, System.Drawing.Imaging.ImageFormat.Png); png = Convert.ToBase64String(bytes.ToArray()); }
        var page = AgentPageSnapshot.ParsePage(api.GetPageContent(pageId, PageInfo.piBasic));
        var (x, y) = PageEditor.NextFreePosition(page);
        var outline = new XElement(One + "Outline", new XElement(One + "Position", new XAttribute("x", x), new XAttribute("y", y), new XAttribute("z", 1)),
            new XElement(One + "OEChildren", new XElement(One + "OE", new XElement(One + "T", new XCData("图片说明"))),
                new XElement(One + "OE", new XElement(One + "Image", new XAttribute("format", "png"),
                    new XElement(One + "Size", new XAttribute("width", 16), new XAttribute("height", 16)), new XElement(One + "Data", png)))));
        ((IOneNotePageAccess)api).UpdatePageContent(PageEditor.BuildPageChanges(pageId, outline), AgentPageSnapshot.Modified(page));
    }
}
