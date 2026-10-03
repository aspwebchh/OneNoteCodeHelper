using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using OneNoteCodeHelper.Services;
using OneNoteCodeHelper.Services.Agent;

/// <summary>显式运行才打开 OneNote，所有写入都在本次新建的独立测试分区。</summary>
internal static class TitleProbe
{
    private static XNamespace One => OneNoteApi.One;

    internal static int Run(string directory)
    {
        IApplication app = null;
        try
        {
            directory = Path.GetFullPath(directory);
            Directory.CreateDirectory(directory);
            Console.WriteLine("Page title probe artifacts: " + directory);
            app = new ApplicationClass();
            app.OpenHierarchy(Path.Combine(directory, "Agent-title-probe-" + Guid.NewGuid().ToString("N") + ".one"), "", out var section, CreateFileType.cftSection);
            var api = new OneNoteApi(app);
            CheckRoundTrip(app, api, section, directory, "empty-page", NewPageStyle.npsBlankPageWithTitle, false);
            CheckRoundTrip(app, api, section, directory, "blank", NewPageStyle.npsBlankPageWithTitle, false);
            CheckRoundTrip(app, api, section, directory, "no-title", NewPageStyle.npsBlankPageNoTitle, false);
            CheckRoundTrip(app, api, section, directory, "rename", NewPageStyle.npsBlankPageWithTitle, true);
            CheckConflict(app, api, section);
            Console.WriteLine("Page title probe: all checks passed. Artifacts: " + directory);
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally { if (app != null) Marshal.ReleaseComObject(app); }
    }

    private static string Create(IApplication app, OneNoteApi api, string section, NewPageStyle style, bool body = true)
    {
        app.CreateNewPage(section, out var id, style);
        if (!body) return id;
        var before = AgentPageSnapshot.ParsePage(api.GetPageContent(id, PageInfo.piBasic));
        var outline = new XElement(One + "Outline", new XElement(One + "Position", new XAttribute("x", 36), new XAttribute("y", 140), new XAttribute("z", 0)),
            new XElement(One + "Size", new XAttribute("width", 500), new XAttribute("height", 100)),
            new XElement(One + "OEChildren", new XElement(One + "OE", new XElement(One + "T", new XCData("主题说明：<b>配置</b>与部署流程。"))),
                new XElement(One + "OE", new XElement(One + "T", new XCData("参考<a href='https://example.com'>文档</a>，保持原有内容。")))));
        api.UpdatePageContent(PageEditor.BuildPageChanges(id, outline), AgentPageSnapshot.Modified(before));
        return id;
    }

    private static void PutTitle(OneNoteApi api, string id, string text)
    {
        var page = AgentPageSnapshot.ParsePage(api.GetPageContent(id, PageInfo.piBasic));
        var title = new XElement(page.Element(One + "Title") ?? AgentPageTitle.Empty());
        var line = title.Element(One + "OE");
        if (line == null) { line = new XElement(One + "OE"); title.Add(line); }
        if (!line.Elements(One + "T").Any()) line.Add(new XElement(One + "T", new XCData("")));
        RichParagraph.Parse(line).Apply(text);
        api.UpdatePageContent(PageEditor.BuildPageChanges(id, title), AgentPageSnapshot.Modified(page));
    }

    private static void CheckRoundTrip(IApplication app, OneNoteApi api, string section, string directory, string kind, NewPageStyle style, bool rename)
    {
        var id = Create(app, api, section, style, kind != "empty-page");
        if (rename) PutTitle(api, id, "原标题：配置说明");
        var beforeXml = api.GetPageContent(id, PageInfo.piBasic);
        File.WriteAllText(Path.Combine(directory, kind + "-before.xml"), beforeXml);
        var s = new AgentPageSnapshot(beforeXml, null, new AgentOptions());
        var c = new AgentCommitter(api);
        var t = new AgentTools(s, c, CancellationToken.None);
        Read(t, s);
        if (rename)
        {
            Execute(t, "set_page_title", new { snapshot_id = s.SnapshotId, title = "默认不得替换" });
            if (s.TitleEdit != null) throw new Exception("Default title preservation failed.");
        }
        const string wanted = "配置  <指南> & C# 😀";
        Execute(t, "set_page_title", new { snapshot_id = s.SnapshotId, title = wanted, replace_existing = rename });
        if (rename || kind == "no-title")
        {
            var titleBlock = s.Blocks.Single(b => b.IsPageTitle);
            Execute(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { titleBlock.Id }, preset_id = "page_title", overrides = new { alignment = "center" } });
            if (rename)
            {
                var body = s.Blocks.First(b => !b.IsPageTitle && b.Editable);
                Execute(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { body.Id }, preset_id = "body" });
            }
        }
        Execute(t, "finish_edit", new { snapshot_id = s.SnapshotId, draft_revision = s.Revision });
        var r = t.Report;
        var afterXml = api.GetPageContent(id, PageInfo.piBasic);
        File.WriteAllText(Path.Combine(directory, kind + "-after.xml"), afterXml);
        var after = AgentPageSnapshot.ParsePage(afterXml);
        Console.WriteLine(kind + " write: " + r.Status + " / " + r.Message);
        if (r.Status != "Verified" || !r.TitleChanged || !r.CanUndo || AgentPageTitle.Text(after.Element(One + "Title")) != wanted || (string)after.Attribute("name") != wanted)
            throw new Exception(kind + " native title verification failed.");
        if (!rename && Body(s.Page) != Body(after)) throw new Exception(kind + " body changed during title-only write.");
        var undo = c.Undo(id, r, s.Options, CancellationToken.None);
        var undoXml = api.GetPageContent(id, PageInfo.piBasic);
        File.WriteAllText(Path.Combine(directory, kind + "-undo.xml"), undoXml);
        var restored = AgentPageSnapshot.ParsePage(undoXml);
        Console.WriteLine(kind + " undo: " + undo.Status + " / " + undo.Message);
        if (undo.Status != "Verified" || AgentPageTitle.Text(restored.Element(One + "Title")) != AgentPageTitle.Text(s.Page.Element(One + "Title")) || Body(s.Page) != Body(restored))
            throw new Exception(kind + " native title undo failed.");
    }

    private static void CheckConflict(IApplication app, OneNoteApi api, string section)
    {
        var id = Create(app, api, section, NewPageStyle.npsBlankPageWithTitle);
        var s = new AgentPageSnapshot(api.GetPageContent(id, PageInfo.piBasic), null, new AgentOptions());
        var t = new AgentTools(s, new AgentCommitter(api), CancellationToken.None);
        Read(t, s); Execute(t, "set_page_title", new { snapshot_id = s.SnapshotId, title = "Agent 标题" });
        PutTitle(api, id, "处理中用户修改的标题");
        var r = new AgentCommitter(api).Commit(s, CancellationToken.None);
        var actual = AgentPageSnapshot.ParsePage(api.GetPageContent(id, PageInfo.piBasic));
        if (r.TitleChanged || r.Conflicts != 1 || AgentPageTitle.Text(actual.Element(One + "Title")) != "处理中用户修改的标题")
            throw new Exception("Native title conflict protection failed.");
        Console.WriteLine("title conflict: passed");
    }

    private static string Body(XElement page) => string.Join("\n", page.Elements(One + "Outline").Descendants(One + "OE").Where(e => e.Elements(One + "T").Any())
        .Select(e => AgentPageSnapshot.SemanticFormat(e, page)));

    private static void Read(AgentTools t, AgentPageSnapshot s)
    {
        var ids = s.Blocks.Where(b => b.Editable).Select(b => b.Id).ToArray();
        if (ids.Length > 0) Execute(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = ids });
    }
    private static void Execute(AgentTools t, string name, object args) => t.Execute(new AgentToolCall { Id = "title-probe", Name = name, Arguments = AgentChatClient.Serializer().Serialize(args) });
}
