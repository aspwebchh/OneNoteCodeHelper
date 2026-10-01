using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using OneNoteCodeHelper.Highlighting;
using OneNoteCodeHelper.Highlighting.Themes;
using OneNoteCodeHelper.Services;
using OneNoteCodeHelper.Services.Agent;

/// <summary>显式运行才创建全新的专用分区和页面；不读写用户已有笔记，不调用 AI。</summary>
internal static class CodeSpacingProbe
{
    private static XNamespace One => OneNoteApi.One;
    internal static int Run(string directory)
    {
        IApplication app = null;
        try
        {
            directory = Path.GetFullPath(directory); Directory.CreateDirectory(directory);
            var path = Path.Combine(directory, "Code-spacing-probe-" + Guid.NewGuid().ToString("N") + ".one");
            app = new ApplicationClass();
            app.OpenHierarchy(path, "", out var section, CreateFileType.cftSection);
            app.CreateNewPage(section, out var pageId, NewPageStyle.npsBlankPageWithTitle);
            var api = new OneNoteApi(app);
            var initial = AgentPageSnapshot.ParsePage(api.GetPageContent(pageId, PageInfo.piBasic));
            var table = CodeBlockBuilder.BuildTable("x = 1\n\nprint(x)", LanguageRegistry.Find("python"), CodeThemes.Light, new AddInSettings { FontFamily = "Consolas" });
            var page = new XElement(One + "Page", new XAttribute("ID", pageId),
                new XElement(One + "Title", Line("Code spacing probe")),
                new XElement(One + "Outline", new XElement(One + "Position", new XAttribute("x", 36), new XAttribute("y", 100), new XAttribute("z", 0)),
                    new XElement(One + "Size", new XAttribute("width", 500), new XAttribute("height", 500)),
                    new XElement(One + "OEChildren", Line("<span style='font-weight:bold'>Before existing<br>&nbsp;<br>&nbsp;</span>"), Mono("<span><br>&nbsp;<br>&nbsp;</span>"), Mono("&nbsp;"),
                        new XElement(One + "OE", table), Line("<span>&nbsp;<br>&nbsp;<br><a href='https://example.com'>After existing</a></span>"),
                        Line("Unrelated A<br><br>second line"), Line("&nbsp;"), Line("&nbsp;"), Line("Unrelated B"),
                        Line("Before conversion"), Mono("print('one')"), Mono("&nbsp;"), Mono("print('two')"), Line("After conversion"))));
            api.UpdatePageContent(page.ToString(SaveOptions.DisableFormatting), AgentPageSnapshot.Modified(initial));
            var before = api.GetPageContent(pageId, PageInfo.piBasic); File.WriteAllText(Path.Combine(directory, "before.xml"), before);
            var s = new AgentPageSnapshot(before, null, new AgentOptions());
            var t = new AgentTools(s, new AgentCommitter(api), CancellationToken.None, new AddInSettings { FontFamily = "Consolas" });
            Execute(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = s.Blocks.Where(b => b.Editable || b.CodeCandidate).Select(b => b.Id).ToArray() });
            var first = s.Blocks.FindIndex(b => b.Text == "print('one')"); var last = s.Blocks.FindIndex(b => b.Text == "print('two')");
            Execute(t, "highlight_code", new { snapshot_id = s.SnapshotId, block_ids = s.Blocks.Skip(first).Take(last - first + 1).Select(b => b.Id).ToArray(), language = "python" });
            var spacing = Execute(t, "normalize_code_spacing", new { snapshot_id = s.SnapshotId });
            Console.WriteLine("Spacing draft: " + AgentChatClient.Serializer().Serialize(spacing));
            var planned = AgentChatClient.Parse(AgentChatClient.Serializer().Serialize(spacing));
            File.WriteAllText(Path.Combine(directory, "expected.xml"), s.CreateDraftPage().ToString());
            Execute(t, "finish_edit", new { snapshot_id = s.SnapshotId, draft_revision = s.Revision });
            var after = api.GetPageContent(pageId, PageInfo.piBasic); File.WriteAllText(Path.Combine(directory, "after.xml"), after);
            Console.WriteLine("Commit: " + t.Report.Status + " " + t.Report.Message);
            if (t.Report.Status != "Verified") return 1;
            var checkedPage = new AgentPageSnapshot(after, null, new AgentOptions());
            foreach (var block in checkedPage.Blocks) block.Read = true;
            var check = AgentCodeSpacing.Build(checkedPage);
            var valid = XNode.DeepEquals(check.Layout, checkedPage.Layout) && check.Skipped.Count == 0 && t.Report.CodeBlocks == 1 && t.Report.Removed > 0 &&
                t.Report.Removed == (int)AiClient.Get(planned, "removed_paragraphs") && t.Report.RemovedSoftLines == (int)AiClient.Get(planned, "removed_soft_lines") &&
                t.Report.InsertedBlankLines == (int)AiClient.Get(planned, "inserted_paragraphs");
            if (t.Report.RemovedSoftLines == 0) Console.WriteLine("This OneNote build trimmed edge BRs when creating the fixture; soft-edge cases are covered by offline regression.");
            var undo = new AgentCommitter(api).Undo(pageId, t.Report, s.Options, CancellationToken.None);
            var restored = api.GetPageContent(pageId, PageInfo.piBasic); File.WriteAllText(Path.Combine(directory, "after-undo.xml"), restored);
            Console.WriteLine("Undo: " + undo.Status + " " + undo.Message);
            var same = Projection(AgentPageSnapshot.ParsePage(before)) == Projection(AgentPageSnapshot.ParsePage(restored));
            Console.WriteLine("One-line gaps verified: " + valid + "; original text and blank lines restored: " + same);
            Console.WriteLine("New dedicated test section: " + path);
            return valid && same && undo.Status == "Verified" ? 0 : 1;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally { if (app != null && Marshal.IsComObject(app)) Marshal.FinalReleaseComObject(app); }
    }
    private static XElement Line(string html) => new XElement(One + "OE", new XElement(One + "T", new XCData(html)));
    private static XElement Mono(string html) { var line = Line(html); line.SetAttributeValue("style", "font-family:Consolas;font-size:10pt"); return line; }
    private static string Projection(XElement page) => string.Join("|", page.Descendants(One + "OE").Where(e => e.Elements(One + "T").Any()).Select(e => new AgentRichText(e).Text.Replace('\u00a0', ' ')));
    private static object Execute(AgentTools tools, string name, object args) => tools.Execute(new AgentToolCall { Id = name, Name = name, Arguments = AgentChatClient.Serializer().Serialize(args) });
}
