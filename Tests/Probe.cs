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

/// <summary>显式运行才激活 OneNote，仅创建并编辑独立的测试分区。</summary>
internal static class Probe
{
    internal static int Run(string[] args)
    {
        if (args.Length != 1) { Console.Error.WriteLine("--probe requires an output directory for a NEW test section."); return 2; }
        IApplication app = null;
        try
        {
            var directory = Path.GetFullPath(args[0]);
            Directory.CreateDirectory(directory);
            var path = Directory.GetFiles(directory, "Agent-format-probe-*.one").FirstOrDefault()
                ?? Path.Combine(directory, "Agent-format-probe-" + Guid.NewGuid().ToString("N") + ".one");
            app = new ApplicationClass();
            app.OpenHierarchy(path, "", out var section, CreateFileType.cftSection);
            app.CreateNewPage(section, out var pageId, NewPageStyle.npsBlankPageWithTitle);
            var api = new OneNoteApi(app);
            var ns = OneNoteApi.One;
            string png;
            using (var bitmap = new System.Drawing.Bitmap(8, 8))
            using (var bytes = new MemoryStream())
            { bitmap.Save(bytes, System.Drawing.Imaging.ImageFormat.Png); png = Convert.ToBase64String(bytes.ToArray()); }
            var initial = AgentPageSnapshot.ParsePage(api.GetPageContent(pageId, PageInfo.piBasic));
            var page = new XElement(ns + "Page", new XAttribute("ID", pageId),
                new XElement(ns + "Title", new XElement(ns + "OE", new XElement(ns + "T", new XCData("Agent 格式回归测试")))),
                new XElement(ns + "Outline", new XElement(ns + "Position", new XAttribute("x", 36), new XAttribute("y", 100), new XAttribute("z", 0)),
                    new XElement(ns + "Size", new XAttribute("width", 500), new XAttribute("height", 200)),
                    new XElement(ns + "OEChildren",
                        new XElement(ns + "OE", new XElement(ns + "T", new XCData("部署前准备"))),
                        new XElement(ns + "OE", new XElement(ns + "T", new XCData("&nbsp;先<b>备份</b>配置。<a href='https://example.com'>参考链接</a><br>第二行😀"))),
                        new XElement(ns + "OE", new XAttribute("style", "font-family:Consolas;font-size:11pt"), new XElement(ns + "T", new XCData("List&lt;String&gt; code;"))),
                        new XElement(ns + "OE", new XElement(ns + "T", "父段落"), new XElement(ns + "OEChildren", new XElement(ns + "OE", new XElement(ns + "T", "保留原样的子段落")))),
                        // 当普通正文输入的代码，由 highlight_code 转为代码框。
                        new XElement(ns + "OE", new XElement(ns + "T", new XCData("public class Demo {"))),
                        new XElement(ns + "OE", new XElement(ns + "T", new XCData("&nbsp;&nbsp;&nbsp;&nbsp;int x = 1;"))),
                        new XElement(ns + "OE", new XElement(ns + "T", new XCData("}"))),
                        new XElement(ns + "OE", new XElement(ns + "Table", new XAttribute("bordersVisible", "true"), new XAttribute("hasHeaderRow", "false"),
                            new XElement(ns + "Columns", new XElement(ns + "Column", new XAttribute("index", 0), new XAttribute("width", 200)),
                                new XElement(ns + "Column", new XAttribute("index", 1), new XAttribute("width", 200), new XAttribute("isLocked", true))),
                            new XElement(ns + "Row", Cell(ns, "表格内文字"), Cell(ns, "保持原样的单元格")))),
                        new XElement(ns + "OE", new XElement(ns + "Image", new XAttribute("format", "png"),
                            new XElement(ns + "Size", new XAttribute("width", 16), new XAttribute("height", 16)), new XElement(ns + "Data", png))))));
            api.UpdatePageContent(page.ToString(SaveOptions.DisableFormatting), AgentPageSnapshot.Modified(initial));
            var beforeXml = api.GetPageContent(pageId, PageInfo.piBasic);
            File.WriteAllText(Path.Combine(directory, "before.xml"), beforeXml);
            var snapshot = new AgentPageSnapshot(beforeXml, null, new AgentOptions { EnableParagraphSpacing = true, EnableMixedOutlines = true });
            foreach (var b in snapshot.Blocks) Console.WriteLine("Probe block " + b.Id + ": " + (b.ProtectedReason ?? "editable") + ", chars=" + b.Text.Length);
            var tools = new AgentTools(snapshot, new AgentCommitter(api), CancellationToken.None);
            var editable = snapshot.Blocks.Where(b => b.Editable).ToList();
            var mono = snapshot.Blocks.Single(b => b.CodeCandidate);
            Execute(tools, "read_blocks", new { snapshot_id = snapshot.SnapshotId, block_ids = editable.Select(b => b.Id).Concat(new[] { mono.Id }).ToArray() });
            var heading = editable.First(b => b.Text.Contains("部署前准备"));
            Execute(tools, "set_paragraph_style", new { snapshot_id = snapshot.SnapshotId, block_ids = new[] { heading.Id }, preset_id = "heading1", overrides = new { alignment = "center" } });
            var body = editable.First(b => b.Text.Contains("先"));
            Execute(tools, "set_paragraph_style", new { snapshot_id = snapshot.SnapshotId, block_ids = new[] { body.Id }, preset_id = "body" });
            Execute(tools, "set_text_style", new { snapshot_id = snapshot.SnapshotId, targets = new[] { new { block_id = body.Id, quote = "备份", occurrence = 1, style = new { bold = false, color = "#1F4E79" } } } });
            Execute(tools, "set_text_style", new { snapshot_id = snapshot.SnapshotId, targets = new[] { new { block_id = body.Id, quote = "第二行", occurrence = 1, style = new { italic = true, underline = true } } } });
            var parent = editable.First(b => b.Text == "父段落");
            Execute(tools, "set_paragraph_style", new { snapshot_id = snapshot.SnapshotId, block_ids = new[] { parent.Id }, preset_id = "heading2" });
            var cell = editable.First(b => b.Text == "表格内文字");
            Execute(tools, "set_paragraph_style", new { snapshot_id = snapshot.SnapshotId, block_ids = new[] { cell.Id }, preset_id = "body", overrides = new { alignment = "right" } });
            Execute(tools, "highlight_code", new { snapshot_id = snapshot.SnapshotId, block_ids = new[] { mono.Id }, language = "java" });
            var plain = editable.Where(b => b.Text.StartsWith("public class", StringComparison.Ordinal) || b.Text.Contains("int x") || b.Text == "}").Select(b => b.Id).ToArray();
            Execute(tools, "highlight_code", new { snapshot_id = snapshot.SnapshotId, block_ids = plain, language = "auto" });
            var expected = new XElement(snapshot.Page);
            expected.Elements(ns + "QuickStyleDef").Remove();
            expected.AddFirst(snapshot.DraftStyles.Elements().Select(e => new XElement(e)));
            foreach (var b in snapshot.Blocks.Where(b => b.Changed)) AgentPageSnapshot.CopyFormat(b.Draft, AgentCommitter.Find(expected, b.ObjectId));
            File.WriteAllText(Path.Combine(directory, "expected.xml"), expected.ToString());
            Execute(tools, "finish_edit", new { snapshot_id = snapshot.SnapshotId, draft_revision = snapshot.Revision });
            File.WriteAllText(Path.Combine(directory, "after.xml"), api.GetPageContent(pageId, PageInfo.piBasic));
            Console.WriteLine("Test section: " + path);
            Console.WriteLine("Result: " + tools.Report.Status + " " + tools.Report.Message);
            Console.WriteLine("Code blocks: " + tools.Report.CodeBlocks);
            var actual = AgentPageSnapshot.ParsePage(api.GetPageContent(pageId, PageInfo.piBasic));
            foreach (var b in snapshot.Blocks.Where(b => b.Changed))
            {
                var left = AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(expected, b.ObjectId), expected);
                var right = AgentPageSnapshot.SemanticFormat(AgentCommitter.Find(actual, b.ObjectId), actual);
                if (left != right)
                {
                    var i = 0; while (i < Math.Min(left.Length, right.Length) && left[i] == right[i]) i++;
                    Console.WriteLine("Mismatch " + b.Id + " expected: " + left.Substring(Math.Max(0, i - 25), Math.Min(220, left.Length - Math.Max(0, i - 25))));
                    Console.WriteLine("Mismatch " + b.Id + " actual: " + right.Substring(Math.Max(0, i - 25), Math.Min(220, right.Length - Math.Max(0, i - 25))));
                }
            }
            var undo = new AgentCommitter(api).Undo(pageId, tools.Report, snapshot.Options, CancellationToken.None);
            File.WriteAllText(Path.Combine(directory, "after-undo.xml"), api.GetPageContent(pageId, PageInfo.piBasic));
            Console.WriteLine("Undo: " + undo.Status + " " + undo.Message);
            Console.WriteLine("The new test page is kept for visual inspection. No existing page was edited.");
            return tools.Report.Status == "Verified" && undo.Status == "Verified" && tools.Report.CodeBlocks == 2 && undo.CodeBlocks == 2 ? 0 : 1;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally { if (app != null && Marshal.IsComObject(app)) Marshal.FinalReleaseComObject(app); }
    }
    private static XElement Cell(XNamespace ns, string text) => new XElement(ns + "Cell",
        new XElement(ns + "OEChildren", new XElement(ns + "OE", new XElement(ns + "T", text))));
    private static void Execute(AgentTools tools, string name, object args) => tools.Execute(new AgentToolCall { Id = name, Name = name, Arguments = AgentChatClient.Serializer().Serialize(args) });
}
