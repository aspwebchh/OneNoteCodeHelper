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
                        // set_list / set_tag：项目符号加重要标记、编号、已完成的待办。
                        new XElement(ns + "OE", new XElement(ns + "T", new XCData("第一步：检查环境"))),
                        new XElement(ns + "OE", new XElement(ns + "T", new XCData("第二步：执行部署"))),
                        new XElement(ns + "OE", new XElement(ns + "T", new XCData("通知测试同事"))),
                        new XElement(ns + "OE", new XElement(ns + "Table", new XAttribute("bordersVisible", "true"), new XAttribute("hasHeaderRow", "false"),
                            new XElement(ns + "Columns", new XElement(ns + "Column", new XAttribute("index", 0), new XAttribute("width", 200)),
                                new XElement(ns + "Column", new XAttribute("index", 1), new XAttribute("width", 200), new XAttribute("isLocked", true))),
                            new XElement(ns + "Row", Cell(ns, "表格内文字"), Cell(ns, "保持原样的单元格")))),
                        new XElement(ns + "OE", new XElement(ns + "Image", new XAttribute("format", "png"),
                            new XElement(ns + "Size", new XAttribute("width", 16), new XAttribute("height", 16)), new XElement(ns + "Data", png))))),
                // 第二个文本框覆盖结构调整：整框提交、整框撤销。第一个文本框没有结构改动，仍按段落逐项提交。
                new XElement(ns + "Outline", new XElement(ns + "Position", new XAttribute("x", 36), new XAttribute("y", 520), new XAttribute("z", 1)),
                    new XElement(ns + "Size", new XAttribute("width", 500), new XAttribute("height", 200)),
                    new XElement(ns + "OEChildren", Line(ns, "结构调整说明"), Line(ns, ""), Line(ns, ""), Line(ns, "第一项"), Line(ns, "第一项的细节"), Line(ns, "第一项的补充"),
                        Line(ns, "| 名称 | 说明 |"), Line(ns, "|---|---|"), Line(ns, "| <b>甲</b> | 见<a href='https://example.com'>链接</a> |"),
                        Line(ns, "应当放在最前面的结论"), Line(ns, ""))),
                // 第三个文本框覆盖跨框移动和合并：先挪走一段，其余（含表格、图片）并进第二个文本框，提交后删掉，撤销时重建。
                new XElement(ns + "Outline", new XElement(ns + "Position", new XAttribute("x", 36), new XAttribute("y", 900), new XAttribute("z", 2)),
                    new XElement(ns + "Size", new XAttribute("width", 500), new XAttribute("height", 150)),
                    new XElement(ns + "OEChildren",
                        new XElement(ns + "OE", new XElement(ns + "T", new XCData("合并来源说明")), new XElement(ns + "OEChildren", Line(ns, "合并来源的细节"))),
                        new XElement(ns + "OE", new XElement(ns + "Table", new XAttribute("bordersVisible", "true"),
                            new XElement(ns + "Columns", new XElement(ns + "Column", new XAttribute("index", 0), new XAttribute("width", 120)),
                                new XElement(ns + "Column", new XAttribute("index", 1), new XAttribute("width", 120))),
                            new XElement(ns + "Row", Cell(ns, "来源表格甲"), Cell(ns, "来源表格乙")), new XElement(ns + "Row", Cell(ns, "来源表格丙"), Cell(ns, "来源表格丁")))),
                        Line(ns, "来源第二段"),
                        new XElement(ns + "OE", new XElement(ns + "Image", new XAttribute("format", "png"),
                            new XElement(ns + "Size", new XAttribute("width", 16), new XAttribute("height", 16)), new XElement(ns + "Data", png))))));
            api.UpdatePageContent(page.ToString(SaveOptions.DisableFormatting), AgentPageSnapshot.Modified(initial));
            var beforeXml = api.GetPageContent(pageId, PageInfo.piBasic);
            File.WriteAllText(Path.Combine(directory, "before.xml"), beforeXml);
            var snapshot = new AgentPageSnapshot(beforeXml, null, new AgentOptions());
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
            var first = editable.First(b => b.Text.StartsWith("第一步", StringComparison.Ordinal)).Id;
            var second = editable.First(b => b.Text.StartsWith("第二步", StringComparison.Ordinal)).Id;
            var todo = editable.First(b => b.Text == "通知测试同事").Id;
            // 父段落带下级段落：撤销时去掉列表要重建它，下级段落应保持原 ID。
            Execute(tools, "set_list", new { snapshot_id = snapshot.SnapshotId, block_ids = new[] { first, parent.Id }, list = "bullet" });
            Execute(tools, "set_list", new { snapshot_id = snapshot.SnapshotId, block_ids = new[] { second }, list = "number" });
            Execute(tools, "set_tag", new { snapshot_id = snapshot.SnapshotId, block_ids = new[] { first }, tag = "important" });
            Execute(tools, "set_tag", new { snapshot_id = snapshot.SnapshotId, block_ids = new[] { todo }, tag = "todo", completed = true });
            var grid = snapshot.Tables.First(t => t.Editable);
            Execute(tools, "set_table_style", new { snapshot_id = snapshot.SnapshotId, table_ids = new[] { grid.Id }, style = new { header_row = true, header_shading = "#DEEAF6" } });
            // 结构调整：删空行、移动、缩进、插入、转表格，再加一处格式修改，一起整框提交。
            string Id(string text) => editable.First(b => b.Text == text).Id;
            var structureOutline = snapshot.Blocks.First(b => b.Text == "结构调整说明").ContainerId;
            Execute(tools, "remove_blank_lines", new { snapshot_id = snapshot.SnapshotId, mode = "collapse" });
            Execute(tools, "move_blocks", new { snapshot_id = snapshot.SnapshotId, block_ids = new[] { Id("应当放在最前面的结论") }, target_id = Id("结构调整说明"), position = "before" });
            Execute(tools, "set_indent", new { snapshot_id = snapshot.SnapshotId, block_ids = new[] { Id("第一项的细节"), Id("第一项的补充") }, direction = "in" });
            Execute(tools, "insert_blocks", new { snapshot_id = snapshot.SnapshotId, target_id = Id("应当放在最前面的结论"), position = "after", paragraphs = new object[] {
                new { text = "摘要", preset_id = "heading2" }, new { text = "要点：先备份再部署", preset_id = "body", list = "bullet" } } });
            Execute(tools, "text_to_table", new { snapshot_id = snapshot.SnapshotId, block_ids = new[] { Id("| 名称 | 说明 |"), Id("|---|---|"), Id("| 甲 | 见链接 |") },
                delimiter = "pipe", header_shading = "#E2EFDA" });
            Execute(tools, "set_paragraph_style", new { snapshot_id = snapshot.SnapshotId, block_ids = new[] { Id("第一项") }, preset_id = "body" });
            // 跨框移动一段，再把第三个文本框其余内容整个并进第二个文本框。
            var mergedOutline = snapshot.Blocks.First(b => b.Text == "合并来源说明").ContainerId;
            Execute(tools, "move_blocks", new { snapshot_id = snapshot.SnapshotId, block_ids = new[] { Id("来源第二段") }, target_id = Id("第一项"), position = "after" });
            Execute(tools, "merge_outlines", new { snapshot_id = snapshot.SnapshotId, source_id = mergedOutline, target_id = Id("结构调整说明"), position = "after" });
            var structureBefore = OutlineText(AgentPageSnapshot.ParsePage(beforeXml), structureOutline);
            var mergedBefore = OutlineText(AgentPageSnapshot.ParsePage(beforeXml), mergedOutline);
            var expected = snapshot.CreateDraftPage();
            File.WriteAllText(Path.Combine(directory, "expected.xml"), expected.ToString());
            Execute(tools, "finish_edit", new { snapshot_id = snapshot.SnapshotId, draft_revision = snapshot.Revision });
            File.WriteAllText(Path.Combine(directory, "after.xml"), api.GetPageContent(pageId, PageInfo.piBasic));
            Console.WriteLine("Test section: " + path);
            Console.WriteLine("Result: " + tools.Report.Status + " " + tools.Report.Message);
            Console.WriteLine("Code blocks: " + tools.Report.CodeBlocks + ", tables: " + tools.Report.Tables + ", text tables: " + tools.Report.TextTables +
                ", outlines: " + tools.Report.Outlines + ", removed: " + tools.Report.Removed + ", moved: " + tools.Report.Moved + ", indented: " + tools.Report.Indented + ", inserted: " + tools.Report.Inserted +
                ", merged: " + tools.Report.Merged + ", leftover: " + tools.Report.Leftover);
            var actual = AgentPageSnapshot.ParsePage(api.GetPageContent(pageId, PageInfo.piBasic));
            // OneNote 实际写回的列表、标记和表格外观，供对照 AgentMarks / TableLook 的写法。
            foreach (var e in actual.Elements(ns + "TagDef").Concat(actual.Descendants(ns + "Tag")).Concat(actual.Descendants(ns + "List")))
                Console.WriteLine("Written " + e.Name.LocalName + ": " + e.ToString(SaveOptions.DisableFormatting));
            var writtenStructure = actual.Elements(ns + "Outline").First(o => (string)o.Attribute("objectID") == structureOutline);
            Console.WriteLine("Written structure: " + string.Join(" | ", writtenStructure.Descendants(ns + "OE").Where(e => e.Elements(ns + "T").Any())
                .Select(e => new string('>', e.Ancestors(ns + "OE").Count()) + AgentCode.PlainText(e))));
            foreach (var e in writtenStructure.Descendants(ns + "Table").Concat(writtenStructure.Descendants(ns + "OE").Take(3)))
                Console.WriteLine("Written " + e.Name.LocalName + ": " + e.ToString(SaveOptions.DisableFormatting));
            var mergedGone = actual.Elements(ns + "Outline").All(o => (string)o.Attribute("objectID") != mergedOutline);
            Console.WriteLine("Merged text box deleted: " + mergedGone + ", images: " + actual.Descendants(ns + "Image").Count() +
                ", binary images: " + AgentPageSnapshot.ParsePage(api.GetPageContent(pageId, PageInfo.piBinaryData)).Descendants(ns + "Image").Count());
            var writtenGrid = AgentTable.Find(actual, grid.ObjectId);
            if (writtenGrid != null) Console.WriteLine("Written table: borders=" + (string)writtenGrid.Attribute("bordersVisible") + " header=" + (string)writtenGrid.Attribute("hasHeaderRow") +
                " shading=" + string.Join(",", writtenGrid.Element(ns + "Row").Elements(ns + "Cell").Select(c => (string)c.Attribute("shadingColor"))));
            foreach (var b in snapshot.Blocks.Where(b => b.Changed))
            {
                var written = AgentCommitter.Find(actual, b.ObjectId);
                if (written == null) continue;
                var left = AgentPageSnapshot.SemanticFormat(AgentLayout.Find(expected, b.Id), expected);
                var right = AgentPageSnapshot.SemanticFormat(written, actual);
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
            var restored = AgentPageSnapshot.ParsePage(api.GetPageContent(pageId, PageInfo.piBasic));
            var marksLeft = restored.Descendants(ns + "Tag").Count() + restored.Descendants(ns + "List").Count();
            var child = snapshot.Blocks.Single(b => b.Text == "保留原样的子段落");
            var childKept = AgentCommitter.Find(restored, child.ObjectId) != null;
            Console.WriteLine("Tags and lists left after undo: " + marksLeft + ", child paragraph kept its ID: " + childKept);
            // 撤销后第二个文本框的文字和原来一致；只移动、缩进过的段落保持原 ID。
            var structureRestored = OutlineText(restored, structureOutline) == structureBefore;
            var movedKept = new[] { "应当放在最前面的结论", "第一项的细节", "第一项" }.All(text => AgentCommitter.Find(restored, snapshot.Blocks.Single(b => b.Text == text).ObjectId) != null);
            Console.WriteLine("Structure restored: " + structureRestored + ", moved paragraphs kept their IDs: " + movedKept + ", tables left: " + restored.Descendants(ns + "Table").Count());
            // 合并删掉的文本框重建：文字一致、ID 是新的；图片按二进制数据仍然完整。
            var rebuilt = restored.Elements(ns + "Outline").FirstOrDefault(o => o.Descendants(ns + "OE").Any(e => e.Elements(ns + "T").Any() && AgentCode.PlainText(e) == "合并来源说明"));
            var mergedRestored = rebuilt != null && (string)rebuilt.Attribute("objectID") != mergedOutline && OutlineText(restored, (string)rebuilt.Attribute("objectID")) == mergedBefore;
            var binaryImages = AgentPageSnapshot.ParsePage(api.GetPageContent(pageId, PageInfo.piBinaryData)).Descendants(ns + "Image").Count(i => i.Element(ns + "Data") != null);
            Console.WriteLine("Merged text box rebuilt: " + mergedRestored + " (" + (string)rebuilt?.Attribute("objectID") + "), images with data: " + binaryImages);
            var report = tools.Report;
            return report.Status == "Verified" && undo.Status == "Verified" && report.CodeBlocks == 2 && undo.CodeBlocks == 2 &&
                report.Tables == 1 && undo.Tables == 1 && marksLeft == 0 && childKept &&
                report.Outlines == 2 && undo.Outlines == 2 && report.TextTables == 1 && report.Removed == 2 && report.Moved == 2 && report.Indented == 2 && report.Inserted == 2 &&
                report.Merged == 1 && report.Leftover == 0 && mergedGone && mergedRestored && binaryImages == 2 && structureRestored && movedKept ? 0 : 1;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally { if (app != null && Marshal.IsComObject(app)) Marshal.FinalReleaseComObject(app); }
    }
    private static XElement Line(XNamespace ns, string html) => new XElement(ns + "OE", new XElement(ns + "T", new XCData(html)));
    private static string OutlineText(XElement page, string outlineId) => string.Join("|", page.Elements(OneNoteApi.One + "Outline")
        .First(o => (string)o.Attribute("objectID") == outlineId).Descendants(OneNoteApi.One + "OE").Where(e => e.Elements(OneNoteApi.One + "T").Any()).Select(AgentCode.PlainText));
    private static XElement Cell(XNamespace ns, string text) => new XElement(ns + "Cell",
        new XElement(ns + "OEChildren", new XElement(ns + "OE", new XElement(ns + "T", text))));
    private static void Execute(AgentTools tools, string name, object args) => tools.Execute(new AgentToolCall { Id = name, Name = name, Arguments = AgentChatClient.Serializer().Serialize(args) });
}
