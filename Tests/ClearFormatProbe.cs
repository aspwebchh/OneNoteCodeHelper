using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using OneNoteCodeHelper.Highlighting;
using OneNoteCodeHelper.Highlighting.Themes;
using OneNoteCodeHelper.Services;
using OneNoteCodeHelper.Services.Agent;

/// <summary>
/// 清除格式相关改动的真机往返探针：高亮写法（background / background-color）、空 T 的格式核验、清除格式和拆框的写回与撤销。
/// 显式运行才创建全新的专用分区，每项一页；不读写用户已有笔记，不调用 AI。
/// </summary>
internal static class ClearFormatProbe
{
    private static XNamespace One => OneNoteApi.One;
    private static readonly Regex Highlighted = new Regex("background-color:(?!transparent)");

    internal static int Run(string directory)
    {
        IApplication app = null;
        try
        {
            directory = Path.GetFullPath(directory); Directory.CreateDirectory(directory);
            var path = Path.Combine(directory, "Clear-format-probe-" + Guid.NewGuid().ToString("N") + ".one");
            app = new ApplicationClass();
            app.OpenHierarchy(path, "", out var section, CreateFileType.cftSection);
            var api = new OneNoteApi(app);
            var results = new List<(string Name, bool Passed)>
            {
                Scenario("1 高亮写法直接写入", () => RawHighlight(app, api, section, directory)),
                Scenario("2 已有工具改写带高亮的文字", () => ToolsKeepHighlight(app, api, section, directory)),
                Scenario("3 未改动的空段落在写回后格式不变", () => BlankParagraphs(app, api, section, directory)),
                Scenario("4 清除格式、拆代码框、表格去底色及撤销", () => ClearEverything(app, api, section, directory))
            };
            Console.WriteLine();
            foreach (var (name, passed) in results) Console.WriteLine((passed ? "PASS " : "FAIL ") + name);
            Console.WriteLine("New dedicated test section: " + path);
            return results.All(r => r.Passed) ? 0 : 1;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally { if (app != null && Marshal.IsComObject(app)) Marshal.FinalReleaseComObject(app); }
    }

    private static (string, bool) Scenario(string name, Func<bool> body)
    {
        Console.WriteLine();
        Console.WriteLine("== " + name);
        try { return (name, body()); }
        catch (Exception ex) { Console.WriteLine("  异常：" + ex); return (name, false); }
    }

    /// <summary>两种高亮写法分别写在 span 和 T 上，回读看 OneNote 保留了什么。</summary>
    private static bool RawHighlight(IApplication app, OneNoteApi api, string section, string directory)
    {
        var page = Create(app, api, section, "高亮写法", Outline(100,
            Line("<span style='background:yellow'>span 旧写法</span>"),
            Line("<span style='background-color:yellow'>span 新写法</span>"),
            Styled(Line("T 旧写法"), "background:yellow"),
            Styled(Line("T 新写法"), "background-color:yellow")));
        var read = AgentPageSnapshot.ParsePage(Read(api, page, directory, "1-readback.xml"));
        var ok = true;
        foreach (var oe in Body(read))
        {
            var kept = Highlighted.IsMatch(new AgentRichText(oe).Signature(read, true));
            Console.WriteLine($"  {AgentCode.PlainText(oe)}：高亮{(kept ? "保留" : "丢失")}；T style=\"{(string)oe.Element(One + "T").Attribute("style")}\"；HTML={oe.Element(One + "T").Value}");
            ok &= kept;
        }
        return ok;
    }

    /// <summary>set_text_style、set_paragraph_style 会经 Css.Read/Write 重写 span 和 T 的样式；写回后高亮仍须在，核验通过，撤销还原。</summary>
    private static bool ToolsKeepHighlight(IApplication app, OneNoteApi api, string section, string directory)
    {
        var page = Create(app, api, section, "工具与高亮", Outline(100,
            Line("普通<span style='background:yellow'>高亮</span>结尾一"),
            Line("普通<span style='background:yellow'>高亮</span>结尾二"),
            Styled(Line("整段高亮三"), "background:yellow")));
        var beforeXml = Read(api, page, directory, "2-before.xml"); var before = AgentPageSnapshot.ParsePage(beforeXml);
        var access = new Logged(api, directory, "2");
        var s = new AgentPageSnapshot(beforeXml, null, new AgentOptions());
        var t = new AgentTools(s, new AgentCommitter(access), CancellationToken.None);
        string Id(string text) => s.Blocks.Single(b => b.Text.Contains(text)).Id;
        Execute(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = s.Blocks.Where(b => b.Editable).Select(b => b.Id).ToArray() });
        Execute(t, "set_text_style", new { snapshot_id = s.SnapshotId, targets = new[] { new { block_id = Id("结尾一"), quote = "高亮", occurrence = 1, style = new { bold = true } } } });
        Execute(t, "set_paragraph_style", new { snapshot_id = s.SnapshotId, block_ids = new[] { Id("结尾二"), Id("整段高亮三") }, preset_id = "body" });
        Console.WriteLine("  草稿 HTML：" + string.Join(" | ", s.Blocks.Where(b => b.Changed).Select(b => $"T style=\"{(string)b.Draft.Element(One + "T").Attribute("style")}\" {b.Draft.Element(One + "T").Value}")));
        Execute(t, "finish_edit", new { snapshot_id = s.SnapshotId, draft_revision = s.Revision });
        Console.WriteLine("  提交：" + t.Report.Status + " " + t.Report.Message);
        var after = AgentPageSnapshot.ParsePage(Read(api, page, directory, "2-after.xml"));
        var kept = true;
        foreach (var oe in Body(after))
        {
            var highlighted = Highlighted.IsMatch(new AgentRichText(oe).Signature(after, true));
            Console.WriteLine($"  {AgentCode.PlainText(oe)}：高亮{(highlighted ? "保留" : "丢失")}；T style=\"{(string)oe.Element(One + "T").Attribute("style")}\"；HTML={oe.Element(One + "T").Value}");
            kept &= highlighted;
        }
        var undo = new AgentCommitter(access).Undo(page, t.Report, s.Options, CancellationToken.None);
        var restored = AgentPageSnapshot.ParsePage(Read(api, page, directory, "2-after-undo.xml"));
        Console.WriteLine("  撤销：" + undo.Status + " " + undo.Message);
        var same = SameFormats(before, restored);
        Console.WriteLine("  撤销后格式与执行前一致：" + same);
        return t.Report.Status == "Verified" && kept && undo.Status == "Verified" && same;
    }

    /// <summary>
    /// 文本框里有几种空段落（空 T、带样式的空 T、段落带样式的空 T、&amp;nbsp;），只改别的段落。
    /// 按段落补丁和整框替换各写一次，空段落都要通过「未指定段落格式不变」的核对。
    /// </summary>
    private static bool BlankParagraphs(IApplication app, OneNoteApi api, string section, string directory)
    {
        var page = Create(app, api, section, "空段落", Outline(100,
            Line("文字一"), Line(""), Styled(Line(""), "font-size:20pt;color:#C00000"), Styled(Line(""), "font-size:20pt", true),
            Styled(Line("&nbsp;"), "font-size:20pt"), Line("文字二")));
        var beforeXml = Read(api, page, directory, "3-before.xml"); var before = AgentPageSnapshot.ParsePage(beforeXml);
        var access = new Logged(api, directory, "3");
        foreach (var oe in Lines(before).Where(e => AgentCode.PlainText(e).Length == 0))
            Console.WriteLine("  回读的空段落：" + oe.ToString(SaveOptions.DisableFormatting));
        // 按段落补丁：改文字一的格式，整个文本框随之写回。
        var s1 = new AgentPageSnapshot(beforeXml, null, new AgentOptions());
        var t1 = new AgentTools(s1, new AgentCommitter(access), CancellationToken.None);
        Execute(t1, "read_blocks", new { snapshot_id = s1.SnapshotId, block_ids = s1.Blocks.Where(b => b.Editable).Select(b => b.Id).ToArray() });
        Execute(t1, "set_text_style", new { snapshot_id = s1.SnapshotId, targets = new[] { new { block_id = s1.Blocks.Single(b => b.Text == "文字一").Id, quote = "文字", occurrence = 1, style = new { bold = true } } } });
        Execute(t1, "finish_edit", new { snapshot_id = s1.SnapshotId, draft_revision = s1.Revision });
        Console.WriteLine("  按段落提交：" + t1.Report.Status + " " + t1.Report.Message);
        var middleXml = Read(api, page, directory, "3-after-patch.xml"); var middle = AgentPageSnapshot.ParsePage(middleXml);
        var blanksKept1 = SameBlanks(before, middle);
        // 整框替换：在文字二后面插入一段。
        var s2 = new AgentPageSnapshot(middleXml, null, new AgentOptions());
        var t2 = new AgentTools(s2, new AgentCommitter(access), CancellationToken.None);
        Execute(t2, "read_blocks", new { snapshot_id = s2.SnapshotId, block_ids = s2.Blocks.Where(b => b.Editable).Select(b => b.Id).ToArray() });
        Execute(t2, "insert_blocks", new { snapshot_id = s2.SnapshotId, target_id = s2.Blocks.Single(b => b.Text == "文字二").Id, position = "after",
            paragraphs = new[] { new { text = "插入的段落", preset_id = "body" } } });
        Execute(t2, "finish_edit", new { snapshot_id = s2.SnapshotId, draft_revision = s2.Revision });
        Console.WriteLine("  整框提交：" + t2.Report.Status + " " + t2.Report.Message);
        var after = AgentPageSnapshot.ParsePage(Read(api, page, directory, "3-after-outline.xml"));
        var blanksKept2 = SameBlanks(middle, after);
        var undo2 = new AgentCommitter(access).Undo(page, t2.Report, s2.Options, CancellationToken.None);
        Console.WriteLine("  撤销整框：" + undo2.Status + " " + undo2.Message);
        var middleRestored = AgentPageSnapshot.ParsePage(Read(api, page, directory, "3-after-undo-outline.xml"));
        var undo1 = new AgentCommitter(access).Undo(page, t1.Report, s1.Options, CancellationToken.None);
        Console.WriteLine("  撤销按段落：" + undo1.Status + " " + undo1.Message);
        var restored = AgentPageSnapshot.ParsePage(Read(api, page, directory, "3-after-undo-patch.xml"));
        var blanksKept3 = SameBlanks(after, middleRestored) && SameBlanks(middleRestored, restored);
        return t1.Report.Status == "Verified" && t2.Report.Status == "Verified" && undo1.Status == "Verified" && undo2.Status == "Verified" &&
            blanksKept1 && blanksKept2 && blanksKept3 && SameFormats(before, restored);
    }

    /// <summary>「清除所有格式，包括代码外层的框」：清除格式、拆代码框、表格恢复默认外观，一次提交并撤销。</summary>
    private static bool ClearEverything(IApplication app, OneNoteApi api, string section, string directory)
    {
        var rich = Styled(Line("<span style='font-weight:bold;color:#C00000;background:yellow'>重点</span>和<a href='https://example.com/x'><span style='font-style:italic'>链接</span></a>&nbsp;x<sup>2</sup>"),
            "font-family:Calibri;font-size:20pt;font-weight:bold", true);
        rich.SetAttributeValue("alignment", "center");
        rich.AddFirst(new XElement(One + "Tag", new XAttribute("index", 0), new XAttribute("completed", "false"), new XAttribute("disabled", "false"),
            new XAttribute("creationDate", DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.000Z"))),
            new XElement(One + "List", new XElement(One + "Bullet", new XAttribute("bullet", 2))));
        var box = CodeBlockBuilder.BuildTable("def f(x):\n    return x + 1\n\nprint(f(1))", LanguageRegistry.Find("python"), CodeThemes.Light, new AddInSettings { FontFamily = "Consolas" });
        var grid = new XElement(One + "Table", new XAttribute("bordersVisible", "false"), new XAttribute("hasHeaderRow", "true"),
            new XElement(One + "Columns", new XElement(One + "Column", new XAttribute("index", 0), new XAttribute("width", 120)),
                new XElement(One + "Column", new XAttribute("index", 1), new XAttribute("width", 120))),
            new XElement(One + "Row", Cell("名称", "#DEEAF6"), Cell("说明", "#DEEAF6")), new XElement(One + "Row", Cell("甲", "#FFF2CC"), Cell("乙", "#E2EFDA")));
        var page = Create(app, api, section, "<span style='color:#FF0000;font-weight:bold'>清除格式</span>", Outline(100,
            rich, Line("运行 <span style='font-family:Consolas'>npm i</span> 即可"), Line("<span style='font-size:30pt'>&nbsp;</span>"),
            Styled(Line("x = 1"), "font-family:Consolas;font-size:10pt", true), new XElement(One + "OE", box), new XElement(One + "OE", grid)),
            new XElement(One + "TagDef", new XAttribute("index", 0), new XAttribute("type", 0), new XAttribute("symbol", 3),
                new XAttribute("fontColor", "automatic"), new XAttribute("highlightColor", "none"), new XAttribute("name", "待办事项")));
        var beforeXml = Read(api, page, directory, "4-before.xml"); var before = AgentPageSnapshot.ParsePage(beforeXml);
        var access = new Logged(api, directory, "4");
        var s = new AgentPageSnapshot(beforeXml, null, new AgentOptions());
        foreach (var b in s.Blocks) Console.WriteLine($"  {b.Id}: {b.ProtectedReason ?? "editable"} {AgentCode.PlainText(b.Original)}");
        var t = new AgentTools(s, new AgentCommitter(access), CancellationToken.None);
        Execute(t, "read_blocks", new { snapshot_id = s.SnapshotId, block_ids = s.Blocks.Where(b => b.Editable || b.CodeCandidate).Select(b => b.Id).ToArray() });
        Console.WriteLine("  clear_format：" + Json(Execute(t, "clear_format", new { snapshot_id = s.SnapshotId, block_ids = s.Blocks.Select(b => b.Id).ToArray() })));
        var code = s.Tables.Single(x => x.ProtectedReason == "highlighted_code").Id;
        var table = s.Tables.Single(x => x.Editable).Id;
        Console.WriteLine("  unwrap_code：" + Json(Execute(t, "unwrap_code", new { snapshot_id = s.SnapshotId, table_ids = new[] { code } })));
        Execute(t, "set_table_style", new { snapshot_id = s.SnapshotId, table_ids = new[] { table }, style = new { borders = true, header_row = false, cell_shading = "none" } });
        File.WriteAllText(Path.Combine(directory, "4-expected.xml"), s.CreateDraftPage().ToString());
        Execute(t, "finish_edit", new { snapshot_id = s.SnapshotId, draft_revision = s.Revision });
        var r = t.Report;
        Console.WriteLine("  提交：" + r.Status + " " + r.Message);
        var after = AgentPageSnapshot.ParsePage(Read(api, page, directory, "4-after.xml"));
        var lines = Lines(after).ToList();
        var noCodeBox = !after.Descendants(One + "Table").Any(x => AgentPageSnapshot.IsCodeBox(x, after));
        var noMono = lines.Where(e => e.Ancestors(One + "Title").Count() == 0).All(e => !new AgentRichText(e).Monospace(after).Any);
        var noHighlight = lines.All(e => !Highlighted.IsMatch(new AgentRichText(e).Signature(after, true)));
        var noMarks = !after.Descendants(One + "Outline").Descendants(One + "List").Any() && !after.Descendants(One + "Outline").Descendants(One + "Tag").Any();
        var noLinks = lines.All(e => new AgentRichText(e).LinkCount == 0);
        var noShading = after.Descendants(One + "Table").Where(x => !AgentPageSnapshot.IsCodeBox(x, after)).Descendants(One + "Cell").All(c => TableLook.Shade((string)c.Attribute("shadingColor")).Length == 0);
        var sameText = Texts(before) == Texts(after);
        Console.WriteLine($"  代码框已拆：{noCodeBox}；无等宽字体：{noMono}；无高亮：{noHighlight}；无列表标记：{noMarks}；无链接：{noLinks}；表格无底色：{noShading}；文字不变：{sameText}");
        foreach (var e in lines.Where(e => AgentCode.PlainText(e).StartsWith(" ", StringComparison.Ordinal) || AgentCode.PlainText(e).Length == 0).Take(3))
            Console.WriteLine("  拆出或空白的行：" + e.ToString(SaveOptions.DisableFormatting));
        var undo = new AgentCommitter(access).Undo(page, r, s.Options, CancellationToken.None);
        Console.WriteLine("  撤销：" + undo.Status + " " + undo.Message);
        var restored = AgentPageSnapshot.ParsePage(Read(api, page, directory, "4-after-undo.xml"));
        var boxBack = restored.Descendants(One + "Table").Any(x => AgentPageSnapshot.IsCodeBox(x, restored));
        var sameAfterUndo = Texts(before) == Texts(restored) && SameFormats(before, restored);
        Console.WriteLine($"  撤销后代码框回来：{boxBack}；文字和格式与执行前一致：{sameAfterUndo}");
        return r.Status == "Verified" && r.Unwrapped == 1 && r.LinksRemoved == 1 && noCodeBox && noMono && noHighlight && noMarks && noLinks && noShading && sameText &&
            undo.Status == "Verified" && boxBack && sameAfterUndo;
    }

    private static string Create(IApplication app, OneNoteApi api, string section, string title, XElement outline, params XElement[] leading)
    {
        app.CreateNewPage(section, out var pageId, NewPageStyle.npsBlankPageWithTitle);
        var initial = AgentPageSnapshot.ParsePage(api.GetPageContent(pageId, PageInfo.piBasic));
        var page = new XElement(One + "Page", new XAttribute("ID", pageId), leading, new XElement(One + "Title", Line(title)), outline);
        api.UpdatePageContent(page.ToString(SaveOptions.DisableFormatting), AgentPageSnapshot.Modified(initial));
        return pageId;
    }

    /// <summary>回读页面并存档。快照要用 OneNote 原样的 XML：重新序列化会加入缩进空白，改变段落指纹。</summary>
    private static string Read(OneNoteApi api, string pageId, string directory, string file)
    {
        var xml = api.GetPageContent(pageId, PageInfo.piBasic);
        File.WriteAllText(Path.Combine(directory, file), xml);
        return xml;
    }

    /// <summary>转发给 OneNote，另把发送的 XML 存档、把写入异常打印出来：提交器会把异常当作「结果不确定」，不显示原因。</summary>
    private sealed class Logged : IOneNotePageAccess
    {
        private readonly OneNoteApi _api; private readonly string _directory, _prefix; private int _writes;
        internal Logged(OneNoteApi api, string directory, string prefix) { _api = api; _directory = directory; _prefix = prefix; }
        public string GetPageContent(string pageId, PageInfo info) => _api.GetPageContent(pageId, info);
        public void UpdatePageContent(string xml, DateTime expected)
        {
            File.WriteAllText(Path.Combine(_directory, _prefix + "-update-" + ++_writes + ".xml"), xml);
            try { _api.UpdatePageContent(xml, expected); }
            catch (Exception ex) { Console.WriteLine("  UpdatePageContent 失败：0x" + ex.HResult.ToString("X8") + " " + ex.Message); throw; }
        }
        public void DeletePageContent(string pageId, string objectId, DateTime expected) => ((IOneNotePageAccess)_api).DeletePageContent(pageId, objectId, expected);
    }

    /// <summary>
    /// 文字段落按页面顺序的语义格式一致（撤销后对照执行前）。重建的段落 ID 会变，按位置对应；
    /// 代码框里的行和插件核验一样跳过空白比较，OneNote 会改写代码行的硬空格。
    /// </summary>
    private static bool SameFormats(XElement left, XElement right)
    {
        string Format(XElement e, XElement page) => e.Ancestors(One + "Table").Any(x => AgentPageSnapshot.IsCodeBox(x, page))
            ? AgentPageSnapshot.CodeLineFormat(e, page) : AgentPageSnapshot.SemanticFormat(e, page);
        var a = Lines(left).Select(e => Format(e, left)).ToList();
        var b = Lines(right).Select(e => Format(e, right)).ToList();
        for (var i = 0; i < Math.Min(a.Count, b.Count); i++)
        {
            if (a[i] == b[i]) continue;
            Console.WriteLine($"  第 {i + 1} 段格式不同：{AgentCode.PlainText(Lines(left).ElementAt(i))}");
            var at = 0; while (at < Math.Min(a[i].Length, b[i].Length) && a[i][at] == b[i][at]) at++;
            string Around(string s) => s.Substring(Math.Max(0, at - 60), Math.Min(s.Length - Math.Max(0, at - 60), 140));
            Console.WriteLine("    之前：…" + Around(a[i])); Console.WriteLine("    之后：…" + Around(b[i]));
        }
        return a.SequenceEqual(b);
    }

    /// <summary>离线比较两份已存档的页面 XML（不连接 OneNote）。</summary>
    internal static int Compare(string before, string after) =>
        SameFormats(AgentPageSnapshot.ParsePage(File.ReadAllText(before)), AgentPageSnapshot.ParsePage(File.ReadAllText(after))) ? 0 : 1;

    /// <summary>两次回读之间，同一 objectID 的空段落语义格式（含空 T 的样式）不变。</summary>
    private static bool SameBlanks(XElement left, XElement right)
    {
        var ok = true;
        foreach (var oe in Lines(left).Where(e => AgentCode.PlainText(e).Length == 0))
        {
            var id = (string)oe.Attribute("objectID");
            var other = AgentCommitter.Find(right, id);
            var same = other != null && AgentPageSnapshot.SemanticFormat(oe, left) == AgentPageSnapshot.SemanticFormat(other, right);
            if (!same) Console.WriteLine("  空段落格式变化：" + id + " → " + other?.ToString(SaveOptions.DisableFormatting));
            ok &= same;
        }
        Console.WriteLine("  空段落格式不变：" + ok);
        return ok;
    }

    private static IEnumerable<XElement> Lines(XElement page) => page.Descendants(One + "OE").Where(e => e.Elements(One + "T").Any());
    /// <summary>文本框里的文字段落，不含页面标题。</summary>
    private static IEnumerable<XElement> Body(XElement page) => page.Elements(One + "Outline").Descendants(One + "OE").Where(e => e.Elements(One + "T").Any());
    private static string Texts(XElement page) => string.Join("|", Lines(page).Select(AgentCode.PlainText));
    private static XElement Outline(int y, params XElement[] paragraphs) => new XElement(One + "Outline",
        new XElement(One + "Position", new XAttribute("x", 36), new XAttribute("y", y), new XAttribute("z", 0)),
        new XElement(One + "Size", new XAttribute("width", 500), new XAttribute("height", 300)), new XElement(One + "OEChildren", paragraphs));
    private static XElement Line(string html) => new XElement(One + "OE", new XElement(One + "T", new XCData(html)));
    /// <summary>给段落的 T（paragraph 为 true 时给 OE）加 style。</summary>
    private static XElement Styled(XElement oe, string style, bool paragraph = false)
    {
        (paragraph ? oe : oe.Element(One + "T")).SetAttributeValue("style", style);
        return oe;
    }
    private static XElement Cell(string text, string shading) => new XElement(One + "Cell", new XAttribute("shadingColor", shading),
        new XElement(One + "OEChildren", Line(text)));
    private static object Execute(AgentTools tools, string name, object args) =>
        tools.Execute(new AgentToolCall { Id = name, Name = name, Arguments = AgentChatClient.Serializer().Serialize(args) });
    private static string Json(object value) => AgentJson.Serialize(value);
}
