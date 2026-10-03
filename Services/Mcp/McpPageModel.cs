using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OneNoteCodeHelper.Services.Agent;

namespace OneNoteCodeHelper.Services.Mcp
{
    /// <summary>只读模型不沿用编辑器的代码保护过滤；无法解释的对象始终显式返回。</summary>
    internal sealed class McpPageModel
    {
        private static XNamespace One => OneNoteApi.One;
        internal readonly List<Dictionary<string, object>> Blocks = new List<Dictionary<string, object>>();
        internal readonly List<string> Issues = new List<string>();
        internal string PageId, Title, Modified;
        internal bool Complete => Issues.Count == 0;
        internal string Markdown => string.Join("\n", Blocks.Select(b => (string)b["markdown"]));

        internal McpPageModel(XElement page, ISet<string> selection = null)
        {
            PageId = (string)page.Attribute("ID"); Title = (string)page.Attribute("name") ?? "";
            Modified = (string)page.Attribute("lastModifiedTime");
            var snapshot = new AgentPageSnapshot(page.ToString(SaveOptions.DisableFormatting), selection, new AgentOptions { MaxPageChars = int.MaxValue }, false);
            var editable = snapshot.Blocks.Where(b => b.Editable).Select(b => b.ObjectId).ToHashSetCompat();
            foreach (var element in page.Elements()) Visit(element, 0, "", null, page, selection, editable);
        }

        private void Visit(XElement e, int depth, string container, string parent, XElement page, ISet<string> selection, ISet<string> editable)
        {
            var id = (string)e.Attribute("objectID");
            var kind = e.Name.LocalName;
            if (kind == "Outline" || kind == "Title")
            { foreach (var c in e.Elements()) Visit(c, depth, id ?? kind, parent, page, selection, editable); return; }
            if (kind == "OEChildren")
            { foreach (var c in e.Elements()) Visit(c, depth, container, parent, page, selection, editable); return; }
            if (kind == "OE")
            {
                var selected = selection == null || selection.Contains(id ?? "");
                if (selected && e.Elements(One + "T").Any())
                {
                    string text, markdown; var readable = true;
                    try { text = new AgentRichText(e).Text; markdown = InlineMarkdown(e); }
                    catch (Exception) { text = PageEditor.ExtractPlainText(e); markdown = text; readable = false; Issues.Add("unsupported_html"); }
                    var code = PageEditor.IsCodeParagraph(e);
                    var prefix = new string(' ', depth * 2);
                    var list = AgentMarks.ListKind(e);
                    var todo = e.Elements(One + "Tag").FirstOrDefault(t => AgentMarks.KindOf(page, t) == "todo");
                    if (todo != null) prefix += AgentMarks.IsCompleted(todo) ? "- [x] " : "- [ ] ";
                    else if (list == "bullet") prefix += "- ";
                    else if (list == "number") prefix += ((string)e.Element(One + "List")?.Element(One + "Number")?.Attribute("restartNumberingAt") ?? "1") + ". ";
                    var preset = page.Elements(One + "QuickStyleDef").FirstOrDefault(d => (string)d.Attribute("index") == (string)e.Attribute("quickStyleIndex"));
                    var name = (string)preset?.Attribute("name") ?? "";
                    if (container == "Title" || e.Ancestors(One + "Title").Any()) prefix = "# ";
                    else if (name == "h1" || name == "h2" || name.IndexOf("heading", StringComparison.OrdinalIgnoreCase) >= 0 || name.Contains("标题"))
                        prefix += name.Contains("2") ? "## " : "# ";
                    else if (name == "quote") prefix += "> ";
                    var block = Add(code ? "code" : "paragraph", id, parent, container, depth, text,
                        code ? Fence(text) : prefix + markdown, readable, !code && editable.Contains(id ?? ""));
                    block["list"] = list; block["tags"] = AgentMarks.Describe(e, page);
                    block["links"] = Links(e).ToArray();
                    if (!readable) block["reason"] = "unsupported_html";
                }
                foreach (var c in e.Elements().Where(c => c.Name != One + "T" && c.Name != One + "List" && c.Name != One + "Tag" && c.Name != One + "Meta"))
                    if (c.Name == One + "OEChildren") Visit(c, depth + 1, container, id, page, selection, editable);
                    else if (selected) Visit(c, depth, container, id, page, selection, editable);
                return;
            }
            if (kind == "Table")
            {
                if (selection != null && !e.Descendants(One + "OE").Any(o => selection.Contains((string)o.Attribute("objectID") ?? ""))) return;
                // 部分选中的表格逐段返回，避免泄漏选区外单元格。
                if (selection != null && e.Descendants(One + "OE").Where(o => o.Elements(One + "T").Any()).Any(o => !selection.Contains((string)o.Attribute("objectID") ?? "")))
                { foreach (var o in e.Elements(One + "Row").Elements(One + "Cell").Elements(One + "OEChildren")) Visit(o, depth, container, parent, page, selection, editable); return; }
                var rows = e.Elements(One + "Row").Select(r => r.Elements(One + "Cell").Select(c =>
                    string.Join("\n", c.Descendants(One + "OE").Where(o => o.Elements(One + "T").Any()).Select(PageEditor.ExtractPlainText))).ToArray()).ToArray();
                var code = e.Descendants(One + "T").Any() && AgentPageSnapshot.IsCodeBox(e, page);
                var text = string.Join("\n", rows.Select(r => string.Join("\t", r)));
                var markdown = code ? Fence(text) : TableMarkdown(rows);
                var block = Add(code ? "code" : "table", id, parent, container, depth, text, markdown, true, false);
                block["rows"] = rows; block["links"] = e.Descendants(One + "OE").SelectMany(Links).ToArray();
                // 单元格也使用同一读取结构，图片、标记或未知对象不能因表格聚合而消失。
                var cells = e.Elements(One + "Row").Select(r => r.Elements(One + "Cell").Select(c =>
                {
                    var start = Blocks.Count;
                    foreach (var children in c.Elements(One + "OEChildren")) Visit(children, depth + 1, container, id ?? parent, page, selection, editable);
                    var contents = Blocks.Skip(start).ToArray(); Blocks.RemoveRange(start, Blocks.Count - start);
                    if (contents.Any(b => !(bool)b["readable"])) block["readable"] = false;
                    return contents;
                }).ToArray()).ToArray();
                block["cells"] = cells;
                if (!code) block["markdown"] = TableMarkdown(cells.Select(r => r.Select(c => string.Join("\n", c.Select(b => (string)b["markdown"]))).ToArray()).ToArray());
                return;
            }
            if (kind == "Image")
            {
                if (selection != null && !selection.Contains(id ?? "") && !selection.Contains(parent ?? "") && (string)e.Attribute("selected") != "all") return;
                var ocr = e.Element(One + "OCRData")?.Element(One + "OCRText")?.Value ?? "";
                var block = Add("image", id ?? parent, parent, container, depth, ocr, "[图片]" + (ocr.Length == 0 ? "" : "\n" + ocr), true, false);
                block["ocr"] = ocr; Issues.Add("image_placeholder"); return;
            }
            if (new[] { "Position", "Size", "Meta", "QuickStyleDef", "TagDef", "PageSettings", "OutlookTask", "Tag" }.Contains(kind)) return;
            Issues.Add("unsupported_object:" + kind);
            Add("object", id ?? parent, parent, container, depth, "", "[未解析对象：" + kind + "]", false, false)["reason"] = "unsupported_object:" + kind;
        }

        private Dictionary<string, object> Add(string type, string id, string parent, string container, int depth, string text, string markdown, bool readable, bool editable)
        {
            var block = new Dictionary<string, object> { ["type"] = type, ["object_id"] = id, ["parent_id"] = parent,
                ["container_id"] = container, ["depth"] = depth, ["text"] = text, ["markdown"] = markdown,
                ["readable"] = readable, ["editable"] = editable };
            Blocks.Add(block); return block;
        }
        private static string Fence(string code)
        { var fence = new string('`', Math.Max(3, Regex.Matches(code, "`+").Cast<Match>().Select(m => m.Length + 1).DefaultIfEmpty(3).Max())); return fence + "\n" + code + "\n" + fence; }
        private static string TableMarkdown(string[][] rows)
        {
            if (rows.Length == 0) return "";
            string Row(string[] cells) => "| " + string.Join(" | ", cells.Select(c => c.Replace("|", "\\|").Replace("\n", "<br>"))) + " |";
            return Row(rows[0]) + "\n" + Row(rows[0].Select(_ => "---").ToArray()) + (rows.Length < 2 ? "" : "\n" + string.Join("\n", rows.Skip(1).Select(Row)));
        }
        private static IEnumerable<object> Links(XElement oe) => oe.Elements(One + "T").SelectMany(t => Regex.Matches(t.Value,
            "<a\\b[^>]*href\\s*=\\s*(?:\"(?<u>[^\"]*)\"|'(?<u>[^']*)')[^>]*>(?<t>.*?)</a>", RegexOptions.IgnoreCase | RegexOptions.Singleline).Cast<Match>())
            .Select(m => (object)new { url = WebUtility.HtmlDecode(m.Groups["u"].Value), text = WebUtility.HtmlDecode(Regex.Replace(m.Groups["t"].Value, "<[^>]*>", "")) });
        private static string InlineMarkdown(XElement oe)
        {
            var html = string.Concat(oe.Elements(One + "T").Select(t => t.Value));
            html = Regex.Replace(html, "<a\\b[^>]*href\\s*=\\s*(?:\"(?<u>[^\"]*)\"|'(?<u>[^']*)')[^>]*>(?<t>.*?)</a>", m => "[" + m.Groups["t"].Value + "](" + m.Groups["u"].Value + ")", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            html = Regex.Replace(html, "</?(?:b|strong)\\b[^>]*>", "**", RegexOptions.IgnoreCase);
            html = Regex.Replace(html, "</?(?:i|em)\\b[^>]*>", "*", RegexOptions.IgnoreCase);
            html = Regex.Replace(html, "<br\\s*/?>", "\n", RegexOptions.IgnoreCase);
            return WebUtility.HtmlDecode(Regex.Replace(html, "<[^>]*>", "")).Replace('\u00a0', ' ');
        }

        // 按语义核对而非原始 XML：忽略回存的 ID、时间、自动高度和定义编号；逐段检查格式、表格及二进制数据。
        internal static string SignatureContent(XElement page, bool identity)
        {
            var data = new StringBuilder();
            foreach (var e in page.Descendants().Where(e => new[] { "Title", "Outline", "OE", "Table", "Row", "Cell", "Image" }.Contains(e.Name.LocalName)))
            {
                data.Append(e.Name.LocalName).Append('/').Append(e.Ancestors().Count()).Append(':');
                if (identity) data.Append((string)e.Attribute("objectID"));
                if (e.Elements(One + "T").Any()) data.Append(AgentPageSnapshot.SemanticFormat(e, page));
                if (e.Name == One + "Table" || e.Name == One + "Cell")
                    data.Append(TableLook.Flag(e, "bordersVisible")).Append(TableLook.Flag(e, "hasHeaderRow")).Append(TableLook.Shade((string)e.Attribute("shadingColor")));
                if (e.Name == One + "Outline") data.Append(e.Element(One + "Position")?.ToString(SaveOptions.DisableFormatting));
                if (e.Name == One + "Image") data.Append(e.Element(One + "Data")?.Value).Append(e.Element(One + "OCRData")?.ToString(SaveOptions.DisableFormatting));
                data.Append('\n');
            }
            data.Append((string)page.Attribute("name"));
            // 未解析对象、页面元数据、列锁定和对象几何也参加指纹；撤销不能删掉后来加入的墨迹或附件。
            var extra = new XElement(page);
            foreach (var e in extra.Descendants(One + "T").ToList()) e.Remove();
            foreach (var e in extra.Descendants(One + "QuickStyleDef").Concat(extra.Descendants(One + "TagDef")).ToList()) e.Remove();
            foreach (var e in extra.DescendantsAndSelf())
            {
                foreach (var a in e.Attributes().Where(a => a.IsNamespaceDeclaration || a.Name.Namespace == AgentLayout.Key.Namespace ||
                    new[] { "ID", "objectID", "selected", "lastModifiedTime", "creationTime", "lastModifiedBy", "author", "authorInitials", "quickStyleIndex", "index", "creationDate", "completionDate" }.Contains(a.Name.LocalName) ||
                    e.Name == One + "Size" && a.Name.LocalName == "height" || e.Name == One + "Column" && a.Name.LocalName == "width" && (string)e.Attribute("isLocked") != "true").ToList()) a.Remove();
                if (e.Name == One + "Tag")
                { var original = page.Descendants(One + "Tag").ElementAt(extra.Descendants(One + "Tag").ToList().IndexOf(e)); var definition = AgentMarks.Definition(page, original); e.SetAttributeValue("definition", definition == null ? null : AgentMarks.Signature(definition)); }
                e.ReplaceAttributes(e.Attributes().OrderBy(a => a.Name.ToString()).ToArray());
            }
            foreach (var n in extra.DescendantNodes().OfType<XText>().Where(t => (t.Parent.HasElements || t.Parent.Name == One + "OE") && string.IsNullOrWhiteSpace(t.Value)).ToList()) n.Remove();
            data.Append(extra.ToString(SaveOptions.DisableFormatting));
            return data.ToString();
        }
        internal static string Signature(XElement page, bool identity)
        { using (var hash = SHA256.Create()) return Convert.ToBase64String(hash.ComputeHash(Encoding.UTF8.GetBytes(SignatureContent(page, identity)))); }
    }
    internal static class McpSetExtensions
    {
        internal static HashSet<T> ToHashSetCompat<T>(this IEnumerable<T> items) => new HashSet<T>(items);
    }
}
