using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OneNoteCodeHelper.Highlighting;
using OneNoteCodeHelper.Highlighting.Themes;
using OneNoteCodeHelper.Services.Agent;

namespace OneNoteCodeHelper.Services.Mcp
{
    internal sealed class McpContent
    {
        private static XNamespace One => OneNoteApi.One;
        internal readonly List<XElement> Nodes = new List<XElement>();
        internal readonly List<string> Warnings = new List<string>();
        internal XElement Styles = new XElement("styles"), Tags = new XElement("tags");

        internal static McpContent Parse(string text, string format, AgentOptions options, AddInSettings settings,
            XElement styles = null, XElement tags = null)
        {
            if (text.Length > AgentTools.MaxInsertedChars) throw new McpFault("content_limit", "每个草稿最多插入 5000 字、50 段。");
            var result = new McpContent();
            if (styles != null) result.Styles = new XElement(styles);
            if (tags != null) result.Tags = new XElement(tags);
            var lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            var stack = new List<(int Indent, XElement Node)>();
            for (var i = 0; i < lines.Length; i++)
            {
                var line = lines[i];
                var fence = format == "markdown" ? Regex.Match(line, @"^\s*(`{3,}|~{3,})([^\s]*)\s*$") : Match.Empty;
                if (fence.Success)
                {
                    if (!options.EnableCodeHighlight) throw new McpFault("capability_disabled", "代码高亮能力已关闭。");
                    var end = i + 1;
                    while (end < lines.Length && !Regex.IsMatch(lines[end], @"^\s*" + Regex.Escape(fence.Groups[1].Value[0].ToString()) + "{" + fence.Groups[1].Length + @",}\s*$")) end++;
                    if (end == lines.Length) { result.Warnings.Add("unclosed_fence_preserved"); result.Nodes.Add(result.Paragraph(line, "body", options, false)); continue; }
                    var code = string.Join("\n", lines.Skip(i + 1).Take(end - i - 1));
                    var language = LanguageRegistry.Find(fence.Groups[2].Value) ?? LanguageRegistry.Detect(code) ?? LanguageRegistry.Find("text");
                    var table = CodeBlockBuilder.BuildTable(code, language, CodeThemes.Find(settings.ThemeId), settings);
                    // BuildTable 默认 TrimEnd；导入必须保留末行空白和空行。
                    var expectedLines = OneNoteHtmlEncoder.BuildLines(code, language.Tokenize(code), CodeThemes.Find(settings.ThemeId), settings.TabWidth);
                    var children = table.Descendants(One + "Cell").First().Element(One + "OEChildren");
                    var prototype = new XElement(children.Elements(One + "OE").First());
                    children.ReplaceNodes(expectedLines.Select(l => { var oe = new XElement(prototype); oe.SetAttributeValue("spaceBefore", null); oe.SetAttributeValue("spaceAfter", null); oe.Element(One + "T").ReplaceNodes(new XCData(l)); return oe; }));
                    result.Nodes.Add(new XElement(One + "OE", table)); i = end; stack.Clear(); continue;
                }
                if (format == "markdown" && i + 1 < lines.Length && line.Contains("|") && Regex.IsMatch(lines[i + 1], @"^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?\s*$"))
                {
                    if (!options.EnableTextTables) throw new McpFault("capability_disabled", "表格插入能力已关闭。");
                    var rows = new List<string[]> { Cells(line) }; i++;
                    while (i + 1 < lines.Length && lines[i + 1].Contains("|") && !string.IsNullOrWhiteSpace(lines[i + 1])) rows.Add(Cells(lines[++i]));
                    var columns = rows.Max(r => r.Length);
                    var table = new XElement(One + "Table", new XAttribute("bordersVisible", "true"), new XAttribute("hasHeaderRow", "true"),
                        new XElement(One + "Columns", Enumerable.Range(0, columns).Select(c => new XElement(One + "Column", new XAttribute("index", c), new XAttribute("width", 100)))),
                        rows.Select(r => new XElement(One + "Row", Enumerable.Range(0, columns).Select(c => new XElement(One + "Cell", new XElement(One + "OEChildren", result.Paragraph(c < r.Length ? r[c] : "", "body", options, true)))))));
                    result.Nodes.Add(new XElement(One + "OE", table)); stack.Clear(); continue;
                }
                var preset = "body"; var value = line; var markdown = format == "markdown";
                var list = markdown ? Regex.Match(line, @"^(\s*)([-+*]|\d+\.)\s+(.*)$") : Match.Empty;
                XElement node;
                if (list.Success)
                {
                    if (!options.EnableLists) throw new McpFault("capability_disabled", "列表能力已关闭。");
                    value = list.Groups[3].Value;
                    var todo = Regex.Match(value, @"^\[([ xX])\]\s*(.*)$");
                    if (todo.Success) { if (!options.EnableTags) throw new McpFault("capability_disabled", "待办标记能力已关闭。"); value = todo.Groups[2].Value; }
                    node = result.Paragraph(value, preset, options, true);
                    var numbered = char.IsDigit(list.Groups[2].Value[0]);
                    AgentMarks.SetList(node, numbered ? "number" : "bullet", numbered ? (int?)int.Parse(list.Groups[2].Value.TrimEnd('.')) : null);
                    if (todo.Success) AgentMarks.SetTag(node, "todo", todo.Groups[1].Value != " ", result.Tags);
                    var indent = list.Groups[1].Value.Replace("\t", "    ").Length;
                    while (stack.Count > 0 && stack.Last().Indent >= indent) stack.RemoveAt(stack.Count - 1);
                    if (stack.Count == 0) result.Nodes.Add(node);
                    else { var parent = stack.Last().Node; if (parent.Element(One + "OEChildren") == null) parent.Add(new XElement(One + "OEChildren")); parent.Element(One + "OEChildren").Add(node); }
                    stack.Add((indent, node));
                }
                else
                {
                    stack.Clear();
                    if (markdown && value.StartsWith("## ")) { preset = "heading2"; value = value.Substring(3); }
                    else if (markdown && value.StartsWith("# ")) { preset = "heading1"; value = value.Substring(2); }
                    else if (markdown && value.StartsWith("> ")) { preset = "quote"; value = value.Substring(2); }
                    else if (markdown && Regex.IsMatch(value, @"^(#{3,}\s|\s{4}\S)|!\[|</?[A-Za-z][^>]*>")) { result.Warnings.Add("unsupported_syntax_preserved"); markdown = false; }
                    result.Nodes.Add(result.Paragraph(value, preset, options, markdown));
                }
            }
            if (result.Nodes.SelectMany(n => n.DescendantsAndSelf(One + "OE")).Count() > AgentTools.MaxInserted)
                throw new McpFault("content_limit", "每个草稿最多插入 50 段（含表格单元格及代码行）。");
            return result;
        }

        private XElement Paragraph(string text, string preset, AgentOptions options, bool markdown)
        {
            if (markdown && Regex.IsMatch(text, @"!\[|</?[A-Za-z][^>]*>")) { Warnings.Add("unsupported_syntax_preserved"); markdown = false; }
            var node = AgentLayout.NewParagraph("import", text, preset, null, options, Styles);
            node.SetAttributeValue(AgentLayout.Key, null);
            node.Element(One + "T").ReplaceNodes(new XCData(markdown ? Inline(text) : OneNoteHtmlEncoder.EncodePlainText(text)));
            return node;
        }
        private static string[] Cells(string line) => Regex.Split(line.Trim().Trim('|'), @"(?<!\\)\|").Select(c => c.Trim().Replace("\\|", "|")).ToArray();
        private static string Inline(string text, int depth = 0)
        {
            if (depth > 8) return OneNoteHtmlEncoder.EncodePlainText(text);
            var result = new StringBuilder();
            for (var i = 0; i < text.Length;)
            {
                if (text[i] == '\\' && i + 1 < text.Length) { result.Append(OneNoteHtmlEncoder.EncodePlainText(text[++i].ToString())); i++; continue; }
                var link = text[i] == '[' ? Regex.Match(text.Substring(i), @"^\[([^\]\n]+)\]\(([^\s)]+)\)") : Match.Empty;
                if (link.Success && Uri.TryCreate(link.Groups[2].Value, UriKind.Absolute, out var uri) && new[] { "http", "https", "onenote", "mailto" }.Contains(uri.Scheme))
                { result.Append("<a href=\"").Append(WebUtility.HtmlEncode(uri.OriginalString)).Append("\">").Append(Inline(link.Groups[1].Value, depth + 1)).Append("</a>"); i += link.Length; continue; }
                var delimiter = text.Substring(i).StartsWith("**") ? "**" : text.Substring(i).StartsWith("__") ? "__" : text[i] == '*' || text[i] == '_' ? text[i].ToString() : null;
                var end = delimiter == null ? -1 : text.IndexOf(delimiter, i + delimiter.Length, StringComparison.Ordinal);
                if (delimiter != null && end > i + delimiter.Length && !(delimiter == "_" && i > 0 && char.IsLetterOrDigit(text[i - 1])))
                { var tag = delimiter.Length == 2 ? "b" : "i"; result.Append('<').Append(tag).Append('>').Append(Inline(text.Substring(i + delimiter.Length, end - i - delimiter.Length), depth + 1)).Append("</").Append(tag).Append('>'); i = end + delimiter.Length; continue; }
                result.Append(OneNoteHtmlEncoder.EncodePlainText(text[i++].ToString()));
            }
            return result.ToString();
        }

        internal static object Apply(AgentPageSnapshot snapshot, AddInSettings settings, string name, IDictionary<string, object> args)
        {
            if (name == "replace_text")
            {
                var block = snapshot.Blocks.FirstOrDefault(b => b.Id == (string)args["block_id"]);
                if (block == null || !block.Editable || block.IsPageTitle || !block.Read || block.Conversion != null || AgentLayout.Find(snapshot.Layout, block.Id) == null)
                    throw new McpFault("protected_target", "请先读取普通可编辑正文段落；标题、代码及混合对象不能用此工具修改。");
                var quote = (string)args["quote"]; var replacement = (string)args["replacement"];
                if (quote.IndexOfAny(new[] { '\r', '\n' }) >= 0 || replacement.IndexOfAny(new[] { '\r', '\n' }) >= 0) throw new McpFault("invalid_arguments", "文字替换不能跨换行。");
                var draft = new XElement(block.Draft); var rich = RichParagraph.Parse(draft);
                if (!rich.IsLossless) throw new McpFault("unsupported_html", "此段落不能无损更新。");
                var at = -quote.Length;
                for (var n = 0; n < Convert.ToInt32(args["occurrence"]); n++) { at = rich.Text.IndexOf(quote, at + quote.Length, StringComparison.Ordinal); if (at < 0) throw new McpFault("quote_not_found", "未找到指定原文及出现序号。"); }
                if (quote == replacement) return new { ok = true, draft_revision = snapshot.Revision, changed = false };
                // 将差异计算限定在已定位片段内，重复文字不能导致未修改字符的格式挪到另一处。
                new AgentRichText(draft).Replace(at, quote.Length, replacement);
                new AgentRichText(draft); block.Draft = draft; block.TextFixes.Add("正文更新"); snapshot.Revision++;
                return new { ok = true, draft_revision = snapshot.Revision, changed = true };
            }
            if (!snapshot.Options.EnableInsert) throw new McpFault("capability_disabled", "插入能力已关闭。");
            if (string.IsNullOrWhiteSpace((string)args["content"])) throw new McpFault("empty_content", "OneNote 不保留全空文本框，请提供正文或使用 insert_blocks 插入空段落。");
            if (name == "append_content" && snapshot.SelectionOnly) throw new McpFault("scope_mismatch", "追加独立文本框只支持整页范围。");
            var content = Parse((string)args["content"], (string)args["format"], snapshot.Options, settings, snapshot.DraftStyles, snapshot.DraftTags);
            var candidate = new XElement(snapshot.Layout);
            string outlineId;
            if (name == "append_content")
            {
                outlineId = "draft-" + Guid.NewGuid().ToString("N");
                var (x, y) = PageEditor.NextFreePosition(candidate);
                candidate.Add(new XElement(One + "Outline", new XAttribute("objectID", outlineId),
                    new XElement(One + "Position", new XAttribute("x", x), new XAttribute("y", y), new XAttribute("z", 0)),
                    new XElement(One + "Size", new XAttribute("width", settings.CodeBlockWidth), new XAttribute("height", 100)),
                    new XElement(One + "OEChildren", content.Nodes)));
            }
            else
            {
                var block = snapshot.Blocks.FirstOrDefault(b => b.Id == (string)args["target_id"]);
                var target = AgentLayout.Find(candidate, (string)args["target_id"]);
                var outline = target?.Ancestors(One + "Outline").FirstOrDefault(); outlineId = (string)outline?.Attribute("objectID");
                if (block == null || !block.Read || !block.Editable || block.Conversion != null || outlineId == null || !snapshot.EditableOutlines.Contains(outlineId))
                    throw new McpFault("protected_target", "插入位置必须是已读取、可调整结构的普通段落。");
                if ((string)args["position"] == "before") target.AddBeforeSelf(content.Nodes); else target.AddAfterSelf(content.Nodes);
            }
            var created = content.Nodes.SelectMany(n => n.DescendantsAndSelf(One + "OE")).ToList();
            if (snapshot.Inserted.Count + created.Count > AgentTools.MaxInserted || snapshot.Inserted.Sum(n => n.Text.Length) + ((string)args["content"]).Length > AgentTools.MaxInsertedChars)
                throw new McpFault("content_limit", "每个草稿最多插入 5000 字、50 段。");
            for (var i = 0; i < created.Count; i++) created[i].SetAttributeValue(AgentLayout.Key, "n" + (snapshot.Inserted.Count + i + 1));
            snapshot.CheckLayout(candidate);
            snapshot.Layout = candidate; snapshot.DraftStyles.ReplaceNodes(content.Styles.Elements()); snapshot.DraftTags.ReplaceNodes(content.Tags.Elements());
            if (name == "append_content") { snapshot.NewOutlines.Add(outlineId); snapshot.EditableOutlines.Add(outlineId); }
            snapshot.LayoutChanges.Add(new AgentLayoutChange { Kind = "inserted", OutlineId = outlineId, Ids = created.Select(AgentLayout.KeyOf).ToArray() });
            snapshot.Inserted.AddRange(created.Select(n => new AgentInserted { Id = AgentLayout.KeyOf(n), OutlineId = outlineId, Text = PageEditor.ExtractPlainText(n) }));
            snapshot.Revision++;
            return new { ok = true, draft_revision = snapshot.Revision, inserted = created.Select(AgentLayout.KeyOf).ToArray(), warnings = content.Warnings.Distinct().ToArray() };
        }
    }
}
