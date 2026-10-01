using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OneNoteCodeHelper.Services.Agent
{
    /// <summary>只规范化代码框与文字的交界。段内删行记录在结构草稿中，合成格式草稿后再应用。</summary>
    internal static class AgentCodeSpacing
    {
        private static XNamespace One => OneNoteApi.One;
        internal static readonly XName Leading = AgentLayout.Key.Namespace + "trim-leading";
        internal static readonly XName Trailing = AgentLayout.Key.Namespace + "trim-trailing";

        internal sealed class Skip
        {
            internal string TextId, CodeId, Reason;
            internal object ToResult() => new { text_id = TextId, code_id = CodeId, reason = Reason };
        }

        internal sealed class Result
        {
            internal XElement Layout;
            internal readonly List<AgentLayoutChange> Changes = new List<AgentLayoutChange>();
            internal readonly List<AgentInserted> Inserted = new List<AgentInserted>();
            internal readonly List<Skip> Skipped = new List<Skip>();
            internal int Removed, SoftLines;
        }

        private sealed class Item
        {
            internal XElement Node;
            internal string Kind, Text;
            /// <summary>非空白文字全是等宽字体：未转换的代码，和代码框之间不算文字交界。</summary>
            internal bool Monospace;
            internal AgentBlock Block;
            internal AgentCodeConversion Conversion;
            internal string Id => AgentLayout.KeyOf(Node) ?? (string)Node.Attribute("objectID") ?? "";
        }

        internal static bool HasTrim(XElement oe) => oe.Attribute(Leading) != null || oe.Attribute(Trailing) != null;
        internal static int TrimCount(XElement oe) => Count(oe, Leading) + Count(oe, Trailing);
        private static int Count(XElement oe, XName name) => (int?)oe.Attribute(name) ?? 0;

        internal static void Apply(XElement oe, XElement record = null)
        {
            record = record ?? oe;
            if (TrimCount(record) == 0) return;
            var rich = new AgentRichText(oe);
            var (start, end) = Visible(rich.Text, record);
            rich.Keep(start, end - start);
        }

        /// <summary>按 record 上的删行记录截取后留下的范围 [Start, End)，即模型读到的文字；没有记录时是全文。</summary>
        internal static (int Start, int End) Visible(string text, XElement record)
        {
            if (record == null || TrimCount(record) == 0) return (0, text.Length);
            var start = Prefix(text, Count(record, Leading)); var end = Suffix(text, Count(record, Trailing));
            if (start > end) throw new AiException("代码框间隔的段内空行记录已失效，请重新规范化间隔。");
            return (start, end);
        }

        private static int Prefix(string text, int count)
        {
            var at = 0;
            for (var i = 0; i < count; i++)
            {
                var next = text.IndexOf('\n', at);
                if (next < 0 || !string.IsNullOrWhiteSpace(text.Substring(at, next - at)))
                    throw new AiException("代码框间隔只能删除完整空白行。");
                at = next + 1;
            }
            return at;
        }

        private static int Suffix(string text, int count)
        {
            var at = text.Length;
            for (var i = 0; i < count; i++)
            {
                var next = at == 0 ? -1 : text.LastIndexOf('\n', at - 1);
                if (next < 0 || !string.IsNullOrWhiteSpace(text.Substring(next + 1, at - next - 1)))
                    throw new AiException("代码框间隔只能删除完整空白行。");
                at = next;
            }
            return at;
        }

        private static int EdgeLines(string text, bool leading)
        {
            var lines = text.Split('\n'); var count = 0;
            for (var i = leading ? 0 : lines.Length - 1; i >= 0 && i < lines.Length && count < lines.Length - 1; i += leading ? 1 : -1)
            {
                if (!string.IsNullOrWhiteSpace(lines[i])) break;
                count++;
            }
            return count;
        }

        private static IEnumerable<XElement> Lines(XElement flow)
        {
            foreach (var oe in flow.Elements(One + "OE"))
            {
                yield return oe;
                foreach (var child in oe.Elements(One + "OEChildren").SelectMany(Lines)) yield return child;
            }
        }

        private static bool IsBox(XElement table, XElement page) => AgentPageSnapshot.IsCodeBox(table, page) &&
            table.Descendants(One + "OE").Any(oe => !string.IsNullOrWhiteSpace(PageEditor.ExtractPlainText(oe)));

        private static List<Item> Items(XElement flow, XElement page, AgentPageSnapshot snapshot)
        {
            var blocks = snapshot.Blocks.ToDictionary(b => b.Id);
            var conversions = snapshot.CodeConversions.SelectMany(c => c.Blocks.Select(b => new { b.Id, Conversion = c })).ToDictionary(x => x.Id, x => x.Conversion);
            var seen = new HashSet<AgentCodeConversion>();
            var items = new List<Item>();
            foreach (var oe in Lines(flow))
            {
                var key = AgentLayout.KeyOf(oe) ?? "";
                blocks.TryGetValue(key, out var block);
                if (conversions.TryGetValue(key, out var conversion))
                {
                    if (seen.Add(conversion)) items.Add(new Item { Node = oe, Kind = conversion.TextTable ? "other" : "code", Conversion = conversion });
                    continue;
                }
                var item = new Item { Node = oe, Block = block, Kind = "other" };
                var table = oe.Element(One + "Table");
                if (table != null && !oe.Elements(One + "T").Any() && IsBox(table, page)) item.Kind = "code";
                else if (oe.Elements().All(e => e.Name == One + "T" || e.Name == One + "Meta" || e.Name == One + "List" || e.Name == One + "Tag" || e.Name == One + "OEChildren"))
                {
                    try
                    {
                        var rich = new AgentRichText(oe);
                        item.Text = rich.Text;
                        if (string.IsNullOrWhiteSpace(item.Text)) item.Kind = "blank";
                        else
                        {
                            item.Monospace = rich.Monospace(page).All;
                            if (block?.Editable == true && !item.Monospace) item.Kind = "text";
                            else if (snapshot.Inserted.Any(n => n.Id == key && n.Text.Length > 0)) item.Kind = "text";
                        }
                    }
                    catch (Exception ex) when (ex is AiException || ex is System.Xml.XmlException || ex is ArgumentException) { }
                }
                items.Add(item);
            }
            return items;
        }

        private static bool PlainBlank(Item item) => item.Node.Elements().All(e => e.Name == One + "T" || e.Name == One + "Meta");
        private static bool IdentityKnown(Item item, XElement page)
        {
            var id = (string)item.Node.Attribute("objectID");
            return id == null ? item.Block == null && AgentLayout.KeyOf(item.Node) != null :
                id.Length > 0 && page.Descendants(One + "OE").Count(e => (string)e.Attribute("objectID") == id) == 1;
        }

        /// <summary>也用于 finish_edit 的只读复核；发布草稿和插入配额由调用方统一处理。</summary>
        internal static Result Build(AgentPageSnapshot snapshot)
        {
            var result = new Result { Layout = new XElement(snapshot.Layout) };
            var page = snapshot.CreateDraftPage();
            var nodes = page.Descendants(One + "OE").Zip(result.Layout.Descendants(One + "OE"), (draft, layout) => new { draft, layout }).ToDictionary(x => x.draft, x => x.layout);
            XElement LayoutNode(Item item) => nodes[item.Node];
            void Change(string kind, Item item) => result.Changes.Add(new AgentLayoutChange { Kind = kind,
                OutlineId = (string)item.Node.Ancestors(One + "Outline").First().Attribute("objectID"), Ids = new[] { item.Id } });
            void Trim(Item item, bool leading, int lines)
            {
                if (lines == 0) return;
                var node = LayoutNode(item); var name = leading ? Leading : Trailing;
                node.SetAttributeValue(name, Count(node, name) + lines);
                result.SoftLines += lines;
                Change("code_spacing", item);
            }
            foreach (var flow in BlankLines.FlowsOf(page).Where(f => !f.Ancestors(One + "Table").Any(t => IsBox(t, page))))
            {
                var items = Items(flow, page, snapshot);
                var previous = -1;
                for (var i = 0; i < items.Count; i++)
                {
                    if (items[i].Kind == "blank") continue;
                    if (previous < 0) { previous = i; continue; }
                    var left = items[previous]; var right = items[i]; var blanks = items.Skip(previous + 1).Take(i - previous - 1).ToList();
                    previous = i;
                    var blankLines = blanks.Sum(b => b.Text.Count(c => c == '\n') + 1);
                    // 已经正好隔一行：空段落都能识别，加上文字边缘的段内空行合计一行。
                    bool Settled(Item edge) => EdgeLines(edge.Text, edge == right) + blankLines == 1 && blanks.All(b => PlainBlank(b) && IdentityKnown(b, page));
                    if (!(left.Kind == "code" && right.Kind == "text" || left.Kind == "text" && right.Kind == "code"))
                    {
                        // 未转换的等宽代码、已排入的表格转换与代码框之间不处理，和已有表格一样不说明；已经隔一行的边界也不必说明。
                        var protectedText = left.Kind == "code" && right.Kind == "other" ? right : right.Kind == "code" && left.Kind == "other" ? left : null;
                        if (protectedText != null && protectedText.Conversion == null && protectedText.Node.Elements(One + "T").Any() && !protectedText.Monospace &&
                            !(protectedText.Text != null && Settled(protectedText)))
                            result.Skipped.Add(new Skip { TextId = protectedText.Id, CodeId = (left.Kind == "code" ? left : right).Id,
                                Reason = snapshot.InSelection(protectedText.Node) ? "protected_text" : "outside_selection" });
                        continue;
                    }
                    var text = left.Kind == "text" ? left : right; var code = left.Kind == "code" ? left : right;
                    var leading = text == right;
                    var soft = EdgeLines(text.Text, leading);
                    var gap = soft + blankLines;
                    if (Settled(text)) continue;
                    string reason = null;
                    var outline = (string)text.Node.Ancestors(One + "Outline").First().Attribute("objectID");
                    if (!snapshot.EditableOutlines.Contains(outline)) reason = "protected_outline";
                    else if (!IdentityKnown(text, page)) reason = "protected_text";
                    else if (!snapshot.InSelection(text.Node) || !(code.Conversion != null ? code.Conversion.Blocks.All(b => AgentLayout.Find(page, b.Id) != null) : snapshot.SelectedObjectContains(code.Node))
                        || blanks.Any(b => !snapshot.InSelection(b.Node))) reason = "outside_selection";
                    else if (text.Block != null && text.Block.Editable && !text.Block.Read) reason = "unread_text";
                    else if (blanks.Any(b => !PlainBlank(b) || !IdentityKnown(b, page))) reason = "protected_blank";
                    else if (gap > 1 && !snapshot.Options.EnableBlankLineRemoval) reason = "removal_disabled";
                    else if (gap > 1 && soft > 0 && text.Block != null && !text.Block.Editable) reason = "protected_text";
                    else if (gap == 0 && !snapshot.Options.EnableInsert) reason = "insert_disabled";
                    else if (gap == 0 && snapshot.Inserted.Count + result.Inserted.Count >= AgentTools.MaxInserted) reason = "insert_limit";
                    if (reason != null) { result.Skipped.Add(new Skip { TextId = text.Id, CodeId = code.Id, Reason = reason }); continue; }
                    if (gap == 0)
                    {
                        var id = "n" + (snapshot.Inserted.Count + result.Inserted.Count + 1);
                        var inserted = AgentLayout.NewBlankLine(id, snapshot.Options, snapshot.DraftStyles);
                        LayoutNode(right).AddBeforeSelf(inserted);
                        result.Inserted.Add(new AgentInserted { Id = id, OutlineId = outline, Text = "" });
                        result.Changes.Add(new AgentLayoutChange { Kind = "inserted_blank", OutlineId = outline, Ids = new[] { id } });
                    }
                    else
                    {
                        var keep = blanks.FirstOrDefault();
                        Trim(text, leading, soft - (keep == null ? 1 : 0));
                        if (keep != null) Trim(keep, false, keep.Text.Count(c => c == '\n'));
                        foreach (var blank in blanks.Skip(1))
                        {
                            Change("removed", blank);
                            AgentLayout.Detach(LayoutNode(blank));
                            result.Removed++;
                        }
                    }
                }
            }
            return result;
        }

        internal static string SkipMessage(IEnumerable<Skip> skipped) => string.Join("、", skipped.Select(s => ReasonText(s.Reason)).Distinct());
        private static string ReasonText(string reason)
        {
            switch (reason)
            {
                case "outside_selection": return "交界处不完全在选区内";
                case "unread_text": return "文字尚未完整读取";
                case "removal_disabled": return "删除空行能力已关闭";
                case "insert_disabled": return "插入段落能力已关闭";
                case "insert_limit": return "插入段落配额已用完";
                default: return "交界处包含受保护内容";
            }
        }
    }
}
