using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using OneNoteCodeHelper.Highlighting;

namespace OneNoteCodeHelper.Services.Agent
{
    /// <summary>工具 JSON Schema 和本地校验共用一个定义，避免协议和实现漂移。</summary>
    internal sealed class AgentSchema
    {
        internal string Type;
        internal Dictionary<string, AgentSchema> Properties;
        internal string[] Required;
        internal AgentSchema Items;
        internal string[] Enum;
        internal double Min;
        internal double Max = 100000;
        internal int MinItems;
        internal int MaxItems = 100;
        internal bool NonEmpty;
        internal static AgentSchema Obj(Dictionary<string, AgentSchema> props, params string[] required) => new AgentSchema { Type = "object", Properties = props, Required = required };
        internal static AgentSchema Str(params string[] values) => new AgentSchema { Type = "string", Enum = values.Length == 0 ? null : values };
        internal static AgentSchema Num(double min, double max, bool integer = false) => new AgentSchema { Type = integer ? "integer" : "number", Min = min, Max = max };
        internal static AgentSchema Array(AgentSchema item) => new AgentSchema { Type = "array", Items = item, MinItems = 1 };

        internal object Json()
        {
            var value = new Dictionary<string, object> { ["type"] = Type };
            if (Properties != null)
            {
                value["properties"] = Properties.ToDictionary(p => p.Key, p => p.Value.Json());
                value["required"] = Required;
                value["additionalProperties"] = false;
                if (NonEmpty) value["minProperties"] = 1;
            }
            if (Items != null) { value["items"] = Items.Json(); value["minItems"] = MinItems; value["maxItems"] = MaxItems; }
            if (Enum != null) value["enum"] = Enum;
            if (Type == "number" || Type == "integer") { value["minimum"] = Min; value["maximum"] = Max; }
            if (Type == "string") { value["minLength"] = 1; value["maxLength"] = 40000; }
            return value;
        }

        internal void Validate(object value, string path = "arguments")
        {
            switch (Type)
            {
                case "object":
                    if (!(value is IDictionary<string, object> map) || map.Keys.Any(k => !Properties.ContainsKey(k)) || Required.Any(k => !map.ContainsKey(k)) || (NonEmpty && map.Count == 0))
                        throw new AiException(path + " 的字段不符合工具定义。");
                    foreach (var item in map) Properties[item.Key].Validate(item.Value, path + "." + item.Key);
                    return;
                case "array":
                    if (!(value is IList list) || list.Count < MinItems || list.Count > MaxItems) throw new AiException(path + " 的数组长度无效。");
                    foreach (var item in list) Items.Validate(item, path + "[]");
                    return;
                case "string":
                    if (!(value is string s) || s.Length == 0 || s.Length > 40000 || (Enum != null && !Enum.Contains(s))) throw new AiException(path + " 的字符串无效。");
                    return;
                case "boolean":
                    if (!(value is bool)) throw new AiException(path + " 必须是布尔值。");
                    return;
                default:
                    if (!(value is int || value is long || value is decimal || value is double)) throw new AiException(path + " 必须是数字。");
                    var n = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                    if (double.IsNaN(n) || double.IsInfinity(n) || n < Min || n > Max || (Type == "integer" && Math.Floor(n) != n)) throw new AiException(path + " 的数字超出范围。");
                    return;
            }
        }
    }

    internal sealed class AgentTools
    {
        private sealed class Tool
        {
            internal string Description;
            internal AgentSchema Schema;
            internal Func<IDictionary<string, object>, object> Execute;
        }
        private readonly Dictionary<string, Tool> _tools = new Dictionary<string, Tool>();
        private readonly AgentPageSnapshot _snapshot;
        private readonly AgentCommitter _committer;
        private readonly CancellationToken _cancellation;
        /// <summary>代码框的主题、字体、字号等，取自功能区当前设置，和「高亮选中」一致。</summary>
        private readonly AddInSettings _code;
        internal AgentReport Report { get; private set; }
        internal object[] Definitions => _tools.Select(t => (object)new { type = "function", function = new { name = t.Key, description = t.Value.Description, parameters = t.Value.Schema.Json() } }).ToArray();

        internal AgentTools(AgentPageSnapshot snapshot, AgentCommitter committer, CancellationToken cancellation, AddInSettings codeSettings = null)
        {
            _snapshot = snapshot; _committer = committer; _cancellation = cancellation; _code = codeSettings ?? new AddInSettings();
            Register("get_page_overview", "获取当前固定页面的段落摘要、保护范围及样式。每页 100 项；通过 offset 翻页。", AgentSchema.Obj(new Dictionary<string, AgentSchema>
            { ["offset"] = AgentSchema.Num(0, 1000, true) }), Overview);
            Register("read_blocks", "完整读取段落正文及样式。修改前必须调用，不能修改受保护段落。", WithIds(), Read);
            var fields = new Dictionary<string, AgentSchema>
            {
                ["alignment"] = AgentSchema.Str("left", "center", "right"), ["font_family"] = AgentSchema.Str(ParagraphStyles.Fonts),
                ["font_size_pt"] = AgentSchema.Num(8, 32), ["color"] = AgentSchema.Str(ParagraphStyles.Colors)
            };
            if (snapshot.Options.EnableParagraphSpacing)
            { fields["space_before_pt"] = AgentSchema.Num(0, 36); fields["space_after_pt"] = AgentSchema.Num(0, 36); }
            var paragraph = WithIds();
            paragraph.Properties["preset_id"] = AgentSchema.Str(ParagraphStyles.Ids);
            paragraph.Properties["overrides"] = AgentSchema.Obj(fields);
            paragraph.Required = new[] { "snapshot_id", "block_ids", "preset_id" };
            Register("set_paragraph_style", "为完整读取的段落设置标题、正文或引用样式，可受限覆盖；只修改草稿，不改文字。原生标题由能力开关决定。", paragraph, Paragraph);
            var style = AgentSchema.Obj(new Dictionary<string, AgentSchema>
            {
                ["bold"] = new AgentSchema { Type = "boolean" }, ["italic"] = new AgentSchema { Type = "boolean" },
                ["underline"] = new AgentSchema { Type = "boolean" }, ["color"] = AgentSchema.Str(ParagraphStyles.Colors)
            });
            style.NonEmpty = true;
            Register("set_text_style", "按精确文字及出现序号（从1开始）设置局部格式；不改变文字或链接。", AgentSchema.Obj(new Dictionary<string, AgentSchema>
            {
                ["snapshot_id"] = AgentSchema.Str(),
                ["targets"] = AgentSchema.Array(AgentSchema.Obj(new Dictionary<string, AgentSchema>
                {
                    ["block_id"] = AgentSchema.Str(), ["quote"] = AgentSchema.Str(), ["occurrence"] = AgentSchema.Num(1, 1000, true), ["style"] = style
                }, "block_id", "quote", "occurrence", "style"))
            }, "snapshot_id", "targets"), TextStyle);
            if (snapshot.Options.EnableCodeHighlight)
            {
                var code = WithIds();
                code.Properties["block_ids"].MaxItems = 1000;
                code.Properties["language"] = AgentSchema.Str(Languages);
                code.Required = new[] { "snapshot_id", "block_ids", "language" };
                Register("highlight_code", "把同一文本块里连续的代码段落（含中间空行）整体换成插件的高亮代码框；先完整读取有文字的段落。" +
                    "language 为 auto 时自动识别，识别不出会报错，可改用 text。已有代码框和行内代码不要转换。", code, HighlightCode);
            }
            Register("get_pending_changes", "检查草稿修订号、改动和尚未读取的段落。", SnapshotOnly(), Pending);
            var finish = SnapshotOnly();
            finish.Properties["draft_revision"] = AgentSchema.Num(0, 10000, true);
            finish.Required = new[] { "snapshot_id", "draft_revision" };
            Register("finish_edit", "独立调用此工具提交草稿并验证结果；必须传入最新修订号。本任务之后不能继续编辑。", finish, Finish);
        }

        private static string[] Languages => new[] { LanguageRegistry.AutoDetectId }.Concat(LanguageRegistry.All.Select(l => l.Id)).ToArray();
        private static AgentSchema SnapshotOnly() => AgentSchema.Obj(new Dictionary<string, AgentSchema> { ["snapshot_id"] = AgentSchema.Str() }, "snapshot_id");
        private static AgentSchema WithIds()
        {
            var value = SnapshotOnly(); value.Properties["block_ids"] = AgentSchema.Array(AgentSchema.Str()); value.Required = new[] { "snapshot_id", "block_ids" }; return value;
        }
        private void Register(string name, string description, AgentSchema schema, Func<IDictionary<string, object>, object> action) =>
            _tools.Add(name, new Tool { Description = description, Schema = schema, Execute = action });

        internal object Execute(AgentToolCall call)
        {
            _cancellation.ThrowIfCancellationRequested();
            if (_snapshot.Frozen) throw new AiException("草稿已冻结，不能继续操作。");
            if (!_tools.TryGetValue(call.Name, out var tool)) throw new AiException("未知工具。");
            var parsed = AgentChatClient.Parse(call.Arguments);
            tool.Schema.Validate(parsed);
            var args = (IDictionary<string, object>)parsed;
            if (args.TryGetValue("snapshot_id", out var id) && (string)id != _snapshot.SnapshotId) throw new AiException("快照 ID 已失效。");
            return tool.Execute(args);
        }

        private object Overview(IDictionary<string, object> args)
        {
            var offset = args.TryGetValue("offset", out var n) ? Convert.ToInt32(n) : 0;
            return new { snapshot_id = _snapshot.SnapshotId, page_title = _snapshot.Title, scope = _snapshot.SelectionOnly ? "selected_paragraphs" : "page",
                draft_revision = _snapshot.Revision, presets = ParagraphStyles.Ids, native_headings = _snapshot.Options.EnableNativeHeadings,
                paragraph_spacing = _snapshot.Options.EnableParagraphSpacing, code_highlight = _snapshot.Options.EnableCodeHighlight,
                languages = _snapshot.Options.EnableCodeHighlight ? Languages : null,
                total = _snapshot.Blocks.Count, next_offset = offset + 100 < _snapshot.Blocks.Count ? (int?)(offset + 100) : null,
                blocks = _snapshot.Blocks.Skip(offset).Take(100).Select(b => new { id = b.Id, container_id = b.ContainerId, parent_id = b.ParentId, depth = b.Depth,
                    editable = b.Editable, reason = b.ProtectedReason,
                    summary = b.Editable || b.CodeCandidate ? b.Text.Substring(0, Math.Min(80, b.Text.Length)) : null }).ToArray() };
        }
        private List<AgentBlock> Targets(IDictionary<string, object> args)
        {
            var ids = ((IList)args["block_ids"]).Cast<string>().ToList();
            if (ids.Distinct().Count() != ids.Count) throw new AiException("目标段落重复。");
            return ids.Select(id => Block(id, false)).ToList();
        }
        private AgentBlock Block(string id, bool writing)
        {
            var b = _snapshot.Blocks.FirstOrDefault(x => x.Id == id);
            if (writing && b != null && (b.CodeCandidate || b.Conversion != null)) throw new AiException("代码段落不能设置样式；用 highlight_code 转换为代码框。");
            // 待高亮的代码可以读取，供判断范围和语言。
            if (b == null || !(b.Editable || b.CodeCandidate)) throw new AiException("目标不存在或受到保护。");
            if (writing && !b.Read) throw new AiException("请先完整读取目标段落。");
            return b;
        }
        private object Read(IDictionary<string, object> args)
        {
            var blocks = Targets(args);
            var page = _snapshot.CreateDraftPage();
            foreach (var b in blocks) b.Read = true;
            return new { snapshot_id = _snapshot.SnapshotId, blocks = blocks.Select(b => new { id = b.Id, kind = b.CodeCandidate ? "unhighlighted_code" : "text", text = b.Text, depth = b.Depth,
                container_id = b.ContainerId, parent_id = b.ParentId, style = Css.Effective(AgentCommitter.Find(page, b.ObjectId), page),
                runs = b.Draft.Elements(OneNoteApi.One + "T").Select(t => t.Value).ToArray() }).ToArray() };
        }

        private object Paragraph(IDictionary<string, object> args)
        {
            var blocks = Targets(args);
            var preset = (string)args["preset_id"];
            var overrides = args.TryGetValue("overrides", out var raw) ? (IDictionary<string, object>)raw : new Dictionary<string, object>();
            if (overrides.TryGetValue("font_family", out var font) && !FontInstalled((string)font)) throw new AiException("指定字体未安装。");
            var drafts = new Dictionary<AgentBlock, XElement>();
            var styles = new XElement(_snapshot.DraftStyles);
            foreach (var b in blocks)
            {
                Block(b.Id, true);
                if (preset == "page_title" && AgentCommitter.Find(_snapshot.Page, b.ObjectId).Parent?.Name != OneNoteApi.One + "Title") throw new AiException("页面标题样式只能用于原标题。");
                var draft = new XElement(b.Draft);
                ParagraphStyles.Apply(draft, preset, overrides, _snapshot.Options);
                if (_snapshot.Options.EnableNativeHeadings && preset != "page_title")
                    draft.SetAttributeValue("quickStyleIndex", ParagraphStyles.EnsureDefinition(styles, ParagraphStyles.Definition(preset, _snapshot.Options)));
                drafts.Add(b, draft);
            }
            var result = Publish(drafts);
            _snapshot.DraftStyles.ReplaceNodes(styles.Elements().Select(e => new XElement(e)));
            return result;
        }
        internal static bool FontInstalled(string name) => System.Windows.Media.Fonts.SystemFontFamilies.Any(f =>
            Css.Normalize(f.Source) == Css.Normalize(name) || f.FamilyNames.Values.Any(v => Css.Normalize(v) == Css.Normalize(name)));

        private object TextStyle(IDictionary<string, object> args)
        {
            var drafts = new Dictionary<AgentBlock, XElement>();
            foreach (IDictionary<string, object> item in (IList)args["targets"])
            {
                var b = Block((string)item["block_id"], true);
                if (!drafts.TryGetValue(b, out var draft)) drafts[b] = draft = new XElement(b.Draft);
                var rich = new AgentRichText(draft);
                var quote = (string)item["quote"];
                var start = -quote.Length;
                for (var i = 0; i < Convert.ToInt32(item["occurrence"]); i++)
                {
                    start = rich.Text.IndexOf(quote, start + quote.Length, StringComparison.Ordinal);
                    if (start < 0) throw new AiException("找不到指定文字或出现序号。");
                }
                var css = new Dictionary<string, string>();
                foreach (var style in (IDictionary<string, object>)item["style"])
                {
                    switch (style.Key)
                    {
                        case "bold": css["font-weight"] = (bool)style.Value ? "bold" : "normal"; break;
                        case "italic": css["font-style"] = (bool)style.Value ? "italic" : "normal"; break;
                        case "underline": css["text-decoration"] = (bool)style.Value ? "underline" : "none"; break;
                        case "color": css["color"] = (string)style.Value; break;
                    }
                }
                rich.Format(start, quote.Length, css);
            }
            return Publish(drafts);
        }

        private object HighlightCode(IDictionary<string, object> args)
        {
            var ids = ((IList)args["block_ids"]).Cast<string>().ToList();
            if (ids.Distinct().Count() != ids.Count) throw new AiException("目标段落重复。");
            var blocks = ids.Select(id => _snapshot.Blocks.FirstOrDefault(b => b.Id == id) ?? throw new AiException("目标不存在。")).ToList();
            var scheduled = blocks.Select(b => b.Conversion).Where(c => c != null).Distinct().ToList();
            if (scheduled.Count == 1 && blocks.All(b => b.Conversion == scheduled[0]) && scheduled[0].Blocks.Count == blocks.Count)
                return new { ok = true, draft_revision = _snapshot.Revision, changed = new string[0], noop = ids };
            if (scheduled.Count > 0) throw new AiException("部分段落已排入另一个代码框。");
            foreach (var b in blocks)
            {
                // 代码中间的空行一并放进代码框。
                if (!(b.Editable || b.CodeCandidate || b.ProtectedReason == "empty")) throw new AiException("目标受到保护，不能转换为代码框。");
                if (b.ProtectedReason != "empty" && !b.Read) throw new AiException("请先完整读取目标段落。");
            }
            // 先全部校验、生成代码框，再发布草稿；失败时这个工具没有副作用。
            var selection = AgentCode.Select(_snapshot.Page, blocks.Select(b => b.ObjectId).ToList());
            var language = LanguageRegistry.Resolve((string)args["language"], selection.Code)
                ?? throw new AiException("无法自动识别代码语言。请用 language 指定语言；不确定时用 text。");
            var table = CodeBlockBuilder.BuildTable(selection.Code, language, _code.Theme, _code);
            var ordered = selection.Paragraphs.Select(oe => blocks.First(b => b.ObjectId == (string)oe.Attribute("objectID"))).ToList();
            var conversion = new AgentCodeConversion { Blocks = ordered, LanguageId = language.Id, Code = selection.Code, Table = table };
            // 转换后这些段落就不在了，之前给它们排的格式草稿作废。
            var discarded = ordered.Where(b => b.Changed).Select(b => b.Id).ToArray();
            foreach (var b in ordered) { b.Draft = new XElement(b.Original); b.Conversion = conversion; }
            _snapshot.CodeConversions.Add(conversion);
            _snapshot.Revision++;
            return new { ok = true, draft_revision = _snapshot.Revision, changed = ordered.Select(b => b.Id).ToArray(), noop = new string[0],
                language = language.Id, discarded_format = discarded };
        }

        private object Publish(Dictionary<AgentBlock, XElement> drafts)
        {
            var changed = new List<string>();
            var noop = new List<string>();
            // 先验证全部段落，再一次性发布草稿。失败时这个工具没有副作用。
            foreach (var pair in drafts)
                if (new AgentRichText(pair.Key.Draft).Signature(_snapshot.Page, false) != new AgentRichText(pair.Value).Signature(_snapshot.Page, false))
                    throw new AiException("格式工具不能改变正文或链接。");
            foreach (var pair in drafts)
            {
                if ((string)pair.Key.Draft.Attribute("quickStyleIndex") == (string)pair.Value.Attribute("quickStyleIndex") &&
                    AgentPageSnapshot.SemanticFormat(pair.Key.Draft, _snapshot.Page) == AgentPageSnapshot.SemanticFormat(pair.Value, _snapshot.Page)) noop.Add(pair.Key.Id);
                else { pair.Key.Draft = pair.Value; changed.Add(pair.Key.Id); }
            }
            if (changed.Count > 0) _snapshot.Revision++;
            return new { ok = true, draft_revision = _snapshot.Revision, changed, noop };
        }
        private object Pending(IDictionary<string, object> args) => new { snapshot_id = _snapshot.SnapshotId, draft_revision = _snapshot.Revision,
            changed = _snapshot.Blocks.Where(b => b.Changed).Select(b => b.Id).ToArray(),
            unread = _snapshot.Blocks.Where(b => b.Editable && !b.Read && b.Conversion == null).Select(b => b.Id).ToArray(),
            code_blocks = _snapshot.CodeConversions.Select(c => new { block_ids = c.Blocks.Select(b => b.Id).ToArray(), language = c.LanguageId }).ToArray(),
            unconverted_code = _snapshot.Blocks.Where(b => b.CodeCandidate && b.Conversion == null).Select(b => b.Id).ToArray(),
            protected_count = _snapshot.Blocks.Count(b => !b.Editable && b.Conversion == null) };
        private object Finish(IDictionary<string, object> args)
        {
            if (Convert.ToInt32(args["draft_revision"]) != _snapshot.Revision) throw new AiException("草稿修订号过期，请先检查待提交修改。");
            _snapshot.Frozen = true;
            Report = _committer.Commit(_snapshot, _cancellation);
            return Report.ToToolResult();
        }
    }
}
