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
        internal int MaxLength = 40000;
        internal bool NonEmpty;
        internal static AgentSchema Obj(Dictionary<string, AgentSchema> props, params string[] required) => new AgentSchema { Type = "object", Properties = props, Required = required };
        internal static AgentSchema Str(params string[] values) => new AgentSchema { Type = "string", Enum = values.Length == 0 ? null : values };
        internal static AgentSchema Short(int maxLength) => new AgentSchema { Type = "string", MaxLength = maxLength };
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
            if (Type == "string") { value["minLength"] = 1; value["maxLength"] = MaxLength; }
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
                    if (!(value is string s) || s.Length == 0 || s.Length > MaxLength || (Enum != null && !Enum.Contains(s))) throw new AiException(path + " 的字符串无效。");
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
            Register("fix_text", $"修正错别字、同音字、形近字和明显的标点误用：把段落里第 occurrence 处（从1开始）精确原文 quote 换成 replacement。" +
                $"quote 取出错处连同前后一两个字，两者都不超过 {MaxFixChars} 字、不含换行；同一段的多处修正按顺序应用。只改出错的字，不润色、不改写，保留原有格式和链接。",
                AgentSchema.Obj(new Dictionary<string, AgentSchema>
                {
                    ["snapshot_id"] = AgentSchema.Str(),
                    ["fixes"] = AgentSchema.Array(AgentSchema.Obj(new Dictionary<string, AgentSchema>
                    {
                        ["block_id"] = AgentSchema.Str(), ["quote"] = AgentSchema.Short(MaxFixChars), ["occurrence"] = AgentSchema.Num(1, 1000, true),
                        ["replacement"] = AgentSchema.Short(MaxFixChars)
                    }, "block_id", "quote", "occurrence", "replacement"))
                }, "snapshot_id", "fixes"), FixText);
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

        /// <summary>fix_text 每处原文和改后文字的字数上限：够放下错字和前后一两个字，放不下整句改写。</summary>
        internal const int MaxFixChars = 30;

        private static string[] Languages => new[] { LanguageRegistry.AutoDetectId }.Concat(LanguageRegistry.All.Select(l => l.Id)).ToArray();

        /// <summary>工具在进度和步骤列表里的中文名。</summary>
        internal static string DisplayName(string name)
        {
            switch (name)
            {
                case "get_page_overview": return "读取页面概况";
                case "read_blocks": return "读取段落";
                case "set_paragraph_style": return "设置段落样式";
                case "set_text_style": return "设置重点文字样式";
                case "fix_text": return "修正错别字";
                case "highlight_code": return "高亮代码";
                case "get_pending_changes": return "检查格式草稿";
                case "finish_edit": return "写回并验证";
                default: return "校验工具请求";
            }
        }

        /// <summary>
        /// 步骤列表里的一行：工具名加上从参数和结果里摘出的要点（段数、样式、语言）。
        /// 结果 ok 为 false 时算失败，附上工具返回的错误说明（插件自己写的中文，不含笔记正文）。
        /// 参数或结果解析不了时只显示工具名。
        /// </summary>
        internal static (string Text, AgentStepState State) DescribeStep(string name, string arguments, string result)
        {
            var args = TryParse(arguments);
            var outcome = TryParse(result);
            string detail = null;
            switch (name)
            {
                case "get_page_overview":
                    detail = AiClient.Get(outcome, "total") is int total ? $"共 {total} 段" : null;
                    break;
                case "read_blocks":
                    detail = CountOf(args, "block_ids");
                    break;
                case "set_paragraph_style":
                    detail = JoinDetail(PresetName(AiClient.Get(args, "preset_id") as string), CountOf(args, "block_ids"));
                    break;
                case "set_text_style":
                    detail = AiClient.Get(args, "targets") is IList targets ? $"{targets.Count} 处" : null;
                    break;
                case "fix_text":
                    detail = AiClient.Get(args, "fixes") is IList fixes ? $"{fixes.Count} 处" : null;
                    break;
                case "highlight_code":
                    detail = JoinDetail(LanguageName(AiClient.Get(outcome, "language") as string ?? AiClient.Get(args, "language") as string),
                        CountOf(args, "block_ids"));
                    break;
            }
            var text = detail == null ? DisplayName(name) : DisplayName(name) + " · " + detail;
            if (Equals(AiClient.Get(outcome, "ok"), false))
                return (text + "：" + (AiClient.Get(outcome, "error") as string ?? "失败"), AgentStepState.Failed);
            return (text, AgentStepState.Done);
        }

        private static object TryParse(string json)
        {
            try { return string.IsNullOrWhiteSpace(json) ? null : AgentChatClient.Parse(json); }
            catch (AiException) { return null; }
        }
        private static string CountOf(object args, string key) => AiClient.Get(args, key) is IList items ? $"{items.Count} 段" : null;
        private static string JoinDetail(string a, string b) => a == null ? b : b == null ? a : a + " · " + b;
        private static string PresetName(string preset)
        {
            switch (preset)
            {
                case "page_title": return "页面标题";
                case "heading1": return "一级标题";
                case "heading2": return "二级标题";
                case "body": return "正文";
                case "quote": return "引用";
                default: return null;
            }
        }
        private static string LanguageName(string id)
        {
            if (string.IsNullOrEmpty(id)) return null;
            if (string.Equals(id, LanguageRegistry.AutoDetectId, StringComparison.OrdinalIgnoreCase)) return "自动识别";
            return LanguageRegistry.Find(id)?.DisplayName ?? id;
        }
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
                    summary = b.Editable || b.CodeCandidate ? b.CurrentText.Substring(0, Math.Min(80, b.CurrentText.Length)) : null }).ToArray() };
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
            if (writing && b != null && (b.CodeCandidate || b.Conversion != null)) throw new AiException("代码段落不能设置样式或修改文字；用 highlight_code 转换为代码框。");
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
            return new { snapshot_id = _snapshot.SnapshotId, blocks = blocks.Select(b => new { id = b.Id, kind = b.CodeCandidate ? "unhighlighted_code" : "text", text = b.CurrentText, depth = b.Depth,
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
                var start = Locate(rich.Text, quote, Convert.ToInt32(item["occurrence"]));
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

        /// <summary>quote 第 occurrence 次（从 1 开始，不重叠计数）出现的位置。</summary>
        private static int Locate(string text, string quote, int occurrence)
        {
            var start = -quote.Length;
            for (var i = 0; i < occurrence; i++)
            {
                start = text.IndexOf(quote, start + quote.Length, StringComparison.Ordinal);
                if (start < 0) throw new AiException("找不到指定文字或出现序号。");
            }
            return start;
        }

        private static readonly char[] LineBreaks = { '\n', '\r' };

        /// <summary>
        /// 唯一能改文字的工具，只做字词级的小修正：每处原文和改后文字都有字数上限，不能动换行，代码段落不能改。
        /// 和格式工具一样只改草稿，写回时照常检查冲突、回读核验，撤销时把文字一起还原。
        /// </summary>
        private object FixText(IDictionary<string, object> args)
        {
            var drafts = new Dictionary<AgentBlock, XElement>();
            var fixes = new Dictionary<AgentBlock, List<string>>();
            foreach (IDictionary<string, object> item in (IList)args["fixes"])
            {
                var b = Block((string)item["block_id"], true);
                var quote = (string)item["quote"];
                var replacement = (string)item["replacement"];
                if (quote == replacement) throw new AiException("replacement 与 quote 相同，没有要修正的文字。");
                if (quote.IndexOfAny(LineBreaks) >= 0 || replacement.IndexOfAny(LineBreaks) >= 0) throw new AiException("修正的文字不能包含换行。");
                if (!drafts.TryGetValue(b, out var draft)) { drafts[b] = draft = new XElement(b.Draft); fixes[b] = new List<string>(); }
                var text = new AgentRichText(draft).Text;
                var start = Locate(text, quote, Convert.ToInt32(item["occurrence"]));
                new AgentRichText(draft).Replace(start, quote.Length, replacement);
                if (new AgentRichText(draft).Text != text.Substring(0, start) + replacement + text.Substring(start + quote.Length))
                    throw new AiException("修正后的文字与预期不一致。");
                fixes[b].Add($"「{quote}」→「{replacement}」");
            }
            // 全部校验通过后才发布，失败时这个工具没有副作用。
            foreach (var pair in drafts) { pair.Key.Draft = pair.Value; pair.Key.TextFixes.AddRange(fixes[pair.Key]); }
            _snapshot.Revision++;
            return new { ok = true, draft_revision = _snapshot.Revision, changed = drafts.Keys.Select(b => b.Id).ToArray(),
                blocks = drafts.Keys.Select(b => new { id = b.Id, text = b.CurrentText }).ToArray() };
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
            foreach (var b in ordered) { b.Draft = new XElement(b.Original); b.TextFixes.Clear(); b.Conversion = conversion; }
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
            text_fixes = _snapshot.Blocks.Where(b => b.TextFixes.Count > 0).Select(b => new { id = b.Id, fixes = b.TextFixes.ToArray() }).ToArray(),
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
