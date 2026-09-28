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
        /// <summary>本次注册了这个工具；系统提示词按实际提供的工具追加说明。</summary>
        internal bool Has(string name) => _tools.ContainsKey(name);
        internal object[] Definitions => _tools.Select(t => (object)new { type = "function", function = new { name = t.Key, description = t.Value.Description, parameters = t.Value.Schema.Json() } }).ToArray();

        internal AgentTools(AgentPageSnapshot snapshot, AgentCommitter committer, CancellationToken cancellation, AddInSettings codeSettings = null)
        {
            _snapshot = snapshot; _committer = committer; _cancellation = cancellation; _code = codeSettings ?? new AddInSettings();
            Register("get_page_overview", "获取当前固定页面的段落摘要、保护范围及样式。每页 100 项；通过 offset 翻页。", AgentSchema.Obj(new Dictionary<string, AgentSchema>
            { ["offset"] = AgentSchema.Num(0, 1000, true) }), Overview);
            Register("read_blocks", "完整读取段落正文及样式。修改前必须调用，不能修改受保护段落。", WithIds(), Read);
            if (snapshot.Images.Count > 0)
            {
                var images = SnapshotOnly();
                images.Properties["image_ids"] = AgentSchema.Array(AgentSchema.Str());
                images.Required = new[] { "snapshot_id", "image_ids" };
                Register("read_image_text", "读取 OneNote 已识别出的图片文字，只用来理解页面；图片本身不能修改。", images, ImageText);
            }
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
            if (snapshot.Options.EnableLists)
            {
                var list = WithIds();
                list.Properties["list"] = AgentSchema.Str(AgentMarks.ListKinds);
                list.Required = new[] { "snapshot_id", "block_ids", "list" };
                Register("set_list", "把完整读取的段落设为项目符号（bullet）、编号（number）列表，或取消列表（none）；只改段落前的符号，不改文字。", list, SetList);
            }
            if (snapshot.Options.EnableTags)
            {
                var tag = WithIds();
                tag.Properties["tag"] = AgentSchema.Str(AgentMarks.TagKinds.Concat(new[] { "none" }).ToArray());
                tag.Properties["completed"] = new AgentSchema { Type = "boolean" };
                tag.Required = new[] { "snapshot_id", "block_ids", "tag" };
                Register("set_tag", "给完整读取的段落加待办（todo）、重要（important）或问题（question）标记；completed 勾选或取消待办，只用于 todo。" +
                    "none 去掉这三种标记，其他标记保留。", tag, SetTag);
            }
            if (snapshot.Options.EnableTableStyles && snapshot.Tables.Any(t => t.Editable))
            {
                var look = AgentSchema.Obj(new Dictionary<string, AgentSchema>
                {
                    ["borders"] = new AgentSchema { Type = "boolean" }, ["header_row"] = new AgentSchema { Type = "boolean" },
                    ["header_shading"] = AgentSchema.Str(TableLook.Shadings.Concat(new[] { "none" }).ToArray())
                });
                look.NonEmpty = true;
                Register("set_table_style", "设置表格边框（borders）、标题行（header_row）和首行底色（header_shading，none 为无底色）；不改单元格文字和行列。",
                    AgentSchema.Obj(new Dictionary<string, AgentSchema>
                    {
                        ["snapshot_id"] = AgentSchema.Str(), ["table_ids"] = AgentSchema.Array(AgentSchema.Str()), ["style"] = look
                    }, "snapshot_id", "table_ids", "style"), TableStyle);
            }
            if (snapshot.Options.EnableCodeHighlight)
            {
                var code = WithIds();
                code.Properties["block_ids"].MaxItems = 1000;
                code.Properties["language"] = AgentSchema.Str(Languages);
                code.Required = new[] { "snapshot_id", "block_ids", "language" };
                Register("highlight_code", "把同一文本块里连续的代码段落（含中间空行）整体换成插件的高亮代码框；先完整读取有文字的段落。" +
                    "language 为 auto 时自动识别，识别不出会报错，可改用 text。已有代码框和行内代码不要转换。", code, HighlightCode);
            }
            var structural = snapshot.EditableOutlines.Count > 0;
            if (snapshot.Options.EnableBlankLineRemoval && snapshot.Blocks.Any(b => b.ProtectedReason == "empty" && snapshot.EditableOutlines.Contains(b.ContainerId)))
            {
                var blank = SnapshotOnly();
                blank.Properties["mode"] = AgentSchema.Str("collapse", "all");
                blank.Required = new[] { "snapshot_id", "mode" };
                Register("remove_blank_lines", "删除多余的空行（只有空白的段落）：collapse 把连续空行合并为一行，并删掉文本框、单元格首尾的空行；all 删掉全部空行。" +
                    "带列表、标记或下级段落的空段落和代码里的空行不删。", blank, RemoveBlankLines);
            }
            if (structural && snapshot.Options.EnableIndent)
            {
                var indent = WithIds();
                indent.Properties["direction"] = AgentSchema.Str("in", "out");
                indent.Required = new[] { "snapshot_id", "block_ids", "direction" };
                Register("set_indent", "调整段落层级：in 挂到上一段下面，out 提到上一级（原来排在它后面的同级段落改挂到它下面）。上下顺序不变，下级段落跟着一起调整。", indent, SetIndent);
            }
            if (structural && snapshot.Options.EnableMoves)
            {
                var move = WithIds();
                move.Properties["target_id"] = AgentSchema.Str();
                move.Properties["position"] = AgentSchema.Str("before", "after");
                move.Required = new[] { "snapshot_id", "block_ids", "target_id", "position" };
                Register("move_blocks", "把段落连同下级段落按原来的先后顺序移到目标段落前面或后面，成为目标的同级段落。可以移到另一个文本框，" +
                    "但表格单元格里的段落只能在同一个单元格里移动，也不能把文本框移空。", move, MoveBlocks);
                if (snapshot.EditableOutlines.Count > 1)
                    Register("merge_outlines", "把一个文本框的全部内容（段落、表格、图片、空行）按原来的顺序移到另一个文本框里目标段落的前面或后面，成为目标的同级段落，" +
                        "再删掉空了的源文本框。source_id 是源文本框的 container_id（见 get_page_overview 的 outlines）。",
                        AgentSchema.Obj(new Dictionary<string, AgentSchema>
                        {
                            ["snapshot_id"] = AgentSchema.Str(), ["source_id"] = AgentSchema.Str(), ["target_id"] = AgentSchema.Str(), ["position"] = AgentSchema.Str("before", "after")
                        }, "snapshot_id", "source_id", "target_id", "position"), MergeOutlines);
            }
            if (structural && snapshot.Options.EnableInsert)
            {
                var item = new Dictionary<string, AgentSchema> { ["text"] = AgentSchema.Short(MaxInsertChars), ["preset_id"] = AgentSchema.Str("heading1", "heading2", "body", "quote") };
                if (snapshot.Options.EnableLists) item["list"] = AgentSchema.Str("bullet", "number");
                var paragraphs = AgentSchema.Array(AgentSchema.Obj(item, "text", "preset_id"));
                paragraphs.MaxItems = MaxInsertPerCall;
                Register("insert_blocks", $"在目标段落前面或后面插入同级的新段落：只写纯文字、不含换行，使用预设样式。每次最多 {MaxInsertPerCall} 段，每段最多 {MaxInsertChars} 字，" +
                    $"每个任务最多 {MaxInserted} 段、{MaxInsertedChars} 字。新段落的 ID 为 n1、n2…，可以作为之后插入、移动的目标。",
                    AgentSchema.Obj(new Dictionary<string, AgentSchema>
                    {
                        ["snapshot_id"] = AgentSchema.Str(), ["target_id"] = AgentSchema.Str(), ["position"] = AgentSchema.Str("before", "after"), ["paragraphs"] = paragraphs
                    }, "snapshot_id", "target_id", "position", "paragraphs"), InsertBlocks);
            }
            if (structural && snapshot.Options.EnableTextTables)
            {
                var table = WithIds();
                table.Properties["block_ids"].MaxItems = 200;
                table.Properties["delimiter"] = AgentSchema.Str(AgentTextTable.Delimiters);
                table.Properties["header_row"] = new AgentSchema { Type = "boolean" };
                table.Properties["borders"] = new AgentSchema { Type = "boolean" };
                table.Properties["header_shading"] = AgentSchema.Str(TableLook.Shadings.Concat(new[] { "none" }).ToArray());
                var header = AgentSchema.Array(AgentSchema.Short(AgentTextTable.MaxHeaderChars));
                header.MaxItems = AgentTextTable.MaxColumns;
                table.Properties["header"] = header;
                table.Required = new[] { "snapshot_id", "block_ids", "delimiter" };
                Register("text_to_table", "把同一文本框里连续的、用制表符（tab）、竖线（pipe）或空格（space）分隔的段落转成表格：每行文字一行（段内 Shift+Enter 换行的也各成一行），" +
                    "中间的空行和 Markdown 分隔行去掉，单元格保留原有文字格式和链接；先完整读取有文字的段落。space 按连续空白拆分。" +
                    "各行列数不一致时按最多的列数建表，缺的单元格留空，结果里的 padded_rows 是补了空单元格的行数。" +
                    "标题等不含分隔符的段落不要放进 block_ids。header_row、borders 默认 true，第一行是数据不是列名时 header_row 设为 false；header_shading 是首行底色。" +
                    "只有用户要求加表头时才用 header 在首行前新增一行列名，个数应等于列数，少了补空，多了按 header 加列。" +
                    $"最多 {AgentTextTable.MaxRows} 行、{AgentTextTable.MaxColumns} 列。", table, TextToTable);
            }
            Register("get_pending_changes", "检查草稿修订号、改动和尚未读取的段落。", SnapshotOnly(), Pending);
            var finish = SnapshotOnly();
            finish.Properties["draft_revision"] = AgentSchema.Num(0, 10000, true);
            finish.Required = new[] { "snapshot_id", "draft_revision" };
            Register("finish_edit", "独立调用此工具提交草稿并验证结果；必须传入最新修订号。本任务之后不能继续编辑。", finish, Finish);
        }

        /// <summary>fix_text 每处原文和改后文字的字数上限：够放下错字和前后一两个字，放不下整句改写。</summary>
        internal const int MaxFixChars = 30;
        /// <summary>insert_blocks 的限额：每次调用、每段、每个任务。</summary>
        internal const int MaxInsertPerCall = 20;
        internal const int MaxInsertChars = 500;
        internal const int MaxInserted = 50;
        internal const int MaxInsertedChars = 5000;

        private static string[] Languages => new[] { LanguageRegistry.AutoDetectId }.Concat(LanguageRegistry.All.Select(l => l.Id)).ToArray();

        /// <summary>工具在进度和步骤列表里的中文名。</summary>
        private static readonly Dictionary<string, string> ToolNames = new Dictionary<string, string>
        {
            ["get_page_overview"] = "读取页面概况",
            ["read_blocks"] = "读取段落",
            ["set_paragraph_style"] = "设置段落样式",
            ["set_text_style"] = "设置重点文字样式",
            ["fix_text"] = "修正错别字",
            ["set_list"] = "设置列表",
            ["set_tag"] = "设置标记",
            ["set_table_style"] = "设置表格样式",
            ["read_image_text"] = "读取图片文字",
            ["highlight_code"] = "高亮代码",
            ["remove_blank_lines"] = "删除空行",
            ["set_indent"] = "调整缩进",
            ["move_blocks"] = "移动段落",
            ["merge_outlines"] = "合并文本框",
            ["insert_blocks"] = "插入段落",
            ["text_to_table"] = "转换为表格",
            ["get_pending_changes"] = "检查格式草稿",
            ["finish_edit"] = "写回并验证"
        };

        internal static string DisplayName(string name) => name != null && ToolNames.TryGetValue(name, out var text) ? text : "校验工具请求";

        /// <summary>
        /// 模型思考时常夹带的工具参数、返回字段里的英文词，换成窗口里的说法。
        /// 只收意思明确的；换不了的英文词留着，由 <see cref="LiveText.Gist"/> 把整句过滤掉。
        /// </summary>
        private static readonly Dictionary<string, string> ThoughtTerms = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["page_title"] = "页面标题", ["heading1"] = "一级标题", ["heading2"] = "二级标题", ["heading"] = "标题", ["headings"] = "标题",
            ["body"] = "正文", ["quote"] = "引用", ["preset"] = "预设样式", ["preset_id"] = "预设样式", ["style"] = "样式", ["styles"] = "样式",
            ["block"] = "段落", ["blocks"] = "段落", ["block_id"] = "段落", ["block_ids"] = "段落", ["paragraph"] = "段落", ["paragraphs"] = "段落",
            ["outline"] = "文本框", ["outlines"] = "文本框", ["container_id"] = "文本框",
            ["table"] = "表格", ["tables"] = "表格", ["image"] = "图片", ["images"] = "图片",
            ["draft"] = "草稿", ["draft_revision"] = "草稿版本", ["bullet"] = "项目符号", ["todo"] = "待办", ["tag"] = "标记", ["list"] = "列表",
            ["indent"] = "缩进", ["spacing"] = "间距", ["font"] = "字体", ["unhighlighted_code"] = "未高亮的代码", ["highlighted_code"] = "已有代码框"
        };

        /// <summary>段落、表格、图片和新插入段落的 ID（p3、t1、i2、n1），前后可能带着「第」「段」。</summary>
        private static readonly System.Text.RegularExpressions.Regex ThoughtId = new System.Text.RegularExpressions.Regex(
            @"(?:第\s*)?(?<![A-Za-z0-9_])(?<kind>[ptin])(?<from>\d+)(?:\s*(?:-|–|—|~|～|到|至)\s*\k<kind>?(?<to>\d+))?(?![A-Za-z0-9_])(?:\s*(?:段落|段|个表格|表格|张图片|图片))?");

        private static readonly System.Text.RegularExpressions.Regex ThoughtWord = new System.Text.RegularExpressions.Regex(
            @"(?<![A-Za-z0-9_])[A-Za-z][A-Za-z0-9_]*(?![A-Za-z0-9_])");

        /// <summary>
        /// 思考摘录里的工具名、预设名和段落 ID 换成中文（见 <see cref="LiveText.Gist"/>）：
        /// 「用 set_paragraph_style 把 p3 设为 heading2」→「用「设置段落样式」把第 3 段设为二级标题」。
        /// </summary>
        internal static string LocalizeThought(string text)
        {
            text = ThoughtId.Replace(text, m =>
            {
                var number = m.Groups["from"].Value + (m.Groups["to"].Success ? "–" + m.Groups["to"].Value : "");
                switch (m.Groups["kind"].Value)
                {
                    case "p": return $"第 {number} 段";
                    case "t": return $"第 {number} 个表格";
                    case "i": return $"第 {number} 张图片";
                    default: return $"新插入的第 {number} 段";
                }
            });
            return ThoughtWord.Replace(text, m =>
                ToolNames.TryGetValue(m.Value, out var tool) ? "「" + tool + "」"
                : ThoughtTerms.TryGetValue(m.Value, out var term) ? term
                : m.Value);
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
                case "set_list":
                    detail = JoinDetail(ListName(AiClient.Get(args, "list") as string), CountOf(args, "block_ids"));
                    break;
                case "set_tag":
                    detail = JoinDetail(TagName(AiClient.Get(args, "tag") as string, AiClient.Get(args, "completed") as bool?), CountOf(args, "block_ids"));
                    break;
                case "set_table_style":
                    detail = AiClient.Get(args, "table_ids") is IList tables ? $"{tables.Count} 个表格" : null;
                    break;
                case "read_image_text":
                    detail = AiClient.Get(args, "image_ids") is IList images ? $"{images.Count} 张" : null;
                    break;
                case "highlight_code":
                    detail = JoinDetail(LanguageName(AiClient.Get(outcome, "language") as string ?? AiClient.Get(args, "language") as string),
                        CountOf(args, "block_ids"));
                    break;
                case "remove_blank_lines":
                    detail = AiClient.Get(outcome, "removed") is IList removed ? $"{removed.Count} 行" : null;
                    break;
                case "set_indent":
                    var direction = AiClient.Get(args, "direction") as string;
                    detail = JoinDetail(direction == "in" ? "增加缩进" : direction == "out" ? "减少缩进" : null, CountOf(args, "block_ids"));
                    break;
                case "move_blocks":
                case "text_to_table":
                    detail = CountOf(args, "block_ids");
                    break;
                case "insert_blocks":
                    detail = AiClient.Get(args, "paragraphs") is IList inserted ? $"{inserted.Count} 段" : null;
                    break;
                case "merge_outlines":
                    detail = AiClient.Get(outcome, "moved") is IList merged ? $"{merged.Count} 段" : null;
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
        private static string ListName(string kind)
        {
            switch (kind)
            {
                case "bullet": return "项目符号";
                case "number": return "编号";
                case "none": return "取消列表";
                default: return null;
            }
        }
        private static string TagName(string kind, bool? completed)
        {
            switch (kind)
            {
                case "todo": return completed == true ? "待办已完成" : completed == false ? "待办未完成" : "待办";
                case "important": return "重要";
                case "question": return "问题";
                case "none": return "去掉标记";
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
            var items = Ordered();
            return new { snapshot_id = _snapshot.SnapshotId, page_title = _snapshot.Title, scope = _snapshot.SelectionOnly ? "selected_paragraphs" : "page",
                draft_revision = _snapshot.Revision, presets = ParagraphStyles.Ids, native_headings = _snapshot.Options.EnableNativeHeadings,
                paragraph_spacing = _snapshot.Options.EnableParagraphSpacing, code_highlight = _snapshot.Options.EnableCodeHighlight,
                languages = _snapshot.Options.EnableCodeHighlight ? Languages : null,
                list_edit = Has("set_list"), tag_edit = Has("set_tag"), table_style = Has("set_table_style"),
                table_shadings = Has("set_table_style") ? TableLook.Shadings : null,
                blank_lines = Has("remove_blank_lines"), indent = Has("set_indent"), move = Has("move_blocks"), insert = Has("insert_blocks"), text_table = Has("text_to_table"),
                total = items.Count, next_offset = offset + 100 < items.Count ? (int?)(offset + 100) : null,
                // 文本框按结构草稿列出，合并掉的不在其中；structure 表示能不能调整结构（移动、合并等）。
                outlines = _snapshot.Layout.Elements(OneNoteApi.One + "Outline").Select(o =>
                {
                    var id = (string)o.Attribute("objectID") ?? "";
                    var inside = items.Where(x => x.Node.Ancestors(OneNoteApi.One + "Outline").FirstOrDefault() == o).ToList();
                    var first = inside.Select(x => x.Block == null ? x.New.Text : x.Block.Editable || x.Block.CodeCandidate ? x.Block.CurrentText : null)
                        .FirstOrDefault(t => !string.IsNullOrWhiteSpace(t));
                    return new { container_id = id, structure = _snapshot.EditableOutlines.Contains(id), blocks = inside.Count,
                        summary = first?.Substring(0, Math.Min(40, first.Length)) };
                }).ToArray(),
                // 按结构草稿里的顺序和层级列出，已删除的空行不在其中，新插入的段落带 inserted。container_id 是现在所在的文本框。
                blocks = items.Skip(offset).Take(100).Select(x => x.Block == null
                    ? (object)new { id = x.Id, container_id = ContainerOf(x.Node), parent_id = ParentKey(x.Node), depth = x.Node.Ancestors(OneNoteApi.One + "OE").Count(),
                        inserted = true, summary = x.New.Text.Substring(0, Math.Min(80, x.New.Text.Length)) }
                    : new { id = x.Id, container_id = ContainerOf(x.Node), parent_id = ParentKey(x.Node), depth = x.Node.Ancestors(OneNoteApi.One + "OE").Count(),
                        table_id = x.Block.TableId, editable = x.Block.Editable, reason = x.Block.ProtectedReason,
                        list = AgentMarks.ListKind(x.Block.Draft), tags = AgentMarks.Describe(x.Block.Draft, _snapshot.DraftTags),
                        summary = x.Block.Editable || x.Block.CodeCandidate ? x.Block.CurrentText.Substring(0, Math.Min(80, x.Block.CurrentText.Length)) : null }).ToArray(),
                // 表格和图片数量有限，不分页。
                tables = _snapshot.Tables.Select(t => new { id = t.Id, container_id = ContainerOf(LayoutObject(t.ObjectId), t.ContainerId), rows = t.Rows, columns = t.Columns,
                    borders = t.Draft.Borders, header_row = t.Draft.HeaderRow, header_shading = t.Draft.ShadingName,
                    editable = t.Editable, reason = t.ProtectedReason, first_row = t.Summary }).ToArray(),
                images = _snapshot.Images.Select(i => new { id = i.Id, container_id = ContainerOf(LayoutObject(i.ObjectId), i.ContainerId), chars = i.Text.Length }).ToArray() };
        }
        /// <summary>结构草稿里对象现在所在的顶层容器：文本框的 ID，标题为 Title。不在文本框、标题里时用 fallback（快照时的位置）。</summary>
        private string ContainerOf(XElement node, string fallback = "")
        {
            var container = node?.Ancestors().FirstOrDefault(e => e.Parent == _snapshot.Layout);
            return container == null ? fallback : (string)container.Attribute("objectID") ?? container.Name.LocalName;
        }
        private XElement LayoutObject(string objectId) => objectId == null ? null : _snapshot.Layout.Descendants().FirstOrDefault(e => (string)e.Attribute("objectID") == objectId);
        /// <summary>结构草稿里按页面顺序排的段落：快照段落和新插入的段落，已删除的不在其中。</summary>
        private List<(string Id, AgentBlock Block, AgentInserted New, XElement Node)> Ordered()
        {
            var blocks = _snapshot.Blocks.ToDictionary(b => b.Id);
            var inserted = _snapshot.Inserted.ToDictionary(i => i.Id);
            var result = new List<(string, AgentBlock, AgentInserted, XElement)>();
            foreach (var oe in _snapshot.Layout.Descendants(OneNoteApi.One + "OE"))
            {
                var key = AgentLayout.KeyOf(oe);
                if (key == null) continue;
                blocks.TryGetValue(key, out var block);
                inserted.TryGetValue(key, out var added);
                if (block != null || added != null) result.Add((key, block, added, oe));
            }
            return result;
        }
        /// <summary>上级段落的短 ID；上级不是快照里的段落（比如表格外层）时为 null。</summary>
        private static string ParentKey(XElement oe) => AgentLayout.KeyOf(oe.Ancestors(OneNoteApi.One + "OE").FirstOrDefault());

        private List<AgentBlock> Targets(IDictionary<string, object> args)
        {
            var ids = ((IList)args["block_ids"]).Cast<string>().ToList();
            if (ids.Distinct().Count() != ids.Count) throw new AiException("目标段落重复。");
            return ids.Select(id => Block(id, false)).ToList();
        }
        private AgentBlock Block(string id, bool writing)
        {
            var b = _snapshot.Blocks.FirstOrDefault(x => x.Id == id);
            if (writing && b?.Conversion != null) throw new AiException("段落已排入代码框或表格转换，不能再修改。");
            if (writing && b != null && b.CodeCandidate) throw new AiException("代码段落不能设置样式或修改文字；用 highlight_code 转换为代码框。");
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
            return new { snapshot_id = _snapshot.SnapshotId, blocks = blocks.Select(b => new { id = b.Id, kind = b.CodeCandidate ? "unhighlighted_code" : "text", text = b.CurrentText,
                depth = AgentLayout.Find(page, b.Id).Ancestors(OneNoteApi.One + "OE").Count(), container_id = ContainerOf(AgentLayout.Find(_snapshot.Layout, b.Id), b.ContainerId),
                parent_id = ParentKey(AgentLayout.Find(page, b.Id)),
                table_id = b.TableId, style = Css.Effective(AgentLayout.Find(page, b.Id), page),
                list = AgentMarks.ListKind(b.Draft), tags = AgentMarks.Describe(b.Draft, _snapshot.DraftTags),
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

        /// <summary>列表和标记是段落前的符号，页面标题上不能加。其余要求和格式工具一样：完整读取过、不受保护、不是代码。</summary>
        private AgentBlock MarkTarget(AgentBlock b)
        {
            Block(b.Id, true);
            if (AgentCommitter.Find(_snapshot.Page, b.ObjectId).Parent?.Name == OneNoteApi.One + "Title") throw new AiException("页面标题不能设置列表或标记。");
            return b;
        }

        private object SetList(IDictionary<string, object> args)
        {
            var kind = (string)args["list"];
            var drafts = new Dictionary<AgentBlock, XElement>();
            foreach (var b in Targets(args).Select(MarkTarget))
            {
                var draft = new XElement(b.Draft);
                AgentMarks.SetList(draft, kind);
                drafts.Add(b, draft);
            }
            return Publish(drafts);
        }

        private object SetTag(IDictionary<string, object> args)
        {
            var kind = (string)args["tag"];
            var completed = args.TryGetValue("completed", out var value) ? (bool?)(bool)value : null;
            if (completed != null && kind != "todo") throw new AiException("completed 只能用于待办（todo）标记。");
            var definitions = new XElement(_snapshot.DraftTags);
            var drafts = new Dictionary<AgentBlock, XElement>();
            foreach (var b in Targets(args).Select(MarkTarget))
            {
                var draft = new XElement(b.Draft);
                AgentMarks.SetTag(draft, kind, completed, definitions);
                drafts.Add(b, draft);
            }
            var result = Publish(drafts);
            _snapshot.DraftTags.ReplaceNodes(definitions.Elements().Select(e => new XElement(e)));
            return result;
        }

        private object TableStyle(IDictionary<string, object> args)
        {
            var ids = ((IList)args["table_ids"]).Cast<string>().ToList();
            if (ids.Distinct().Count() != ids.Count) throw new AiException("目标表格重复。");
            var tables = ids.Select(id => _snapshot.Tables.FirstOrDefault(t => t.Id == id) ?? throw new AiException("目标表格不存在。")).ToList();
            if (tables.Any(t => !t.Editable)) throw new AiException("目标表格受到保护，不能设置样式。");
            var style = (IDictionary<string, object>)args["style"];
            // 全部算好再发布，失败时这个工具没有副作用。
            var looks = tables.Select(t => (Table: t, Look: t.Draft.With(style))).ToList();
            var changed = looks.Where(l => !l.Look.SameAs(l.Table.Draft)).ToList();
            foreach (var l in changed) l.Table.Draft = l.Look;
            if (changed.Count > 0) _snapshot.Revision++;
            return new { ok = true, draft_revision = _snapshot.Revision, changed = changed.Select(l => l.Table.Id).ToArray(),
                noop = looks.Except(changed).Select(l => l.Table.Id).ToArray() };
        }

        private object ImageText(IDictionary<string, object> args)
        {
            var ids = ((IList)args["image_ids"]).Cast<string>().ToList();
            if (ids.Distinct().Count() != ids.Count) throw new AiException("目标图片重复。");
            var images = ids.Select(id => _snapshot.Images.FirstOrDefault(i => i.Id == id) ?? throw new AiException("图片不存在或没有识别出的文字。")).ToList();
            return new { snapshot_id = _snapshot.SnapshotId, images = images.Select(i => new { id = i.Id, text = i.Text, truncated = i.Truncated }).ToArray() };
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
            var selection = AgentCode.Select(_snapshot.Layout, blocks.Select(b => b.ObjectId).ToList());
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

        /// <summary>把用制表符、竖线或空格分隔的连续段落排入草稿，提交时换成表格；和代码框一样整段替换，撤销时换回原段落。</summary>
        private object TextToTable(IDictionary<string, object> args)
        {
            var blocks = Ids(args, "block_ids").Select(id => _snapshot.Blocks.FirstOrDefault(b => b.Id == id) ?? throw new AiException("目标不存在。")).ToList();
            foreach (var b in blocks)
            {
                if (b.Conversion != null) throw new AiException("部分段落已排入代码框或表格转换。");
                if (b.CodeCandidate) throw new AiException("代码段落请用 highlight_code 转换为代码框。");
                if (!(b.Editable || b.ProtectedReason == "empty") || !_snapshot.EditableOutlines.Contains(ContainerOf(AgentLayout.Find(_snapshot.Layout, b.Id))))
                    throw new AiException("目标受到保护或已删除，不能转换为表格。");
                if (b.ProtectedReason != "empty" && !b.Read) throw new AiException("请先完整读取目标段落。");
            }
            // 先全部校验、生成表格，再发布草稿；失败时这个工具没有副作用。
            var selection = AgentCode.Select(_snapshot.Layout, blocks.Select(b => b.ObjectId).ToList(), "表格");
            if (selection.Block.Name == OneNoteApi.One + "Cell") throw new AiException("表格单元格里的段落不能再转换为表格。");
            var shading = args.TryGetValue("header_shading", out var shade) && (string)shade != "none" ? TableLook.Shade((string)shade) : null;
            var names = args.TryGetValue("header", out var header) ? ((IList)header).Cast<string>().ToList() : null;
            var table = AgentTextTable.Build(selection.Paragraphs, (string)args["delimiter"],
                !args.TryGetValue("header_row", out var headerRow) || (bool)headerRow, !args.TryGetValue("borders", out var borders) || (bool)borders, shading, names, out var padded);
            var ordered = selection.Paragraphs.Select(oe => blocks.First(b => b.Id == AgentLayout.KeyOf(oe))).ToList();
            var conversion = new AgentCodeConversion { Blocks = ordered, TextTable = true, Code = selection.Code, Table = table };
            // 转换后这些段落就不在了，之前给它们排的格式和文字修正作废。
            var discarded = ordered.Where(b => b.Changed).Select(b => b.Id).ToArray();
            foreach (var b in ordered) { b.Draft = new XElement(b.Original); b.TextFixes.Clear(); b.Conversion = conversion; }
            _snapshot.CodeConversions.Add(conversion);
            _snapshot.Revision++;
            return new { ok = true, draft_revision = _snapshot.Revision, changed = ordered.Select(b => b.Id).ToArray(), rows = table.Elements(OneNoteApi.One + "Row").Count(),
                columns = table.Element(OneNoteApi.One + "Columns").Elements().Count(), padded_rows = padded, discarded_format = discarded };
        }

        private object RemoveBlankLines(IDictionary<string, object> args)
        {
            var blanks = new HashSet<string>(_snapshot.Blocks.Where(b => b.ProtectedReason == "empty" && b.Conversion == null && _snapshot.EditableOutlines.Contains(b.ContainerId)).Select(b => b.Id));
            var candidate = new XElement(_snapshot.Layout);
            var outlines = candidate.Descendants(OneNoteApi.One + "OE").Where(e => blanks.Contains(AgentLayout.KeyOf(e) ?? "")).ToDictionary(e => e, OutlineOf);
            var removed = AgentLayout.RemoveBlankLines(candidate, e => blanks.Contains(AgentLayout.KeyOf(e) ?? ""), (string)args["mode"] == "all");
            PublishLayout(candidate, "removed", removed.Select(e => (outlines[e], (string)null, AgentLayout.KeyOf(e))).ToList());
            return new { ok = true, draft_revision = _snapshot.Revision, removed = removed.Select(AgentLayout.KeyOf).ToArray() };
        }

        private object SetIndent(IDictionary<string, object> args)
        {
            var ids = Ids(args, "block_ids");
            var candidate = new XElement(_snapshot.Layout);
            var nodes = ids.Select(id => Movable(candidate, id)).ToList();
            var listed = new HashSet<XElement>(nodes);
            // 上级段落也在列表里时，下级段落跟着它走，不再单独调整。
            var tops = nodes.Where(n => !n.Ancestors(OneNoteApi.One + "OE").Any(listed.Contains)).InDocumentOrder().ToList();
            var changes = tops.Select(n => (OutlineOf(n), (string)null, AgentLayout.KeyOf(n))).ToList();
            foreach (var n in tops) AgentLayout.Indent(n, (string)args["direction"] == "in", AgentLayout.KeyOf(n));
            var changed = PublishLayout(candidate, "indented", changes);
            return new { ok = true, draft_revision = _snapshot.Revision, changed = changed ? ids.ToArray() : new string[0], noop = changed ? new string[0] : ids.ToArray() };
        }

        private object MoveBlocks(IDictionary<string, object> args)
        {
            var ids = Ids(args, "block_ids");
            var targetId = (string)args["target_id"];
            if (ids.Contains(targetId)) throw new AiException("目标段落不能是要移动的段落。");
            var candidate = new XElement(_snapshot.Layout);
            var nodes = ids.Select(id => Movable(candidate, id)).ToList();
            var target = Node(candidate, targetId, out _);
            var listed = new HashSet<XElement>(nodes);
            if (nodes.Any(n => n.Ancestors(OneNoteApi.One + "OE").Any(listed.Contains))) throw new AiException("下级段落会随上级段落一起移动，不要同时列出。");
            if (target.Ancestors(OneNoteApi.One + "OE").Any(listed.Contains)) throw new AiException("目标段落不能在要移动的段落下面。");
            // 可以移到另一个文本框；表格单元格里的段落只能在同一个单元格里移动。
            var block = PageEditor.TextBlockOf(target);
            if (nodes.Any(n => PageEditor.TextBlockOf(n) != block && (block.Name != OneNoteApi.One + "Outline" || PageEditor.TextBlockOf(n).Name != OneNoteApi.One + "Outline")))
                throw new AiException("表格单元格里的段落只能在同一个单元格里移动，也不能从外面移进单元格。");
            var ordered = nodes.InDocumentOrder().ToList();
            var into = OutlineOf(target);
            var changes = ordered.Select(n => (into, OutlineOf(n) == into ? null : OutlineOf(n), AgentLayout.KeyOf(n))).ToList();
            var sources = changes.Select(c => c.Item2).Where(from => from != null).Distinct().ToList();
            AgentLayout.Move(ordered, target, (string)args["position"] == "before");
            if (sources.Any(from => !candidate.Elements(OneNoteApi.One + "Outline").First(o => (string)o.Attribute("objectID") == from).Descendants(OneNoteApi.One + "OE").Any()))
                throw new AiException("这样会把文本框移空；要把整个文本框并入另一个文本框，请用 merge_outlines。");
            var changed = PublishLayout(candidate, "moved", changes);
            return new { ok = true, draft_revision = _snapshot.Revision, changed = changed ? ids.ToArray() : new string[0], noop = changed ? new string[0] : ids.ToArray() };
        }

        /// <summary>把整个文本框并到另一个文本框里目标段落的前后，提交时删掉空了的源文本框。</summary>
        private object MergeOutlines(IDictionary<string, object> args)
        {
            var sourceId = (string)args["source_id"];
            var candidate = new XElement(_snapshot.Layout);
            var source = candidate.Elements(OneNoteApi.One + "Outline").FirstOrDefault(o => (string)o.Attribute("objectID") == sourceId);
            if (source == null) throw new AiException("源文本框不存在或已经合并。");
            if (!_snapshot.EditableOutlines.Contains(sourceId)) throw new AiException("源文本框受到保护，不能调整结构。");
            var target = Node(candidate, (string)args["target_id"], out _);
            if (target.Ancestors(OneNoteApi.One + "Outline").First() == source) throw new AiException("目标段落不能在源文本框里。");
            if (PageEditor.TextBlockOf(target).Name != OneNoteApi.One + "Outline") throw new AiException("目标段落不能在表格单元格里。");
            // 只处理选中范围时，没选中的段落不在快照里，不能跟着搬走。
            if (source.Descendants(OneNoteApi.One + "OE").Any(e => e.Elements(OneNoteApi.One + "T").Any() && AgentLayout.KeyOf(e) == null))
                throw new AiException("源文本框里有没选中的段落，不能整体合并。");
            var moved = source.Descendants(OneNoteApi.One + "OE").Select(AgentLayout.KeyOf).Where(k => k != null).ToArray();
            var into = OutlineOf(target);
            AgentLayout.Merge(source, target, (string)args["position"] == "before");
            PublishLayout(candidate, new[] { new AgentLayoutChange { Kind = "merged", OutlineId = into, From = sourceId, Ids = moved } });
            return new { ok = true, draft_revision = _snapshot.Revision, merged = sourceId, into, moved };
        }

        private static readonly char[] Breaks = { '\n', '\r', '\t' };

        private object InsertBlocks(IDictionary<string, object> args)
        {
            var candidate = new XElement(_snapshot.Layout);
            var target = Node(candidate, (string)args["target_id"], out _);
            var items = ((IList)args["paragraphs"]).Cast<IDictionary<string, object>>().ToList();
            var texts = items.Select(i => ((string)i["text"]).Trim()).ToList();
            if (texts.Any(t => t.Length == 0)) throw new AiException("插入的段落不能是空白。");
            if (texts.Any(t => t.IndexOfAny(Breaks) >= 0)) throw new AiException("插入的文字不能含换行或制表符，每段单独写一项。");
            if (_snapshot.Inserted.Count + items.Count > MaxInserted || _snapshot.Inserted.Sum(i => i.Text.Length) + texts.Sum(t => t.Length) > MaxInsertedChars)
                throw new AiException($"每个任务最多插入 {MaxInserted} 段、{MaxInsertedChars} 字。");
            var styles = new XElement(_snapshot.DraftStyles);
            var outline = OutlineOf(target);
            var created = items.Select((item, i) => AgentLayout.NewParagraph("n" + (_snapshot.Inserted.Count + i + 1), texts[i], (string)item["preset_id"],
                item.TryGetValue("list", out var list) ? (string)list : null, _snapshot.Options, styles)).ToList();
            if ((string)args["position"] == "before") target.AddBeforeSelf(created); else target.AddAfterSelf(created);
            PublishLayout(candidate, "inserted", created.Select(e => (outline, (string)null, AgentLayout.KeyOf(e))).ToList());
            _snapshot.DraftStyles.ReplaceNodes(styles.Elements().Select(e => new XElement(e)));
            _snapshot.Inserted.AddRange(created.Select((e, i) => new AgentInserted { Id = AgentLayout.KeyOf(e), OutlineId = outline, Text = texts[i] }));
            return new { ok = true, draft_revision = _snapshot.Revision, inserted = created.Select(AgentLayout.KeyOf).ToArray() };
        }

        private static List<string> Ids(IDictionary<string, object> args, string key)
        {
            var ids = ((IList)args[key]).Cast<string>().ToList();
            if (ids.Distinct().Count() != ids.Count) throw new AiException("目标段落重复。");
            return ids;
        }

        private static string OutlineOf(XElement oe) => (string)oe.Ancestors(OneNoteApi.One + "Outline").First().Attribute("objectID");

        /// <summary>结构工具的对象：快照段落或新插入的段落，返回结构草稿 layout 里的 OE。只能在可调整结构的文本框里。</summary>
        private XElement Node(XElement layout, string id, out AgentBlock block)
        {
            block = _snapshot.Blocks.FirstOrDefault(b => b.Id == id);
            var node = AgentLayout.Find(layout, id);
            if (node == null || (block == null && !_snapshot.Inserted.Any(i => i.Id == id))) throw new AiException($"段落 {id} 不存在或已删除。");
            var container = node.Ancestors().FirstOrDefault(e => e.Parent == layout);
            if (container == null || container.Name != OneNoteApi.One + "Outline" || !_snapshot.EditableOutlines.Contains((string)container.Attribute("objectID") ?? ""))
                throw new AiException($"段落 {id} 在页面标题或受保护的文本框里，不能调整结构。");
            if (block?.Conversion != null) throw new AiException($"段落 {id} 已排入代码框或表格转换，不能调整结构。");
            if (block?.ProtectedReason == "highlighted_code") throw new AiException($"段落 {id} 在代码框里，不能调整结构。");
            return node;
        }

        /// <summary>要移动、调整缩进的段落：读过的文字段落、待转换的代码、空行或新插入的段落。</summary>
        private XElement Movable(XElement layout, string id)
        {
            var node = Node(layout, id, out var block);
            if (block == null || block.ProtectedReason == "empty") return node;
            if (!(block.Editable || block.CodeCandidate)) throw new AiException($"段落 {id} 受到保护，不能移动或调整缩进。");
            if (!block.Read) throw new AiException("请先完整读取目标段落。");
            return node;
        }

        /// <summary>结构工具共用的发布：通过不变量检查后换上新的结构草稿并记下改动。结构没有变化时不动修订号，返回 false。</summary>
        private bool PublishLayout(XElement candidate, string kind, IList<(string Outline, string From, string Id)> changes) =>
            PublishLayout(candidate, changes.GroupBy(c => (c.Outline, c.From)).Select(g =>
                new AgentLayoutChange { Kind = kind, OutlineId = g.Key.Outline, From = g.Key.From, Ids = g.Select(c => c.Id).ToArray() }));
        private bool PublishLayout(XElement candidate, IEnumerable<AgentLayoutChange> changes)
        {
            if (XNode.DeepEquals(candidate, _snapshot.Layout)) return false;
            _snapshot.CheckLayout(candidate);
            _snapshot.Layout = candidate;
            _snapshot.LayoutChanges.AddRange(changes);
            _snapshot.Revision++;
            return true;
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
                    AgentMarks.DraftKey(pair.Key.Draft) == AgentMarks.DraftKey(pair.Value) &&
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
            code_blocks = _snapshot.CodeConversions.Where(c => !c.TextTable).Select(c => new { block_ids = c.Blocks.Select(b => b.Id).ToArray(), language = c.LanguageId }).ToArray(),
            text_tables = _snapshot.CodeConversions.Where(c => c.TextTable).Select(c => new { block_ids = c.Blocks.Select(b => b.Id).ToArray() }).ToArray(),
            layout = new { removed = LayoutIds("removed"), moved = LayoutIds("moved"), indented = LayoutIds("indented"),
                inserted = _snapshot.Inserted.Select(i => new { id = i.Id, text = i.Text }).ToArray(),
                merged = _snapshot.LayoutChanges.Where(c => c.Kind == "merged").Select(c => new { from = c.From, into = c.OutlineId }).ToArray() },
            unconverted_code = _snapshot.Blocks.Where(b => b.CodeCandidate && b.Conversion == null).Select(b => b.Id).ToArray(),
            tables_changed = _snapshot.Tables.Where(t => t.Changed).Select(t => t.Id).ToArray(),
            protected_count = _snapshot.Blocks.Count(b => !b.Editable && b.Conversion == null) };
        private string[] LayoutIds(string kind) => _snapshot.LayoutChanges.Where(c => c.Kind == kind).SelectMany(c => c.Ids).Distinct().ToArray();

        private object Finish(IDictionary<string, object> args)
        {
            if (Convert.ToInt32(args["draft_revision"]) != _snapshot.Revision) throw new AiException("草稿修订号过期，请先检查待提交修改。");
            _snapshot.Frozen = true;
            Report = _committer.Commit(_snapshot, _cancellation);
            return Report.ToToolResult();
        }
    }
}
