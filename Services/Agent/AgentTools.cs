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
        /// <summary>字符串可以为空，如 text_to_table rows 里的空单元格。</summary>
        internal bool AllowEmpty;
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
            if (Type == "string") { value["minLength"] = AllowEmpty ? 0 : 1; value["maxLength"] = MaxLength; }
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
                    if (!(value is string s) || (s.Length == 0 && !AllowEmpty) || s.Length > MaxLength || (Enum != null && !Enum.Contains(s))) throw new AiException(path + " 的字符串无效。");
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
        /// <summary>执行开始时的围栏上下文；不随删围栏、改字或移动段落重新分类。</summary>
        private readonly Dictionary<string, List<AgentMarkdown.LineRole>> _markdownRoles = new Dictionary<string, List<AgentMarkdown.LineRole>>();
        /// <summary>规范化代码框间隔前的结构草稿，以及当时已有的插入段落数和结构改动数；之后没有间隔改动时为 null。</summary>
        private (XElement Layout, int Inserted, int Changes)? _spacingBase;
        /// <summary>会改变结构或转换范围的工具：执行前先撤回间隔调整，免得旧的删补留在已不相邻的段落上。</summary>
        private static readonly HashSet<string> SpacingSensitive = new HashSet<string>
            { "remove_blank_lines", "set_indent", "move_blocks", "merge_outlines", "insert_blocks", "strip_markdown", "highlight_code", "text_to_table" };
        internal AgentReport Report { get; private set; }
        // 只由运行器在最后一轮开启；模型的工具参数不能绕过完整读取检查。
        internal bool AllowIncompleteFinish { get; set; }
        internal const int ReadBatchSize = 100;
        internal int UnreadCount => UnreadBlockIds().Length;
        /// <summary>本次注册了这个工具；系统提示词按实际提供的工具追加说明。</summary>
        internal bool Has(string name) => _tools.ContainsKey(name);
        internal object[] Definitions => _tools.Select(Definition).ToArray();
        /// <summary>只取指定的工具定义，比如最后一轮只给 finish_edit。</summary>
        internal object[] DefinitionsOf(params string[] names) => _tools.Where(t => names.Contains(t.Key)).Select(Definition).ToArray();
        private static object Definition(KeyValuePair<string, Tool> t) =>
            new { type = "function", function = new { name = t.Key, description = t.Value.Description, parameters = t.Value.Schema.Json() } };

        internal AgentTools(AgentPageSnapshot snapshot, AgentCommitter committer, CancellationToken cancellation, AddInSettings codeSettings = null)
        {
            _snapshot = snapshot; _committer = committer; _cancellation = cancellation; _code = codeSettings ?? new AddInSettings();
            if (snapshot.Options.EnableMarkdownCleanup) CacheFenceRoles();
            Register("get_page_overview", "获取当前固定页面的段落摘要、保护范围及样式。每页 100 项；通过 offset 翻页。", AgentSchema.Obj(new Dictionary<string, AgentSchema>
            { ["offset"] = AgentSchema.Num(0, 1000, true) }), Overview);
            Register("read_blocks", "完整读取段落正文及样式。修改前必须调用，不能修改受保护段落；传入的受保护段落会跳过，列在 skipped 里并附原因。", WithIds(), Read);
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
            Register("set_paragraph_style", "为完整读取的段落设置标题、正文或引用样式，可受限覆盖；只修改草稿，不改文字。原生标题由能力开关决定。" +
                "若原生样式会影响未指定的下级段落，则只设置本段外观并保留原有标题层级，结果 appearance_only 列出这些段落。", paragraph, Paragraph);
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
            if (snapshot.Options.EnableMarkdownCleanup)
            {
                var markdown = WithIds();
                markdown.Properties["block_ids"].MaxItems = 1000;
                markdown.Properties["kinds"] = AgentSchema.Array(AgentSchema.Str(AgentMarkdown.Kinds));
                markdown.Properties["emphasis"] = AgentSchema.Str("remove", "format");
                Register("strip_markdown", "去掉完整读取的段落里的 Markdown 标记，只删标记字符，其余文字、格式和链接不变：" +
                    "行首的 # 标题、> 引用、- * + 和 1. 列表符号（含 [ ] 待办框），行内 `代码`、**粗体**、*斜体*、~~删除线~~ 的符号，" +
                    "以及整行的 ``` 围栏和 --- 分隔线（能删整段时删掉这一段，否则清空文字）。kinds 只处理其中几类，默认全部。" +
                    "emphasis 默认 remove 只去掉符号；format 同时把内容设为粗体、斜体或删除线。" +
                    "markdown/md 围栏里照常处理；其他围栏里的代码不动，列在 code_lines 里。链接、图片和表格不处理。" +
                    "结果 changed 里的 heading（级别）、list、todo、quote、indent 是原来的标记，可以据此设置标题、列表、待办和引用。", markdown, StripMarkdown);
            }
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
            if (structural && (snapshot.Options.EnableBlankLineRemoval || snapshot.Options.EnableInsert))
                Register("normalize_code_spacing", "只规范化同一文本框或单元格内代码框与文字的交界，空段落和文字首尾的 Shift+Enter 空行合计保留一行，多删少补。" +
                    "先转换代码并完成其他结构调整，最后调用此工具；之后再调整结构、转换代码或清理 Markdown 会撤回已做的间隔调整。代码内部、文字之间及文本框首尾不处理。" +
                    "删除和补入分别受删空行、插入段落开关控制，关闭时对应边界返回 skipped。",
                    SnapshotOnly(), NormalizeCodeSpacing);
            if (structural && snapshot.Options.EnableBlankLineRemoval && (snapshot.Options.EnableInsert ||
                snapshot.Blocks.Any(b => b.ProtectedReason == "empty" && snapshot.EditableOutlines.Contains(b.ContainerId))))
            {
                var blank = SnapshotOnly();
                blank.Properties["mode"] = AgentSchema.Str("collapse", "all");
                blank.Required = new[] { "snapshot_id", "mode" };
                Register("remove_blank_lines", "删除多余的空行（只有空白的段落）：collapse 把连续空行合并为一行，并删掉文本框、单元格首尾的空行；all 删掉全部空行。" +
                    "也处理本次用 insert_blocks 插入的空行；带列表、标记或下级段落的空段落和代码里的空行不删。", blank, RemoveBlankLines);
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
                        "再删掉空了的源文本框。source_id 是源文本框的 container_id（见 get_page_overview 的 outlines）。" +
                        "结果里的 mergeable_left 是能在目标框内至少一个合法位置继续合并的文本框数（mergeable_outline_ids 列出它们），已核验选区、格式和转换限制。" +
                        "用户要求合并全部文本框时继续处理这些框；为 0 表示当前目标框已无合法合并来源，受保护或不满足限制的框保留并说明原因。" +
                        "某个位置失败时按错误选择其他合法位置，不要重复相同的失败调用。",
                        AgentSchema.Obj(new Dictionary<string, AgentSchema>
                        {
                            ["snapshot_id"] = AgentSchema.Str(), ["source_id"] = AgentSchema.Str(), ["target_id"] = AgentSchema.Str(), ["position"] = AgentSchema.Str("before", "after")
                        }, "snapshot_id", "source_id", "target_id", "position"), MergeOutlines);
            }
            if (structural && snapshot.Options.EnableInsert)
            {
                var item = new Dictionary<string, AgentSchema> { ["text"] = AgentSchema.Short(MaxInsertChars), ["preset_id"] = AgentSchema.Str("heading1", "heading2", "body", "quote"),
                    ["blank"] = new AgentSchema { Type = "boolean" } };
                if (snapshot.Options.EnableLists) item["list"] = AgentSchema.Str("bullet", "number");
                // 文字项要 text 和 preset_id，空行项只写 blank，由 InsertBlocks 逐项校验。
                var entry = AgentSchema.Obj(item);
                entry.NonEmpty = true;
                var paragraphs = AgentSchema.Array(entry);
                paragraphs.MaxItems = MaxInsertPerCall;
                Register("insert_blocks", $"在目标段落前面或后面插入同级的新段落：只写纯文字、不含换行，使用预设样式。每次最多 {MaxInsertPerCall} 段，每段最多 {MaxInsertChars} 字，" +
                    $"每个任务最多 {MaxInserted} 段、{MaxInsertedChars} 字。新段落的 ID 为 n1、n2…，可以作为之后插入、移动的目标。" +
                    "需要空行时这一项只写 {\"blank\": true}（不带 text、preset_id、list），插入一行正文外观的空段落，比如合并文本框后隔开交界处；空行也算段数，已经有空行的地方不要再加。",
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
                table.Properties["lines_per_row"] = AgentSchema.Num(1, AgentTextTable.MaxColumns, true);
                table.Properties["header_row"] = new AgentSchema { Type = "boolean" };
                table.Properties["borders"] = new AgentSchema { Type = "boolean" };
                table.Properties["header_shading"] = AgentSchema.Str(TableLook.Shadings.Concat(new[] { "none" }).ToArray());
                var header = AgentSchema.Array(AgentSchema.Short(AgentTextTable.MaxHeaderChars));
                header.MaxItems = AgentTextTable.MaxColumns;
                table.Properties["header"] = header;
                var cells = AgentSchema.Array(new AgentSchema { Type = "string", MaxLength = AgentTextTable.MaxCellChars, AllowEmpty = true });
                cells.MaxItems = AgentTextTable.MaxColumns;
                var rows = AgentSchema.Array(cells);
                rows.MaxItems = AgentTextTable.MaxRows;
                table.Properties["rows"] = rows;
                table.Required = new[] { "snapshot_id", "block_ids" };
                Register("text_to_table", "把同一文本框里连续的段落转成表格，单元格保留原有文字格式和链接；先完整读取有文字的段落。delimiter 和 rows 二选一。" +
                    "用 delimiter 按规则拆分：tab 为制表符，pipe 为竖线，space 按连续空白拆分；每行文字一行（段内 Shift+Enter 换行的也各成一行），" +
                    "中间的空行和 Markdown 分隔行去掉。一条记录分成几行（如名称一行、地址一行，中间可以有空行）时，lines_per_row 设为每条记录的行数（空行不算），" +
                    "每几行合成表格的一行，各行不再拆分时 delimiter 用 none；记录里除最后一行外，行尾的冒号去掉。" +
                    "分隔不统一、键值对、有缺项等 delimiter 处理不了的文本用 rows 逐行给出单元格：每格逐字复制原文、按原文顺序排列，不改写、不合并、不重复，" +
                    "缺项用空字符串，单元格不能含换行。单元格之间的空白、制表符、| : ， ; 、 = 等分隔符可以省略；「IP：10.0.0.1」里的标签只有等于该列 header 时才能省略，" +
                    "其余文字都要放进单元格，工具会逐字核验。" +
                    "各行列数不一致时按最多的列数建表，缺的单元格留空，结果里的 padded_rows 是补了空单元格的行数。" +
                    "标题等不属于表格的段落不要放进 block_ids。header_row、borders 默认 true，第一行是数据不是列名时 header_row 设为 false；header_shading 是首行底色。" +
                    "只有用户要求加表头或要把标签挪成列名时才用 header 在首行前新增一行列名，个数应等于列数，少了补空，多了按 header 加列。" +
                    $"最多 {AgentTextTable.MaxRows} 行、{AgentTextTable.MaxColumns} 列。", table, TextToTable);
            }
            Register("get_pending_changes", "检查草稿修订号、改动和尚未读取的段落；next_read_block_ids 是下一批可完整读取的段落（最多 100 段）。", SnapshotOnly(), Pending);
            var finish = SnapshotOnly();
            finish.Properties["draft_revision"] = AgentSchema.Num(0, 10000, true);
            finish.Required = new[] { "snapshot_id", "draft_revision" };
            Register("finish_edit", "独立调用此工具提交草稿并验证结果；必须传入最新修订号，并先读完范围内所有可读取的段落。未读完时返回下一批 ID，继续 read_blocks 后重试；提交成功后不能继续编辑。", finish, Finish);
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
            ["strip_markdown"] = "去除 Markdown 符号",
            ["set_list"] = "设置列表",
            ["set_tag"] = "设置标记",
            ["set_table_style"] = "设置表格样式",
            ["read_image_text"] = "读取图片文字",
            ["highlight_code"] = "高亮代码",
            ["remove_blank_lines"] = "删除空行",
            ["normalize_code_spacing"] = "规范化代码框间隔",
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
                    detail = AiClient.Get(outcome, "blocks") is IList readBlocks
                        ? JoinDetail($"{readBlocks.Count} 段", AiClient.Get(outcome, "skipped") is IList skippedBlocks ? $"跳过受保护 {skippedBlocks.Count} 段" : null)
                        : CountOf(args, "block_ids");
                    break;
                case "set_paragraph_style":
                    detail = JoinDetail(PresetName(AiClient.Get(args, "preset_id") as string), CountOf(args, "block_ids"));
                    if (AiClient.Get(outcome, "appearance_only") is IList appearance && appearance.Count > 0)
                        detail = JoinDetail(detail, $"{appearance.Count} 段仅设置外观");
                    break;
                case "set_text_style":
                    detail = AiClient.Get(args, "targets") is IList targets ? $"{targets.Count} 处" : null;
                    break;
                case "fix_text":
                    detail = AiClient.Get(args, "fixes") is IList fixes ? $"{fixes.Count} 处" : null;
                    break;
                case "strip_markdown":
                    detail = AiClient.Get(outcome, "changed") is IList stripped ? $"{stripped.Count} 段" : CountOf(args, "block_ids");
                    if (AiClient.Get(outcome, "removed_lines") is IList separators && separators.Count > 0)
                        detail = JoinDetail(detail, $"删除围栏和分隔线 {separators.Count} 行");
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
                case "normalize_code_spacing":
                    detail = $"删除 {AiClient.Get(outcome, "removed_paragraphs") ?? 0} 个空段落、{AiClient.Get(outcome, "removed_soft_lines") ?? 0} 个段内空行，补入 {AiClient.Get(outcome, "inserted_paragraphs") ?? 0} 行";
                    if (AiClient.Get(outcome, "skipped") is IList spacingSkipped && spacingSkipped.Count > 0)
                        detail += $"，跳过 {spacingSkipped.Count} 处：" + (AiClient.Get(outcome, "skip_message") as string);
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
                    if (AiClient.Get(args, "paragraphs") is IList inserted)
                    {
                        var blankCount = inserted.Cast<object>().Count(p => Equals(AiClient.Get(p, "blank"), true));
                        detail = JoinDetail(inserted.Count > blankCount ? $"{inserted.Count - blankCount} 段" : null, blankCount > 0 ? $"空行 {blankCount} 行" : null);
                    }
                    break;
                case "merge_outlines":
                    detail = AiClient.Get(outcome, "moved") is IList merged ? $"{merged.Count} 段" : null;
                    if (AiClient.Get(outcome, "mergeable_left") is int left && left > 0) detail = JoinDetail(detail, $"还剩 {left} 个");
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
            if (_spacingBase == null || !SpacingSensitive.Contains(call.Name)) return tool.Execute(args);
            // 先撤回间隔调整，在不含删补的草稿上执行；提交前复核会要求重新规范化。工具失败或没有改动时原样恢复。
            var saved = (Layout: _snapshot.Layout, Inserted: _snapshot.Inserted.ToList(), Changes: _snapshot.LayoutChanges.ToList(), Base: _spacingBase, Revision: _snapshot.Revision);
            void Restore()
            {
                _snapshot.Layout = saved.Layout;
                _snapshot.Inserted.Clear(); _snapshot.Inserted.AddRange(saved.Inserted);
                _snapshot.LayoutChanges.Clear(); _snapshot.LayoutChanges.AddRange(saved.Changes);
                _spacingBase = saved.Base;
                _snapshot.Revision = saved.Revision;
            }
            var spacingBase = _spacingBase.Value;
            _snapshot.Layout = new XElement(spacingBase.Layout);
            _snapshot.Inserted.RemoveRange(spacingBase.Inserted, _snapshot.Inserted.Count - spacingBase.Inserted);
            _snapshot.LayoutChanges.RemoveRange(spacingBase.Changes, _snapshot.LayoutChanges.Count - spacingBase.Changes);
            _spacingBase = null;
            object result;
            try { result = tool.Execute(args); }
            catch { Restore(); throw; }
            if (_snapshot.Revision == saved.Revision) Restore();
            return result;
        }

        private object Overview(IDictionary<string, object> args)
        {
            var offset = args.TryGetValue("offset", out var n) ? Convert.ToInt32(n) : 0;
            var items = Ordered();
            return new { snapshot_id = _snapshot.SnapshotId, page_title = _snapshot.Title, scope = _snapshot.SelectionOnly ? "selected_paragraphs" : "page",
                draft_revision = _snapshot.Revision, presets = ParagraphStyles.Ids, native_headings = _snapshot.Options.EnableNativeHeadings,
                paragraph_spacing = _snapshot.Options.EnableParagraphSpacing, code_highlight = _snapshot.Options.EnableCodeHighlight,
                languages = _snapshot.Options.EnableCodeHighlight ? Languages : null,
                list_edit = Has("set_list"), tag_edit = Has("set_tag"), markdown_cleanup = Has("strip_markdown"), table_style = Has("set_table_style"),
                table_shadings = Has("set_table_style") ? TableLook.Shadings : null,
                blank_lines = Has("remove_blank_lines"), code_spacing = Has("normalize_code_spacing"), indent = Has("set_indent"), move = Has("move_blocks"), insert = Has("insert_blocks"), text_table = Has("text_to_table"),
                total = items.Count, next_offset = offset + 100 < items.Count ? (int?)(offset + 100) : null,
                // 文本框按结构草稿列出，合并掉的不在其中；structure 表示能不能调整结构（移动、合并等）。
                outlines = _snapshot.Layout.Elements(OneNoteApi.One + "Outline").Select(o =>
                {
                    var id = (string)o.Attribute("objectID") ?? "";
                    var inside = items.Where(x => x.Node.Ancestors(OneNoteApi.One + "Outline").FirstOrDefault() == o).ToList();
                    var first = inside.Select(x => x.Block == null ? x.New.Text : x.Block.Editable || x.Block.CodeCandidate ? VisibleText(x.Block) : null)
                        .FirstOrDefault(t => !string.IsNullOrWhiteSpace(t));
                    return new { container_id = id, structure = _snapshot.EditableOutlines.Contains(id), blocks = inside.Count,
                        summary = first?.Substring(0, Math.Min(40, first.Length)) };
                }).ToArray(),
                // 按结构草稿里的顺序和层级列出，已删除的空行不在其中，新插入的段落带 inserted。container_id 是现在所在的文本框。
                blocks = items.Skip(offset).Take(100).Select(x => x.Block == null
                    ? (object)new { id = x.Id, container_id = ContainerOf(x.Node), parent_id = ParentKey(x.Node), depth = x.Node.Ancestors(OneNoteApi.One + "OE").Count(),
                        inserted = true, blank = x.New.Text.Length == 0, summary = x.New.Text.Substring(0, Math.Min(80, x.New.Text.Length)) }
                    : new { id = x.Id, container_id = ContainerOf(x.Node), parent_id = ParentKey(x.Node), depth = x.Node.Ancestors(OneNoteApi.One + "OE").Count(),
                        table_id = x.Block.TableId, editable = x.Block.Editable, reason = x.Block.ProtectedReason,
                        list = AgentMarks.ListKind(x.Block.Draft), tags = AgentMarks.Describe(x.Block.Draft, _snapshot.DraftTags),
                        summary = x.Block.Editable || x.Block.CodeCandidate ? Summary(VisibleText(x.Block)) : null }).ToArray(),
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
            if (writing && AgentLayout.Find(_snapshot.Layout, b.Id) == null) throw new AiException("目标段落已删除。");
            return b;
        }
        private object Read(IDictionary<string, object> args)
        {
            var ids = ((IList)args["block_ids"]).Cast<string>().ToList();
            if (ids.Distinct().Count() != ids.Count) throw new AiException("目标段落重复。");
            var found = ids.Select(id => _snapshot.Blocks.FirstOrDefault(x => x.Id == id) ?? throw new AiException($"段落 {id} 不存在。")).ToList();
            // 读取没有副作用：按范围读时夹在中间的空行、代码框等受保护段落跳过并说明原因，不让整批失败再重试。
            // strip_markdown 整段删掉的围栏、分隔线已不在结构草稿里。
            bool Present(AgentBlock b) => AgentLayout.Find(_snapshot.Layout, b.Id) != null;
            var blocks = found.Where(b => (b.Editable || b.CodeCandidate) && Present(b)).ToList();
            var skipped = found.Where(b => !blocks.Contains(b)).Select(b => new { id = b.Id, reason = Present(b) ? b.ProtectedReason : "removed" }).ToArray();
            if (blocks.Count == 0) throw new AiException("目标段落都受到保护：" + string.Join("、", skipped.Select(s => $"{s.id}（{s.reason}）")) + "。");
            var page = _snapshot.CreateDraftPage();
            foreach (var b in blocks) b.Read = true;
            return new { snapshot_id = _snapshot.SnapshotId, skipped = skipped.Length > 0 ? skipped : null, blocks = blocks.Select(b => new { id = b.Id, kind = b.CodeCandidate ? "unhighlighted_code" : "text", text = new AgentRichText(AgentLayout.Find(page, b.Id)).Text,
                depth = AgentLayout.Find(page, b.Id).Ancestors(OneNoteApi.One + "OE").Count(), container_id = ContainerOf(AgentLayout.Find(_snapshot.Layout, b.Id), b.ContainerId),
                parent_id = ParentKey(AgentLayout.Find(page, b.Id)),
                table_id = b.TableId, style = Css.Effective(AgentLayout.Find(page, b.Id), page),
                list = AgentMarks.ListKind(b.Draft), tags = AgentMarks.Describe(b.Draft, _snapshot.DraftTags),
                // 有加粗、链接等行内格式时才给原始 HTML；纯文字和 text 一样，不重复输出。
                runs = AgentLayout.Find(page, b.Id).Elements(OneNoteApi.One + "T").Any(t => t.Value.IndexOf('<') >= 0) ? AgentLayout.Find(page, b.Id).Elements(OneNoteApi.One + "T").Select(t => t.Value).ToArray() : null }).ToArray() };
        }

        private object Paragraph(IDictionary<string, object> args)
        {
            var blocks = Targets(args);
            var preset = (string)args["preset_id"];
            var overrides = args.TryGetValue("overrides", out var raw) ? (IDictionary<string, object>)raw : new Dictionary<string, object>();
            if (overrides.TryGetValue("font_family", out var font) && !FontInstalled((string)font)) throw new AiException("指定字体未安装。");
            var drafts = new Dictionary<AgentBlock, XElement>();
            var styles = new XElement(_snapshot.DraftStyles);
            var before = _snapshot.CreateDraftPage();
            foreach (var b in blocks)
            {
                Block(b.Id, true);
                if (preset == "page_title" && AgentCommitter.Find(_snapshot.Page, b.ObjectId).Parent?.Name != OneNoteApi.One + "Title") throw new AiException("页面标题样式只能用于原标题。");
                var draft = new XElement(b.Draft);
                var current = AgentLayout.Find(before, b.Id) ?? throw new AiException("目标段落已删除。");
                ParagraphStyles.Apply(draft, preset, overrides, _snapshot.Options, current.Elements(OneNoteApi.One + "OEChildren").Any());
                if (_snapshot.Options.EnableNativeHeadings && preset != "page_title")
                    draft.SetAttributeValue("quickStyleIndex", ParagraphStyles.EnsureDefinition(styles, ParagraphStyles.Definition(preset, _snapshot.Options)));
                drafts.Add(b, draft);
            }
            var after = _snapshot.CreateDraftPage(formats: drafts, styles: styles);
            var targeted = new HashSet<string>(blocks.Select(b => b.Id));
            var appearanceOnly = new HashSet<string>();
            // 两份草稿页出自同一个结构草稿，只差格式，段落按位置一一对应。页面上可能有重复或缺失的 objectID，不按 ID 找。
            var originals = after.Descendants(OneNoteApi.One + "OE").Zip(before.Descendants(OneNoteApi.One + "OE"), (a, o) => (a, o)).ToDictionary(p => p.a, p => p.o);
            // 从内向外检查，先消除内层目标的影响，避免把无关的外层标题也降级。
            foreach (var b in blocks.OrderByDescending(b => AgentLayout.Find(before, b.Id).Ancestors().Count()))
            {
                var target = AgentLayout.Find(after, b.Id);
                var untouched = target.Descendants(OneNoteApi.One + "OE").Where(e => e.Elements(OneNoteApi.One + "T").Any() && !targeted.Contains(AgentLayout.KeyOf(e) ?? "")).ToList();
                if (untouched.All(e => SameInheritedFormat(e, originals[e], after, before))) continue;
                var native = new AgentRichText(target).Signature(after, true);
                var draft = drafts[b];
                foreach (var name in new[] { "style", "quickStyleIndex" }) draft.SetAttributeValue(name, (string)b.Draft.Attribute(name));
                ParagraphStyles.PinEmphasis(draft, ParagraphStyles.Definition(preset, _snapshot.Options));
                AgentPageSnapshot.CopyFormat(draft, target);
                // 格式草稿带的是未截取的正文，按结构草稿的记录重新去掉代码框间隔删除的空白行。
                AgentCodeSpacing.Apply(target);
                // 仅设置外观必须和原生样式看起来一样，做不到就不改，不把别的外观当成完成。
                if (new AgentRichText(target).Signature(after, true) != native) throw new AiException($"段落 {b.Id} 无法在保留下级段落格式的同时设置这种外观，没有修改。");
                appearanceOnly.Add(b.Id);
            }
            var result = Publish(drafts, styles: styles, appearanceOnly: appearanceOnly);
            foreach (var b in blocks) b.AppearanceOnly = appearanceOnly.Contains(b.Id);
            return result;
        }

        /// <param name="original">node 在修改前草稿页 before 里对应的段落。</param>
        private static bool SameInheritedFormat(XElement node, XElement original, XElement page, XElement before)
        {
            try { return AgentPageSnapshot.SemanticFormat(original, before) == AgentPageSnapshot.SemanticFormat(node, page); }
            catch (Exception ex) when (ex is System.Xml.XmlException || ex is AiException || ex is ArgumentException)
            {
                // 未指定的未知 HTML 本身没动，仍须检查它继承的格式，不能因解析不了而放过原生样式的影响。
                return Css.Write(Css.Effective(original, before)) == Css.Write(Css.Effective(node, page));
            }
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
                var start = LocateVisible(b, rich.Text, quote, Convert.ToInt32(item["occurrence"]));
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
            foreach (var b in Targets(args).Select(MarkTarget)) drafts.Add(b, new XElement(b.Draft));
            foreach (var draft in drafts.Values) AgentMarks.SetList(draft, kind);
            NormalizeMarkdownNumbers(drafts);
            return Publish(drafts);
        }

        /// <summary>按这次调用后的列表状态重算 Markdown 起点，包括先前分批恢复的相邻项；原生编号不动。</summary>
        private void NormalizeMarkdownNumbers(Dictionary<AgentBlock, XElement> drafts)
        {
            var blocks = _snapshot.Blocks.ToDictionary(b => b.Id);
            XElement Draft(AgentBlock b) => drafts.TryGetValue(b, out var d) ? d : b.Draft;
            foreach (var item in Ordered())
            {
                var b = item.Block;
                if (b?.MarkdownList?.Number == null || !b.Editable || !b.Read || b.Conversion != null || AgentMarks.ListKind(b.Original) == "number") continue;
                var number = Draft(b).Element(OneNoteApi.One + "List")?.Element(OneNoteApi.One + "Number");
                if ((string)number?.Attribute("numberSequence") != "0" || (string)number.Attribute("numberFormat") != "##.") continue;
                var previous = item.Node.ElementsBeforeSelf(OneNoteApi.One + "OE").LastOrDefault();
                var key = previous == null ? null : AgentLayout.KeyOf(previous);
                var prior = key != null && blocks.TryGetValue(key, out var p) ? p : null;
                var priorNumber = prior == null ? null : Draft(prior).Element(OneNoteApi.One + "List")?.Element(OneNoteApi.One + "Number");
                var marker = prior?.MarkdownList;
                // 仅同一父节点、同层的连续 Markdown 列表继续编号；连续的 1. 写法也按 Markdown 惯例递增。
                var continues = marker?.Number != null && prior.Conversion == null &&
                    (string)priorNumber?.Attribute("numberSequence") == "0" && (string)priorNumber.Attribute("numberFormat") == "##." &&
                    marker.Indent == b.MarkdownList.Indent && marker.Quote == b.MarkdownList.Quote &&
                    (b.MarkdownList.Number == marker.Number + 1 || b.MarkdownList.Number == 1 && marker.Number == 1);
                int? start = continues ? (int?)null : b.MarkdownList.Number;
                if ((string)number.Attribute("restartNumberingAt") == start?.ToString(CultureInfo.InvariantCulture)) continue;
                // 不在本批目标里的后项也可能要清掉临时起点；只改副本，一起经 Publish 核验、发布。
                if (!drafts.TryGetValue(b, out var draft)) drafts.Add(b, draft = new XElement(b.Draft));
                draft.Element(OneNoteApi.One + "List").Element(OneNoteApi.One + "Number").SetAttributeValue("restartNumberingAt", start);
            }
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
            return Publish(drafts, tags: definitions);
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

        /// <summary>
        /// 在模型读到的文字里定位，返回草稿全文中的位置。代码框间隔删掉的首尾空白行仍在草稿里、提交时才截掉，
        /// 不能参与出现序号的计数，也不能被改成非空白（否则删行记录失效）。
        /// </summary>
        private int LocateVisible(AgentBlock b, string text, string quote, int occurrence)
        {
            var (from, to) = AgentCodeSpacing.Visible(text, AgentLayout.Find(_snapshot.Layout, b.Id));
            return from + Locate(text.Substring(from, to - from), quote, occurrence);
        }

        private static string Summary(string text) => text.Substring(0, Math.Min(80, text.Length));

        /// <summary>草稿文字去掉代码框间隔删掉的首尾空白行，和 read_blocks 返回的一致。</summary>
        private string VisibleText(AgentBlock b)
        {
            var text = b.CurrentText;
            var (from, to) = AgentCodeSpacing.Visible(text, AgentLayout.Find(_snapshot.Layout, b.Id));
            return text.Substring(from, to - from);
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
                var start = LocateVisible(b, text, quote, Convert.ToInt32(item["occurrence"]));
                new AgentRichText(draft).Replace(start, quote.Length, replacement);
                if (new AgentRichText(draft).Text != text.Substring(0, start) + replacement + text.Substring(start + quote.Length))
                    throw new AiException("修正后的文字与预期不一致。");
                fixes[b].Add($"「{quote}」→「{replacement}」");
            }
            // 全部校验通过后才发布，失败时这个工具没有副作用。
            foreach (var pair in drafts) { pair.Key.Draft = pair.Value; pair.Key.TextFixes.AddRange(fixes[pair.Key]); }
            _snapshot.Revision++;
            return new { ok = true, draft_revision = _snapshot.Revision, changed = drafts.Keys.Select(b => b.Id).ToArray(),
                blocks = drafts.Keys.Select(b => new { id = b.Id, text = VisibleText(b) }).ToArray() };
        }

        /// <summary>转换或删掉的段落不再逐段写回；需要保留的草稿须先存入转换记录。</summary>
        private static void Discard(AgentBlock b)
        {
            b.Draft = new XElement(b.Original); b.TextFixes.Clear(); b.MarkdownMarks = 0; b.MarkdownList = null; b.AppearanceOnly = false;
        }

        /// <summary>
        /// 在初始完整页面上按文本框（单元格）跟踪围栏，包括未选中的上下文；只缓存目标段落的角色，不向模型返回选区外文字。
        /// 任一段 HTML 无法解析时，该文本流不参与 Markdown 清理，避免漏掉围栏而误改代码。
        /// </summary>
        private void CacheFenceRoles()
        {
            foreach (var flow in _snapshot.Layout.Descendants(OneNoteApi.One + "OE").Where(oe => oe.Elements(OneNoteApi.One + "T").Any())
                .GroupBy(PageEditor.TextBlockOf))
            {
                var fences = new AgentMarkdown.Fences();
                var roles = new Dictionary<string, List<AgentMarkdown.LineRole>>();
                try
                {
                    foreach (var oe in flow)
                    {
                        var lines = new AgentRichText(oe).Text.Split('\n').Select(fences.Next).ToList();
                        var key = AgentLayout.KeyOf(oe);
                        if (key != null) roles.Add(key, lines);
                    }
                }
                catch (Exception ex) when (ex is System.Xml.XmlException || ex is AiException || ex is ArgumentException) { continue; }
                foreach (var pair in roles) _markdownRoles.Add(pair.Key, pair.Value);
            }
        }

        /// <summary>
        /// 去掉 Markdown 标记：只删标记字符，逐处删除，其余文字的格式和链接不变，和 fix_text 一样写回时核验、撤销时还原。
        /// 整段只有 ``` 围栏或 --- 分隔线的段落在结构草稿里删掉（和删空行一样受 EnableBlankLineRemoval 控制）；删不了的清空文字。
        /// 受保护、代码、已删的段落跳过并说明原因；单个段落改不了也只跳过这一段。
        /// </summary>
        private object StripMarkdown(IDictionary<string, object> args)
        {
            var kinds = new HashSet<string>(args.TryGetValue("kinds", out var only) ? ((IList)only).Cast<string>() : AgentMarkdown.Kinds);
            var format = args.TryGetValue("emphasis", out var emphasis) && (string)emphasis == "format";
            var targets = new List<AgentBlock>();
            var skipped = new List<object>();
            foreach (var id in Ids(args, "block_ids"))
            {
                var b = _snapshot.Blocks.FirstOrDefault(x => x.Id == id) ?? throw new AiException($"段落 {id} 不存在。");
                var reason = !b.Editable ? b.ProtectedReason : b.Conversion != null ? "converted"
                    : AgentLayout.Find(_snapshot.Layout, b.Id) == null ? "removed" : null;
                if (reason != null) { skipped.Add(new { id, reason }); continue; }
                if (!b.Read) throw new AiException("请先完整读取目标段落。");
                if (!_markdownRoles.ContainsKey(b.Id)) { skipped.Add(new { id, reason = "markdown_context_unavailable" }); continue; }
                targets.Add(b);
            }
            var results = targets.ToDictionary(b => b, b => AgentMarkdown.Analyze(b.CurrentText, _markdownRoles[b.Id], kinds, format));
            var noop = targets.Where(b => results[b].Marks.Count == 0 && results[b].CodeLines == 0).Select(b => b.Id).ToList();
            var codeLines = targets.Where(b => results[b].CodeLines > 0).Select(b => b.Id).ToList();

            // 整段是围栏或分隔线：能调整结构时删掉这一段，一个文本框（单元格）里至少留一段；不能删的清空文字，标题里的不动。
            var separators = targets.Where(b => results[b].Separator).ToList();
            var emptied = new List<AgentBlock>();
            var candidate = new XElement(_snapshot.Layout);
            var removable = new List<(AgentBlock Block, XElement Node)>();
            foreach (var b in separators)
            {
                var node = AgentLayout.Find(candidate, b.Id);
                if (PageEditor.TextBlockOf(node)?.Name == OneNoteApi.One + "Title") { skipped.Add(new { id = b.Id, reason = "title" }); results.Remove(b); continue; }
                // ContainerOf 认的是结构草稿本身，不是这里的副本。
                if (_snapshot.Options.EnableBlankLineRemoval && _snapshot.EditableOutlines.Contains(ContainerOf(AgentLayout.Find(_snapshot.Layout, b.Id))) &&
                    node.Elements().All(e => e.Name == OneNoteApi.One + "T" || e.Name == OneNoteApi.One + "Meta"))
                    removable.Add((b, node));
                else emptied.Add(b);
            }
            foreach (var flow in removable.GroupBy(r => PageEditor.TextBlockOf(r.Node)).ToList())
                if (flow.Key.Descendants(OneNoteApi.One + "OE").Count() == flow.Count())
                {
                    removable.Remove(flow.First());
                    emptied.Add(flow.First().Block);
                }
            var removed = new List<AgentBlock>();
            if (removable.Count > 0)
            {
                var changes = removable.Select(r => (OutlineOf(r.Node), (string)null, r.Block.Id)).ToList();
                foreach (var r in removable) AgentLayout.Detach(r.Node);
                // 结构校验不过（比如会越出选区）时退回清空文字，不让整批失败。
                try { PublishLayout(candidate, "markdown", changes); removed.AddRange(removable.Select(r => r.Block)); }
                catch (AiException) { emptied.AddRange(removable.Select(r => r.Block)); }
            }

            // 逐段在草稿副本上删标记：先加格式（不改文字，下标不变），再从后往前删，最后按纯文字核对。
            var drafts = new Dictionary<AgentBlock, XElement>();
            foreach (var b in targets.Where(b => results.ContainsKey(b) && results[b].Marks.Count > 0 && !removed.Contains(b)))
            {
                var result = results[b];
                var ordered = result.Marks.OrderBy(m => m.Start).ToList();
                var draft = new XElement(b.Draft);
                try
                {
                    for (var i = 1; i < ordered.Count; i++)
                        if (ordered[i].Start < ordered[i - 1].Start + ordered[i - 1].Length) throw new AiException("标记范围重叠。");
                    var text = new AgentRichText(draft).Text;
                    var rich = new AgentRichText(draft);
                    foreach (var (start, length, css) in result.Formats) rich.Format(start, length, css);
                    foreach (var (start, length) in ordered.AsEnumerable().Reverse()) rich.Replace(start, length, "");
                    if (new AgentRichText(draft).Text != AgentMarkdown.Remove(text, ordered))
                        throw new AiException("去掉标记后的文字与预期不一致。");
                }
                catch (AiException ex) { skipped.Add(new { id = b.Id, reason = ex.Message }); emptied.Remove(b); continue; }
                drafts.Add(b, draft);
            }
            // 全部算好才发布；整段删掉的段落不再有草稿。
            foreach (var b in removed) Discard(b);
            foreach (var pair in drafts)
            {
                pair.Key.Draft = pair.Value;
                pair.Key.MarkdownMarks += results[pair.Key].Marks.Count;
                if (results[pair.Key].List != null) pair.Key.MarkdownList = results[pair.Key];
            }
            if (drafts.Count > 0) _snapshot.Revision++;
            return new
            {
                ok = true, draft_revision = _snapshot.Revision,
                changed = drafts.Keys.Where(b => !emptied.Contains(b)).Select(b =>
                {
                    var r = results[b];
                    return new { id = b.Id, heading = r.Heading, list = r.List, todo = r.Todo, quote = r.Quote ? (bool?)true : null,
                        indent = r.Indent > 0 ? (int?)r.Indent : null, code_lines = r.CodeLines > 0 ? (int?)r.CodeLines : null };
                }).ToArray(),
                removed_lines = removed.Select(b => b.Id).ToArray(),
                emptied = emptied.Where(drafts.ContainsKey).Select(b => b.Id).ToArray(),
                code_lines = codeLines.Where(id => !drafts.Keys.Any(b => b.Id == id)).ToArray(),
                noop, skipped
            };
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
            foreach (var b in ordered) { Discard(b); b.Conversion = conversion; }
            _snapshot.CodeConversions.Add(conversion);
            _snapshot.Revision++;
            return new { ok = true, draft_revision = _snapshot.Revision, changed = ordered.Select(b => b.Id).ToArray(), noop = new string[0],
                language = language.Id, discarded_format = discarded };
        }

        /// <summary>把连续段落按分隔符或模型给出的单元格排入草稿，提交时换成表格；和代码框一样整段替换，撤销时换回原段落。</summary>
        private object TextToTable(IDictionary<string, object> args)
        {
            var byRows = args.TryGetValue("rows", out var given);
            if (byRows == args.ContainsKey("delimiter")) throw new AiException("delimiter 和 rows 须二选一：规整的文本用 delimiter，其他用 rows 逐格给出。");
            if (byRows && args.ContainsKey("lines_per_row")) throw new AiException("lines_per_row 只能配合 delimiter 使用；用 rows 时每行直接给出一条记录的全部单元格。");
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
            // Markdown 清理后的转换使用当前草稿；原始 selection 仍用于冲突、结构核验及撤销。
            var preserve = blocks.Any(b => b.MarkdownMarks > 0);
            var source = preserve ? AgentCode.Select(_snapshot.CreateDraftPage(), blocks.Select(b => b.ObjectId).ToList(), "表格") : selection;
            var shading = args.TryGetValue("header_shading", out var shade) && (string)shade != "none" ? TableLook.Shade((string)shade) : null;
            var names = args.TryGetValue("header", out var header) ? ((IList)header).Cast<string>().ToList() : null;
            var headerRow = !args.TryGetValue("header_row", out var hasHeader) || (bool)hasHeader;
            var borders = !args.TryGetValue("borders", out var bordered) || (bool)bordered;
            int padded;
            var table = byRows
                ? AgentTextTable.BuildFromRows(source.Paragraphs, ((IList)given).Cast<IList>().Select(r => (IList<string>)r.Cast<string>().ToList()).ToList(),
                    headerRow, borders, shading, names, out padded)
                : AgentTextTable.Build(source.Paragraphs, (string)args["delimiter"], args.TryGetValue("lines_per_row", out var group) ? Convert.ToInt32(group) : 1,
                    headerRow, borders, shading, names, out padded);
            var ordered = selection.Paragraphs.Select(oe => blocks.First(b => b.Id == AgentLayout.KeyOf(oe))).ToList();
            var conversion = new AgentCodeConversion { Blocks = ordered, TextTable = true, Code = selection.Code, Table = table,
                MarkdownMarks = preserve ? ordered.Sum(b => b.MarkdownMarks) : 0,
                TextFixes = preserve ? ordered.SelectMany(b => b.TextFixes).ToList() : new List<string>() };
            // 清理后的草稿已经进入表格，不再逐段写回；其他转换仍丢弃文字和格式草稿。
            var discarded = preserve ? new string[0] : ordered.Where(b => b.Changed).Select(b => b.Id).ToArray();
            foreach (var b in ordered) { Discard(b); b.Conversion = conversion; }
            _snapshot.CodeConversions.Add(conversion);
            _snapshot.Revision++;
            return new { ok = true, draft_revision = _snapshot.Revision, changed = ordered.Select(b => b.Id).ToArray(), rows = table.Elements(OneNoteApi.One + "Row").Count(),
                columns = table.Element(OneNoteApi.One + "Columns").Elements().Count(), padded_rows = padded, discarded_format = discarded };
        }

        private object RemoveBlankLines(IDictionary<string, object> args)
        {
            var blanks = new HashSet<string>(_snapshot.Blocks.Where(b => b.ProtectedReason == "empty" && b.Conversion == null).Select(b => b.Id)
                .Concat(_snapshot.Inserted.Where(i => i.Text.Length == 0).Select(i => i.Id)));
            var candidate = new XElement(_snapshot.Layout);
            // 按草稿里的当前位置判断范围；插入历史保留，避免清理后复用 ID 或释放本次任务的插入配额。
            var outlines = candidate.Elements(OneNoteApi.One + "Outline").Where(o => _snapshot.EditableOutlines.Contains((string)o.Attribute("objectID")))
                .Descendants(OneNoteApi.One + "OE").Where(e => blanks.Contains(AgentLayout.KeyOf(e) ?? "") && _snapshot.InSelection(e)).ToDictionary(e => e, OutlineOf);
            var removed = AgentLayout.RemoveBlankLines(candidate, outlines.ContainsKey, (string)args["mode"] == "all");
            PublishLayout(candidate, "removed", removed.Select(e => (outlines[e], (string)null, AgentLayout.KeyOf(e))).ToList());
            return new { ok = true, draft_revision = _snapshot.Revision, removed = removed.Select(AgentLayout.KeyOf).ToArray() };
        }

        private object NormalizeCodeSpacing(IDictionary<string, object> args)
        {
            var result = AgentCodeSpacing.Build(_snapshot);
            var changed = PublishSpacing(result);
            return new { ok = true, draft_revision = _snapshot.Revision, removed_paragraphs = result.Removed,
                removed_soft_lines = result.SoftLines, inserted_paragraphs = result.Inserted.Count, noop = !changed,
                skipped = result.Skipped.Select(s => s.ToResult()).ToArray(), skip_message = AgentCodeSpacing.SkipMessage(result.Skipped) };
        }

        /// <summary>发布间隔调整。首次产生删补时记下调整前的草稿，后续结构工具据此撤回。</summary>
        private bool PublishSpacing(AgentCodeSpacing.Result result)
        {
            var before = (_snapshot.Layout, _snapshot.Inserted.Count, _snapshot.LayoutChanges.Count);
            var changed = PublishLayout(result.Layout, result.Changes);
            if (changed && _spacingBase == null) _spacingBase = before;
            _snapshot.Inserted.AddRange(result.Inserted);
            _snapshot.CodeSpacingRequested = true;
            return changed;
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
            var prepared = PrepareMerge(_snapshot.Layout, sourceId, (string)args["target_id"], (string)args["position"] == "before");
            // 在发布前完成只读预检：取消或非预期异常不能留下已发布的部分草稿。
            var left = MergeableOutlines(prepared.Layout, prepared.Into);
            PublishLayout(prepared.Layout, new[] { new AgentLayoutChange { Kind = "merged", OutlineId = prepared.Into, From = sourceId, Ids = prepared.Moved } });
            return new { ok = true, draft_revision = _snapshot.Revision, merged = sourceId, into = prepared.Into, moved = prepared.Moved,
                mergeable_left = left.Length, mergeable_outline_ids = left };
        }

        /// <summary>实际合并与剩余计数共用的准备：只修改布局副本，核验全部不变量，不发布草稿。</summary>
        private (XElement Layout, string Into, string[] Moved) PrepareMerge(XElement layout, string sourceId, string targetId, bool before,
            Action<string> formatMismatch = null)
        {
            _cancellation.ThrowIfCancellationRequested();
            var candidate = new XElement(layout);
            var source = candidate.Elements(OneNoteApi.One + "Outline").FirstOrDefault(o => (string)o.Attribute("objectID") == sourceId);
            if (source == null) throw new AiException("源文本框不存在或已经合并。");
            if (!_snapshot.EditableOutlines.Contains(sourceId)) throw new AiException("源文本框受到保护，不能调整结构。");
            var target = Node(candidate, targetId, out _);
            if (target.Ancestors(OneNoteApi.One + "Outline").First() == source) throw new AiException("目标段落不能在源文本框里。");
            if (PageEditor.TextBlockOf(target).Name != OneNoteApi.One + "Outline") throw new AiException("目标段落不能在表格单元格里。");
            if (HasUnselected(source)) throw new AiException("源文本框里有没选中的段落，不能整体合并。");
            var moved = source.Descendants(OneNoteApi.One + "OE").Select(AgentLayout.KeyOf).Where(k => k != null).ToArray();
            var into = OutlineOf(target);
            AgentLayout.Merge(source, target, before);
            _snapshot.CheckLayout(candidate, formatMismatch);
            _cancellation.ThrowIfCancellationRequested();
            return (candidate, into, moved);
        }

        /// <summary>只处理选中范围时，没选中的段落不在快照里，不能跟着整框搬走。</summary>
        private static bool HasUnselected(XElement outline) =>
            outline.Descendants(OneNoteApi.One + "OE").Any(e => e.Elements(OneNoteApi.One + "T").Any() && AgentLayout.KeyOf(e) == null);

        /// <summary>按页面顺序列出还能并进 into 的文本框；目标框内任一合法段落的前后通过实际合并校验才计入。</summary>
        private string[] MergeableOutlines(XElement layout, string into)
        {
            _cancellation.ThrowIfCancellationRequested();
            // 必须取完整格式草稿：待应用的父段格式、QuickStyleDef 也影响新位置的继承环境。
            var draft = _snapshot.CreateDraftPage(layout);
            var draftNodes = draft.Descendants(OneNoteApi.One + "OE").Where(e => AgentLayout.KeyOf(e) != null).ToDictionary(AgentLayout.KeyOf);
            var targets = new List<(string Id, string InheritedStyle)>();
            var targetParents = new HashSet<XElement>();
            var destination = layout.Elements(OneNoteApi.One + "Outline").First(o => (string)o.Attribute("objectID") == into);
            foreach (var id in destination.Descendants(OneNoteApi.One + "OE").Select(AgentLayout.KeyOf).Where(id => id != null))
            {
                _cancellation.ThrowIfCancellationRequested();
                try
                {
                    var node = Node(layout, id, out _);
                    // 同一父节点下的合法目标都不在转换范围内；整框插到其前后具有相同的祖先和继承环境。
                    // 每个父节点只预检一组代表位置，避免样式不兼容时把一整框的兄弟段落反复核验。
                    if (PageEditor.TextBlockOf(node).Name == OneNoteApi.One + "Outline" && targetParents.Add(node.Parent))
                        targets.Add((id, Css.Write(Css.Effective(draftNodes[id].Parent, draft))));
                }
                catch (AiException) { /* 已转换的段落或代码框不能作为目标。 */ }
            }
            var left = new List<string>();
            foreach (var source in layout.Elements(OneNoteApi.One + "Outline"))
            {
                _cancellation.ThrowIfCancellationRequested();
                var id = (string)source.Attribute("objectID");
                if (id == null || id == into || !_snapshot.EditableOutlines.Contains(id) || HasUnselected(source)) continue;
                if (CanMerge(source)) left.Add(id);
            }
            return left.ToArray();

            bool CanMerge(XElement source)
            {
                var sourceId = (string)source.Attribute("objectID");
                var sourceBlocks = new HashSet<string>(source.Descendants(OneNoteApi.One + "OE").Select(AgentLayout.KeyOf).Where(id => id != null));
                var rejectedStyles = new HashSet<string>(StringComparer.Ordinal);
                foreach (var target in targets)
                {
                    _cancellation.ThrowIfCancellationRequested();
                    if (rejectedStyles.Contains(target.InheritedStyle)) continue;
                    foreach (var before in new[] { true, false })
                    {
                        _cancellation.ThrowIfCancellationRequested();
                        string mismatch = null;
                        try { PrepareMerge(layout, sourceId, target.Id, before, id => mismatch = id); return true; }
                        catch (AiException)
                        {
                            // 只复用已由完整检查确认的源段格式失败。相同继承环境下源段的字符格式必然相同；
                            // 选区、转换和其他段落的失败不能据此跳过位置，成功候选仍走完整检查。
                            if (mismatch != null && sourceBlocks.Contains(mismatch))
                            {
                                rejectedStyles.Add(target.InheritedStyle);
                                break;
                            }
                        }
                    }
                }
                return false;
            }
        }

        private static readonly char[] Breaks = { '\n', '\r', '\t' };

        private object InsertBlocks(IDictionary<string, object> args)
        {
            var candidate = new XElement(_snapshot.Layout);
            var target = Node(candidate, (string)args["target_id"], out _);
            var items = ((IList)args["paragraphs"]).Cast<IDictionary<string, object>>().ToList();
            var blanks = items.Select(i => i.TryGetValue("blank", out var b) && (bool)b).ToList();
            for (var i = 0; i < items.Count; i++)
            {
                if (blanks[i] && new[] { "text", "preset_id", "list" }.Any(items[i].ContainsKey)) throw new AiException("空行项只写 blank: true，不要带 text、preset_id 或 list。");
                if (!blanks[i] && !(items[i].ContainsKey("text") && items[i].ContainsKey("preset_id"))) throw new AiException("每项要么写 text 和 preset_id，要么写 blank: true 插入空行。");
            }
            // 空行的文字记为空，和代码框间隔补的空行一样。
            var texts = items.Select((i, n) => blanks[n] ? "" : ((string)i["text"]).Trim()).ToList();
            if (texts.Where((t, n) => !blanks[n]).Any(t => t.Length == 0)) throw new AiException("插入的段落不能是空白；要插空行请写 blank: true。");
            if (texts.Any(t => t.IndexOfAny(Breaks) >= 0)) throw new AiException("插入的文字不能含换行或制表符，每段单独写一项。");
            if (_snapshot.Inserted.Count + items.Count > MaxInserted || _snapshot.Inserted.Sum(i => i.Text.Length) + texts.Sum(t => t.Length) > MaxInsertedChars)
                throw new AiException($"每个任务最多插入 {MaxInserted} 段、{MaxInsertedChars} 字。");
            var styles = new XElement(_snapshot.DraftStyles);
            var outline = OutlineOf(target);
            var created = items.Select((item, i) =>
            {
                var id = "n" + (_snapshot.Inserted.Count + i + 1);
                return blanks[i] ? AgentLayout.NewBlankLine(id, _snapshot.Options, styles)
                    : AgentLayout.NewParagraph(id, texts[i], (string)item["preset_id"], item.TryGetValue("list", out var list) ? (string)list : null, _snapshot.Options, styles);
            }).ToList();
            if ((string)args["position"] == "before") target.AddBeforeSelf(created); else target.AddAfterSelf(created);
            // 空行记 inserted_blank：提交时按空行核验（允许 OneNote 把 &nbsp; 回存成空 T），结果计入补空行。
            var added = created.Where((e, i) => !blanks[i]).Select(AgentLayout.KeyOf).ToArray();
            var blankIds = created.Where((e, i) => blanks[i]).Select(AgentLayout.KeyOf).ToArray();
            PublishLayout(candidate, new[] { ("inserted", added), ("inserted_blank", blankIds) }.Where(c => c.Item2.Length > 0)
                .Select(c => new AgentLayoutChange { Kind = c.Item1, OutlineId = outline, Ids = c.Item2 }).ToList());
            _snapshot.DraftStyles.ReplaceNodes(styles.Elements().Select(e => new XElement(e)));
            _snapshot.Inserted.AddRange(created.Select((e, i) => new AgentInserted { Id = AgentLayout.KeyOf(e), OutlineId = outline, Text = texts[i] }));
            return new { ok = true, draft_revision = _snapshot.Revision, inserted = created.Select(AgentLayout.KeyOf).ToArray(), blank_lines = blankIds };
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
            _cancellation.ThrowIfCancellationRequested();
            _snapshot.Layout = candidate;
            _snapshot.LayoutChanges.AddRange(changes);
            _snapshot.Revision++;
            return true;
        }

        private object Publish(Dictionary<AgentBlock, XElement> drafts, XElement styles = null, XElement tags = null, ISet<string> appearanceOnly = null)
        {
            var changed = new List<string>();
            var noop = new List<string>();
            var before = _snapshot.CreateDraftPage();
            var after = _snapshot.CreateDraftPage(formats: drafts, styles: styles, tags: tags);
            // 先验证全部段落，再一次性发布草稿。失败时这个工具没有副作用。
            foreach (var pair in drafts)
            {
                var old = AgentLayout.Find(before, pair.Key.Id) ?? throw new AiException("目标段落已删除。");
                var proposed = AgentLayout.Find(after, pair.Key.Id);
                if (new AgentRichText(old).Signature(before, false) != new AgentRichText(proposed).Signature(after, false))
                    throw new AiException("格式工具不能改变正文或链接。");
                var expected = AgentPageSnapshot.SemanticFormat(proposed, after);
                // 同批父段可能改变继承值。即使最终外观与原来相同，抵消继承变化的显式样式也必须发布。
                AgentPageSnapshot.CopyFormat(old, proposed);
                var without = AgentPageSnapshot.SemanticFormat(proposed, after);
                AgentPageSnapshot.CopyFormat(pair.Value, proposed);
                if ((string)pair.Key.Draft.Attribute("quickStyleIndex") == (string)pair.Value.Attribute("quickStyleIndex") &&
                    AgentMarks.DraftKey(pair.Key.Draft) == AgentMarks.DraftKey(pair.Value) &&
                    AgentPageSnapshot.SemanticFormat(old, before) == expected && without == expected) noop.Add(pair.Key.Id);
                else changed.Add(pair.Key.Id);
            }
            foreach (var pair in drafts.Where(p => changed.Contains(p.Key.Id))) pair.Key.Draft = pair.Value;
            if (styles != null) _snapshot.DraftStyles.ReplaceNodes(styles.Elements().Select(e => new XElement(e)));
            if (tags != null) _snapshot.DraftTags.ReplaceNodes(tags.Elements().Select(e => new XElement(e)));
            if (changed.Count > 0) _snapshot.Revision++;
            return new { ok = true, draft_revision = _snapshot.Revision, changed, noop, appearance_only = appearanceOnly?.ToArray() };
        }
        /// <summary>按当前草稿顺序计算；删除、转换和新插入的段落不再需要读取。</summary>
        private string[] UnreadBlockIds() => Ordered().Where(x => x.Block != null &&
            (x.Block.Editable || x.Block.CodeCandidate) && !x.Block.Read && x.Block.Conversion == null).Select(x => x.Id).ToArray();

        private object Pending(IDictionary<string, object> args)
        {
            var unread = UnreadBlockIds();
            var present = new HashSet<string>(_snapshot.Layout.Descendants(OneNoteApi.One + "OE").Select(AgentLayout.KeyOf));
            return new { snapshot_id = _snapshot.SnapshotId, draft_revision = _snapshot.Revision,
                changed = _snapshot.Blocks.Where(b => b.Changed).Select(b => b.Id).ToArray(),
                text_fixes = _snapshot.Blocks.Where(b => b.TextFixes.Count > 0).Select(b => new { id = b.Id, fixes = b.TextFixes.ToArray() }).ToArray(),
                markdown_stripped = _snapshot.Blocks.Where(b => b.MarkdownMarks > 0).Select(b => b.Id).ToArray(),
                appearance_only = _snapshot.Blocks.Where(b => b.Changed && b.AppearanceOnly).Select(b => b.Id).ToArray(),
                unread, unread_count = unread.Length, next_read_block_ids = unread.Take(ReadBatchSize).ToArray(),
                code_blocks = _snapshot.CodeConversions.Where(c => !c.TextTable).Select(c => new { block_ids = c.Blocks.Select(b => b.Id).ToArray(), language = c.LanguageId }).ToArray(),
                text_tables = _snapshot.CodeConversions.Where(c => c.TextTable).Select(c => new { block_ids = c.Blocks.Select(b => b.Id).ToArray() }).ToArray(),
                layout = new { removed = LayoutIds("removed"), markdown_removed = LayoutIds("markdown"), moved = LayoutIds("moved"), indented = LayoutIds("indented"),
                    code_spacing = _snapshot.CodeSpacingRequested, soft_blank_lines = _snapshot.Layout.Descendants(OneNoteApi.One + "OE").Sum(AgentCodeSpacing.TrimCount),
                inserted = _snapshot.Inserted.Where(i => present.Contains(i.Id)).Select(i => new { id = i.Id, text = i.Text }).ToArray(),
                merged = _snapshot.LayoutChanges.Where(c => c.Kind == "merged").Select(c => new { from = c.From, into = c.OutlineId }).ToArray() },
                unconverted_code = _snapshot.Blocks.Where(b => b.CodeCandidate && b.Conversion == null).Select(b => b.Id).ToArray(),
                tables_changed = _snapshot.Tables.Where(t => t.Changed).Select(t => t.Id).ToArray(),
                protected_count = _snapshot.Blocks.Count(b => !b.Editable && b.Conversion == null) };
        }
        private string[] LayoutIds(string kind) => _snapshot.LayoutChanges.Where(c => c.Kind == kind).SelectMany(c => c.Ids).Distinct().ToArray();

        private object Finish(IDictionary<string, object> args)
        {
            if (Convert.ToInt32(args["draft_revision"]) != _snapshot.Revision) throw new AiException("草稿修订号过期，请先检查待提交修改。");
            var unread = UnreadBlockIds();
            if (unread.Length > 0 && !AllowIncompleteFinish)
                return new { ok = false, error = $"范围内还有 {unread.Length} 段未完整读取，未提交草稿。请分批 read_blocks 后重试。",
                    snapshot_id = _snapshot.SnapshotId, draft_revision = _snapshot.Revision,
                    unread_count = unread.Length, next_read_block_ids = unread.Take(ReadBatchSize).ToArray() };
            var spacing = _snapshot.CodeSpacingRequested ? AgentCodeSpacing.Build(_snapshot) : null;
            if (spacing != null && !XNode.DeepEquals(spacing.Layout, _snapshot.Layout))
            {
                if (!AllowIncompleteFinish)
                    return new { ok = false, error = "后续修改改变了代码框间隔，未提交草稿。请重新调用 normalize_code_spacing 后重试。",
                        snapshot_id = _snapshot.SnapshotId, draft_revision = _snapshot.Revision };
                // 最后一轮不能再调用其他工具：按当前草稿重新规范化后提交，而不是丢掉整个任务。
                PublishSpacing(spacing);
            }
            _snapshot.Frozen = true;
            Report = _committer.Commit(_snapshot, _cancellation);
            Report.SpacingSkipped = spacing?.Skipped.Select(s => s.ToResult()).ToArray() ?? new object[0];
            if (Report.SpacingSkipped.Length > 0) Report.Message += $"\n代码框间隔跳过 {Report.SpacingSkipped.Length} 处：" + AgentCodeSpacing.SkipMessage(spacing.Skipped) + "。";
            Report.UnreadCount = unread.Length;
            if (unread.Length > 0)
            {
                if (Report.Status == "Verified" && Report.CanUndo) Report.Status = "PartiallyApplied";
                Report.Message += $"\n仍有 {unread.Length} 段未读取，本次只提交已完成的草稿。";
            }
            return Report.ToToolResult();
        }
    }
}
