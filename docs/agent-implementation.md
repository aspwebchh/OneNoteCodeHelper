# Agent 实施与验证记录

日期：2026-09-26。已实现第一版页面格式 Agent，保持 .NET Framework 4.8、WPF、强名称签名和现有 Office PIA 早绑定引用。没有新增第三方生产依赖。

## 实际执行链路

1. 「开始 → AI 助手 → Agent」通过 `AddIn.OnShowAgentWindow` 启动独立 STA 窗口，固定当前页及打开窗口时的选区。
2. 用户选择整页或选中段落并输入需求。后台读取固定页面，构建短 ID、保护范围、内容和格式指纹。
3. `AgentRunner` 把模型原生 `tool_calls` 映射到本地函数（按能力开关和页面内容注册，最多十八个），完整读取后才能修改段落。模型的文字声明不能触发写回。
4. 格式和改字工具只更新内存草稿。每个批量调用全部校验通过后才发布，返回修订号；参数错误不会留下半个工具调用的修改。
5. 模型独立调用 `finish_edit` 后，提交器在页面锁内重读、识别冲突、重建完整的受影响容器、带时间戳提交，并回读核验。
6. UI 展示实际核验结果，保存本次已确认段落的撤销记录。撤销再次校验指纹，避免覆盖后来编辑的内容。
   撤销记录同时存下原段落引用的 QuickStyleDef、TagDef：OneNote 回存后会给它们重新编号，撤销时按内容对应到当前编号，不照搬原来的 `quickStyleIndex`、标记 `index`。

页面正文作为数据传给模型，不作为系统指令。工具不接受任意 XML、HTML、外部页面 ID、文件路径或脚本。修改范围不能超出快照；图片二进制、代码和受保护对象不传入模型。模型能读取的正文和富文本格式会发送给用户配置的模型接口。

## 工具

本地校验和 JSON Schema 共用 `AgentSchema` 定义；所有对象拒绝未知字段，限制数组长度、数值、枚举和字符串长度。
系统提示词按本次实际注册的工具拼接（`AgentRunner.SystemPrompt`），没有提供的工具不提。

| 工具 | 参数与执行规则 |
|---|---|
| `get_page_overview` | 可选 `offset`，每页最多 100 条；返回快照 ID、短段落 ID、最多 80 字摘要、嵌套深度、保护原因、可用预设与能力开关 |
| `read_blocks` | `snapshot_id`、`block_ids`；最多 100 段，返回完整文字、富文本 runs、当前草稿样式；读取受保护段落会拒绝 |
| `set_paragraph_style` | `snapshot_id`、`block_ids`、`preset_id`、可选 `overrides`；预设为 page_title/heading1/heading2/body/quote，page_title 只用于原标题。原生样式影响未指定子段时保留原有继承属性和样式引用，仅设置父段自身文字外观，返回 `appearance_only` 段落 ID 列表 |
| `set_text_style` | `snapshot_id`、`targets`；每项指定 `block_id`、原文 `quote`、从 1 开始的 `occurrence`、`style`；支持 bold/italic/underline/color |
| `fix_text` | `snapshot_id`、`fixes`；每项指定 `block_id`、原文 `quote`、从 1 开始的 `occurrence`、改后文字 `replacement`，两者各 1–30 字、不含换行，同一段的多处按顺序应用；代码段落拒绝。除 `strip_markdown` 外唯一能改文字的工具，系统提示词要求只在用户要求时修正错别字 |
| `strip_markdown` | `snapshot_id`、`block_ids`（≤1000，须完整读取），可选 `kinds`（heading/quote/list/emphasis/inline_code/fence/rule，默认全部）、`emphasis`（remove 默认，format 同时设粗体、斜体、删除线）。只删标记字符（`AgentMarkdown` 分析，逐处 `Replace` 为空），返回 `changed`（每段原来的 `heading` 级别、`list`、`todo`、`quote`、`indent`）、`removed_lines`、`emptied`、`code_lines`、`noop`、`skipped`（受保护、代码、已删的段落，或单段处理失败的原因）。整段只有围栏、分隔线的段落在结构草稿里删掉（变更种类 `markdown`），需要 `EnableBlankLineRemoval`、可调整结构的文本框、段落只有 T/Meta，且文本框（单元格）里至少留一段；做不到或结构校验不过时清空文字，标题里的不动。围栏按文本框逐行跟踪整页：`markdown`/`md` 围栏里照常处理，其里带语言的围栏算嵌套代码块，其他围栏里的行不动。`EnableMarkdownCleanup=false` 时不注册 |
| `highlight_code` | `snapshot_id`、`block_ids`（最多 1000）、`language`（`auto` 或已支持的语言 id）；把连续的代码段落排入草稿，提交时换成高亮代码框。`Agent/EnableCodeHighlight=false` 时不注册 |
| `set_list` | `snapshot_id`、`block_ids`、`list`（bullet/number/none）；已是同一种列表时不动，保留原符号样式。页面标题、未读和代码段落拒绝。`EnableLists=false` 时不注册 |
| `set_tag` | `snapshot_id`、`block_ids`、`tag`（todo/important/question/none）、可选 `completed`（只用于 todo）；同类标记不重复添加，none 只去掉这三种。`EnableTags=false` 时不注册 |
| `set_table_style` | `snapshot_id`、`table_ids`、`style`（`borders`、`header_row`、`header_shading` 至少一项）；不需要先读。`EnableTableStyles=false` 或页面没有可编辑表格时不注册 |
| `read_image_text` | `snapshot_id`、`image_ids`；只读，返回 OneNote 识别出的图片文字，每张最多 4000 字并标记 `truncated`。页面上有带识别文字的图片时才注册 |
| `remove_blank_lines` | `snapshot_id`、`mode`（collapse/all）；collapse 同「排版优化」，连续空行留一行、删掉文本框和单元格首尾的空行；all 删掉全部，但不会把一摞段落删空。带列表、标记、下级段落的空段落和已排转换里的空行不删。`EnableBlankLineRemoval=false` 或没有可删空行时不注册 |
| `set_indent` | `snapshot_id`、`block_ids`、`direction`（in/out）；in 挂到前一个兄弟段落下，out 移到上级之后、原来排在后面的兄弟段落改挂到它下面，上下顺序不变。上级也在列表里的段落跟着上级走。`EnableIndent=false` 时不注册 |
| `move_blocks` | `snapshot_id`、`block_ids`、`target_id`、`position`（before/after）；连同下级段落按原顺序移到目标前后，成为目标的同级段落。可以移到另一个文本框，单元格里的段落只能在同一单元格里移动，不能把文本框移空。`EnableMoves=false` 时不注册 |
| `merge_outlines` | `snapshot_id`、`source_id`（源文本框的 container_id）、`target_id`、`position`；源文本框的全部顶层段落（含表格、图片、空行）按顺序移到目标前后，提交时删掉源文本框。目标不能在源文本框或单元格里；只处理选区时源文本框里的文字段落必须都在选区内。`EnableMoves=false` 或可调整结构的文本框少于两个时不注册 |
| `insert_blocks` | `snapshot_id`、`target_id`、`position`、`paragraphs`（1–20 项，每项 `text` ≤500 字、`preset_id` 为 heading1/heading2/body/quote、可选 `list`）；纯文字按 HTML 转义，每个任务最多 50 段、5000 字，新段落短 ID 为 n1…。`EnableInsert=false` 时不注册 |
| `text_to_table` | `snapshot_id`、`block_ids`（≤200），`delimiter` 和 `rows` 二选一，可选 `header_row`、`borders`（默认 true）、`header_shading` 和 `header`（首行前新增的列名，每项 ≤30 字）；最多 10 列、100 行。**delimiter**（tab/pipe/space/none，none 为整行一格）、可选 `lines_per_row`（1–10，每条记录占的行数，空行不算；大于 1 时每几行合成一行，记录中非末行的行尾冒号去掉）：每行文字一行（段内 `<br>` 切开的也各成一行），至少分出 2 列；space 按连续空白（含 `&nbsp;`、全角空格）拆分。**rows**（每行 ≤10 格、每格 ≤500 字，空字符串为空单元格，不能含换行）：模型逐格给出切分结果，`BuildFromRows` 把各段文字用 `\n` 连起来按顺序定位每格（空白彼此等价，换行除外），单元格按定位到的范围从原段落截取；原文里没进单元格的字只能是空白和 `\t | ｜ : ： , ， ; ； 、 =`，或去掉这些后与该列 `header` 相同的标签，否则报「没有按顺序找到」或「没有放进任何单元格」。各行按最多的列数补空单元格，header 少了补空、多了加列；返回 `padded_rows`（补了空单元格的行数）。`EnableTextTables=false` 时不注册 |
| `get_pending_changes` | `snapshot_id`；返回草稿修订号、修改段落 ID、已排的文字修正、未完整读取的可编辑段落 ID、已排入的代码框和表格转换、结构改动（删除、移动、缩进、插入）、尚未转换的等宽代码及保护计数 |
| `finish_edit` | `snapshot_id`、`draft_revision`；必须是当轮唯一工具，修订号匹配后冻结草稿，只提交一次 |

段落 overrides：`alignment` 为 left/center/right；`font_family` 为 Microsoft YaHei/Calibri/Arial 且需已安装；字号 8–32 pt；段前后间距 0–36 pt；颜色限于 `#1F4E79`、`#365F91`、`#222222`、`#666666`。
局部文字按精确匹配定位，重复短语的出现序号按不重叠匹配计数。切开代理对、组合字符或常见 emoji 序列的请求会拒绝。
`AgentRichText.Save` 将同一 `one:T`、相同标签路径下连续文字先聚合再序列化，避免把 emoji、补充平面汉字的高低代理项分开输出；多个文字片段、链接、局部格式、硬空格和段内换行仍保留。

格式草稿的修改前后比较在完整结构草稿中进行，包含祖先继承格式和本批新增的样式、标记定义；整批校验通过才发布。取消继承的加粗、斜体、下划线会生成有效覆盖，重复设置不增加修订号。
段落预设使用当前草稿的子段关系（包括先缩进再设置样式），由内向外检查未指定子段的有效格式；只有原生样式确实影响子段时才降级。
`get_pending_changes` 同样返回 `appearance_only`；最终 `finish_edit` 结果只列入回读核验通过的降级段落，冲突跳过或未验证的不计入。撤销恢复原有格式和样式引用。

`get_page_overview` 另外返回每段的 `list`、`tags`、`table_id`，以及 `tables`（t1…，行列数、当前外观、首行前 40 字、是否可编辑）和 `images`（i1…，识别文字字数）；
`read_blocks` 每段也带 `list`、`tags`，`get_pending_changes` 带 `tables_changed`。
概况按结构草稿里的顺序列出段落：已删除的空行不在其中，新插入的段落带 `inserted`；`depth`、`parent_id` 按结构草稿计算，`parent_id` 是上级段落的短 ID。
段落、表格、图片的 `container_id` 是现在所在的文本框；`outlines` 列出结构草稿里的文本框（`structure` 表示能否调整结构、段落数、首段摘要），合并掉的不列。

`fix_text` 在原文和改后文字之间逐字比对：没变的字连同格式、链接原样保留，新字沿用被替换的字（纯插入时沿用前一个字）的格式。修正后 `read_blocks`、概况摘要和局部格式定位都按草稿里的新文字。
提交时只有排过文字修正的段落可以改正文，而且必须和草稿一致；其他段落仍要求正文和链接不变。冲突检测、回读核验照旧，撤销记录带上修正，撤销时文字一起还原。修正清单只进窗口里的结果，日志只记条数。
`strip_markdown` 走同一套机制：段落记下去掉的标记处数（`AgentBlock.MarkdownMarks`，与 `TextFixes` 一起构成 `TextEdited`，写回时允许正文变化），先按原文下标加格式（不改文字），再从后往前删标记，最后按纯文字核对。
结果、撤销消息只报处数（整段删掉的围栏、分隔线各算一处），不列清单。已删的段落 `read_blocks` 跳过（reason=removed），格式和文字工具拒绝。

## 代码框转换

代码段落分三类：单行单格表格里、格内全是等宽段落的视为已有代码框（`highlighted_code`），受保护；整段非空白字符都是等宽字体的是待转换代码（`unhighlighted_code`），可读取、可转换、不能设置样式；只有部分文字等宽的是行内代码（`protected_code`），受保护。普通字体输入的代码仍是可编辑正文，由模型判断后转换。

`highlight_code` 复用「高亮选中」的 `CodeSelection` 拼源码（每深一层补一个制表符）、`LanguageRegistry.Resolve` 定语言、`CodeBlockBuilder.BuildTable` 生成代码框，主题和字号等取打开窗口时的功能区设置。为保证能原样撤销，只接受同一文本块里连续的段落，第一段在最外层，每段的下级段落都在范围内，且不含列表或待办标记；代码中间的空段落可以一并传入。转换会丢弃这些段落已排的格式草稿。

提交时和格式修改在同一次重读、写入、回读中完成。任何源段落的指纹变化都会跳过整个代码框。回读核验中，写入前不存在的 objectID 视为新建对象；新代码行和还原段落只比较规范化的纯文字（硬空格按空格、去掉行尾空白），因为 OneNote 会改写代码行的 span。
核验通过后，以实际回读状态记录转换结果的 Table ID 和撤销指纹：完整 XML 内容、格式、链接、列表、标记、图片和结构，加上所在文本框和上级段落的身份、文本框坐标、祖先继承属性与引用的 QuickStyleDef、TagDef。样式和标记引用按定义内容比较，不看编号，OneNote 重新编号不算改动。代码框与文字转表格共用此保护。
不比较同级序号：在转换结果前后增删段落、增删别处的文本框时原地换回不会覆盖后续编辑，不算冲突。
指纹只排除已知瞬态差异（选中状态、修改时间、纯表格包装段落 ID、未锁定列的自动宽度），并规范化属性顺序及 XML 排版空白；正文空白、锁定列宽仍严格检查。

撤销时转换结果指纹不变才换回原段落；后续编辑触发冲突时整处跳过。所有转换的冲突检查均对照本轮撤销写入前的页面，避免先恢复的段落改变后续目标的位置或继承样式而产生假冲突。
还原时去掉 objectID 和编辑记录让 OneNote 新建，`quickStyleIndex` 按当前页面的样式定义重新对应，再核验文字和格式。有结构调整的转换继续按整框、跨框联动组撤销。

## 列表、标记和表格样式

2026-09-28 追加，三者都不改变段落数量、顺序和嵌套，沿用草稿、提交、回读核验和撤销流程。

- **列表**：新建时写 `<one:List><one:Bullet bullet="2"/></one:List>` 或 `<one:List><one:Number numberSequence="0" numberFormat="##."/></one:List>`，不写字体属性，OneNote 回存时自己补上编号文字。
  列表类型可以在原段落上直接切换；但 OneNote 不接受在原段落上去掉列表（见 README「几个踩过的坑」），所以取消列表和撤销新加的列表时，
  提交器去掉这一段的 objectID 和编辑记录，让 OneNote 重建它（`AgentCode.StripIdentity`）。下级段落保持原 ID；
  回读时重建的段落按位置对应，撤销记录记新 ID。
- **标记**：预设按 OneNote 默认标记库，todo = symbol 3、important = 13、question = 15，本机回存与预期一致。
  草稿里按图标复用页面已有的 TagDef（包括用户的中文「待办事项」），没有才新增；提交时按完整定义对应到重新读取的页面，只有新增了 TagDef 才把 TagDef 一起提交。
  省略 `one:Tag` 可以删掉标记，OneNote 会一并清掉不再引用的 TagDef。
- **核验**：`SemanticFormat` 追加列表种类和「标记图标:完成状态」（`AgentMarks.Projection`），不看 OneNote 补上的字号、编号文字、时间。
  本次写入的段落和重建的段落不逐字比较 List/Tag 的 XML；其余段落仍严格比较，标记按引用的 TagDef 内容比较，不受 TagDef 重新编号影响。
- **表格**：外观为 `bordersVisible`、`hasHeaderRow` 和首行各单元格的 `shadingColor`，逐格记录，撤销能还原各格不同的底色。
  指纹只含表格 ID、外观、行列数和首行单元格 ID，处理期间改单元格文字不算冲突。外观比较把缺省的开关当 false、把缺省/none/automatic 底色当无底色。
  单行单格且全是等宽段落的表格是代码框，只保护。
- **图片文字**：读 `one:Image/one:OCRData/one:OCRText`；合成页面上产生不了 OCR，真实识别结果需要在 OneNote 里人工验收。

## 结构调整

2026-09-28 追加。删空行、缩进、移动、插入和转表格改变段落数量、顺序或嵌套，不能沿用按段落重放的办法，改为按文本框处理：

- **结构草稿**：快照另存一份页面副本 `Layout`，段落用私有命名空间属性 `urn:onenote-code-helper:agent` 记短 ID，新段落没有 objectID。
  结构工具在副本上操作，校验通过才整个换掉；格式草稿仍在各段落的 `Draft` 里，`CreateDraftPage` 把两者合起来。
- **不变量**：每次结构操作后，已排入的代码框和表格转换必须仍能选中且内容不变；已有段落的语义格式不能变
  （`Css.Effective` 会继承上级 OE 的样式，挂到带样式的段落下面会被带成别的格式，这种调整拒绝）。
- **选区边界**：发布结构草稿前，对比选区外原有 OE（包括无文字的图片、表格包装）的存在、祖先链和彼此顺序。
  减少缩进、移动或合并连带改变这些节点时，要求扩大选区并拒绝整个调用，草稿和修订号不变；在未选段落前插入或移动选中段落不算越界。当前页范围不增加此限制。
  整体选中的表格（与 `CollectTables` 同一口径：有文字的格全部选中，含代码框）和图片（本身或所在段落 `selected="all"`）记入 `SelectedObjects`，
  它们的外层段落和格内段落（包括没选中的空行）属于选区，不参与边界比较，可随选中段落缩进、移动或合并。
- **整框提交**：有结构改动的文本框在提交时比较整框指纹（整框 XML，去掉选中状态、修改时间和图片识别文字，加上引用的样式和标记定义）。
  一致时用结构草稿替换整个文本框：套上格式草稿、代码框和表格转换、表格外观，样式和标记编号按内容对应到重新读取的页面，去掉列表的段落照旧重建。
  不一致时结构改动按冲突跳过，框里的格式、代码框、表格改动照常逐项提交，新插入的段落记为冲突。
- **核验**：内容核验照旧比较期望页面和回读页面的拓扑与正文；替换框里每个文字段落再按位置比较语义格式，新表格的单元格另比正文和链接。
  只移动、缩进的段落留在「未指定段落」里，写入前后都要求格式不变。
- **整框撤销**：记下写入前的文本框、引用的定义和写入后的指纹，框里不再另记逐段撤销项。撤销时指纹一致才换回；
  已经不在页面上的对象（删掉的空行、转换前的段落、表格外层）去掉 ID 重建，插入的段落和新表格写回时一并删除。
- **转表格**复用代码框转换：选段规则相同（连续、下级段落在范围内、不带列表和标记），没有结构改动的文本框里按段落补丁提交，撤销时换回原段落。
  单元格复制源段落的 `style` 和 `T`，用 `AgentRichText.Keep` 截取，保留加粗和链接。

### 跨文本框与合并

2026-09-28 追加。`move_blocks` 可以把段落移到另一个文本框，`merge_outlines` 把整个文本框并进另一个文本框：

- **联动组**：跨框的改动记下源文本框 `AgentLayoutChange.From`，`AgentPageSnapshot.LayoutGroups` 按 (OutlineId, From) 并查集分组。
  提交时组内每个文本框的整框指纹都和快照一致才整组替换，有一个不一致就整组按冲突跳过（框里的格式修改照旧逐项补丁提交）；撤销按 `Group` 整组恢复或整组跳过。
- **按实际位置取改动**：`AgentBlock.ContainerId` 等是快照时的位置，整框替换改为按结构草稿里的实际位置取格式草稿、转换和表格外观，整次提交共用一份草稿页面；
  「原来有没有列表」和「对象原来在哪个文本框」按写入前的整页判断。补丁模式仍用快照时的位置，回退时段落就在原处。
- **跨框的对象去掉 ID**：OneNote 把移到另一个文本框的对象一律按新对象建立，`AgentLayout.StripForeign` 事先去掉原来不在这个框里的对象的 ID，
  其中的图片由 `FillImageData` 从 piBinaryData 补上 `one:Data`（只带 CallbackID 会成坏图）。新建段落里的图片写入后按二进制数据核对数量和内容。
- **删框**：合并掉的文本框写成一行空白，OneNote 收到后直接删掉它；期望页面里不算这个框。写入后还留着占位空白时，核验通过后用 `DeletePageContent` 删，
  失败则留着一行空白并在结果里说明（`Leftover`）。
- **撤销重建**：联动组的撤销记录先存下图片数据。被删的文本框按写入前的 Position、Size 和内容新建（没有 objectID），原来就在这个框里、现在还在的对象保留 ID，
  其余去 ID 重建；OneNote 按位置把新文本框排进页面 XML，核对前 `AlignNewOutlines` 把回读页面里的新文本框挪到期望页面的顺序。

本机实测（OneNote `16.0.20326.20158`）：整框回传时，同一文本框里调换顺序、挂到别的段落下、提到上一级的段落都保留 objectID；
省略的段落被删除；无 ID 的段落和表格按新对象建立，`hasHeaderRow`、`bordersVisible`、首行底色照写。
`one:Column` 缺 `width` 会被架构拒绝，但未锁定列的宽度由 OneNote 按内容重算，所以只写占位值。
整框按写入前的样子恢复可行，文本框 ID 不变。
跨文本框移动（两种发送顺序都试过）的段落和下级段落都得到新 ID，只写目标框时源框里的原段落照旧留着；图片带 `one:Data` 跨框写入两次，数据不变，只带 CallbackID 则成坏图。
`one:Outline` 不带 `one:OEChildren` 或 OEChildren 为空都被架构拒绝；写成只剩一行空白时 OneNote 直接删掉这个文本框；无 ID 的新 Outline 按位置排进页面 XML。

## 代码位置

| 文件 | 职责 |
|---|---|
| `Ribbon.xml`、`AddIn.cs` | 按钮、STA 窗口启动、避免重复窗口、关闭时取消任务 |
| `Views/AgentWindow.xaml(.cs)` | 功能（Agent 或文字功能）、模型、思考强度、范围、需求、进度（轮次、思考摘录、执行步骤）、取消、结果、会话撤销；窗口大小不随内容变，可拖动调整 |
| `Services/Agent/AgentRunner.cs` | 模型循环、历史消息、工具分派、调用幂等和预算 |
| `Services/Agent/AgentChatClient.cs` | Chat Completions、HTTP/SSE、工具片段聚合、消息 DTO |
| `Services/Agent/AgentTools.cs` | 工具注册、Schema 与本地校验、草稿发布 |
| `Services/Agent/AgentMarks.cs` | 列表与标记：写法、TagDef 复用与对应、元素顺序、核验投影 |
| `Services/Agent/AgentTables.cs` | 表格外观、表格快照项、指纹和撤销记录；text_to_table 的拆分和表格生成 |
| `Services/Agent/AgentLayout.cs` | 结构草稿的私有 ID、删空行、缩进、移动、插入、合并、整框指纹、跨框去 ID 和图片数据、整框撤销记录 |
| `Services/Agent/AgentFormatting.cs` | 富文本解析与局部样式、等宽判定、CSS 归一化、预设、QuickStyleDef 管理 |
| `Services/Agent/AgentCode.cs` | 代码框和表格转换的范围校验、原段落记录、代码框指纹和撤销还原 |
| `Services/Agent/AgentPageSnapshot.cs` | 范围、短 ID、保护对象、指纹、配置和语义投影 |
| `Services/Agent/AgentCommitter.cs` | 时间戳提交、冲突重建、整框（联动组）替换、删框和重建、核验、撤销 |
| `Services/PageEditCoordinator.cs` | 有界的页面提交锁，与原有编辑器共用 |
| `Services/OneNoteApi.cs` | 显式 xs2013、非强制更新、COM 在途调用和断开协调 |
| `Tests/`、`Tools/agent-test.ps1` | 独立签名测试程序、Fake 页面、脚本模型、HTTP handler 模拟 |
| `Tools/agent-format-probe.ps1` | 显式创建并编辑专用合成页面的真实 COM 探针 |

与原设计的文件布局相比，协议 DTO 留在 `AgentChatClient.cs`，工具注册表留在 `AgentTools.cs`，两个样式编辑器合并为 `AgentFormatting.cs`。测试使用单独的 net48 EXE，没有向生产 `RenderDiagnostics` 增加测试接口。

## 协议、限额与异常

- 单轮请求发送 `tools`、`tool_choice=auto`、`stream=true`，不沿用原 AI 优化的 JSON 输出模式。支持 SSE 以及服务端返回的普通 JSON 消息。
- SSE 按 `tool_calls[index]` 拼接 arguments，保留工具 ID、名称和 `reasoning_content`；收到有效 `finish_reason` 并通过完整校验前不执行任何工具。
- 无结束原因的截断、流内错误、重复工具 ID 携带不同参数等情况停止执行，报错按结束原因说明（`length`、`content_filter`、连接中途结束等）。重复 ID 且参数相同则复用上次结果。
- `length`（输出达到 `MaxTokens`）同样不执行该轮工具，但草稿没动：不是最后一轮时整轮丢掉（半截的思考和工具调用不进历史），追加提醒让模型少思考、分批调用后重试，每个任务最多重试 `AgentRunner.MaxTruncatedRetries`（2）次，再截断就停止并提示调大「最大输出 token」。
- 追加给模型的提醒（截断、没调用工具、收尾）遇到上一条已是 user 消息时并进去，不发连续两条 user 消息。
- 默认最多 24 轮、96 次工具、总时限 600 秒、200,000 个可编辑字符、1000 段；单轮请求体最多 1,000,000 字符（按 DeepSeek 1M token 上下文、约 0.6 token/汉字估算）。限制可在允许范围内配置，上下文较短的模型应调小。
- 请求体和工具结果用 `AgentJson` 序列化：JavaScriptSerializer 转义的 `< > & '`（`<` 等）还原成原字符，避免工具结果作为字符串再序列化时膨胀成 7 个字符；工具结果去掉值为 null 的字段，空数组保留。read_blocks 只对含行内格式（T 里有 `<`）的段落给 `runs`，纯文字段落只有 `text`。历史工具消息和 `reasoning_content` 不裁剪。`Serializer()` 的 `MaxJsonLength` 为 16 MiB，只作序列化保护，预算由 `MaxRequestChars` 控制。
- HTTP 单次超时沿用 AI 配置；无新数据 60 秒终止。SSE 传输上限为 8 MiB（包括每个分片重复的 JSON 包装），本轮真正需要保存的思考、正文和工具参数合计上限为 500,000 字符；单条事件与非流式 JSON 上限 600,000 字符，单工具参数 64,000 字符，JSON 最大嵌套 40 层。超限时不执行不完整工具。
- 系统提示词末尾写明本次的轮数和工具次数上限，并要求互不依赖的读取和修改放在同一轮调用。轮数快用完时收尾：剩 2 轮时追加提醒；最后一轮追加当前 `draft_revision`，请求只带 `finish_edit` 的定义，调用其他工具按工具错误返回。收尾阶段提交成功时，结果消息注明是轮数用完提前提交的；最后一轮只回文字时按 `NoChange` 返回，不再报错。
- 上下文预算、工具次数、总时限达到上限，或在提交前取消，丢弃草稿；最后一轮仍没有提交也丢弃草稿并报错。COM 提交已经开始后，即使取消也要先尝试核验结果。
- 不把供应商错误响应正文、工具参数或笔记正文写进新增日志；日志只记录状态、次数和异常类型。每轮请求另记一行：模型、思考强度、请求字数、耗时、结束原因、token 用量（输入/输出/思考，需 `StreamUsage`）以及思考、回复、工具参数的字数。

旧 `ai-settings.xml` 没有 `Agent` 节点时使用默认配置。不会改写用户现有配置。模型地址、Key、模型列表和思考强度沿用现有设置；开关和完整示例见 README。

## 写回保护和回存适配

- 原段落指纹包括正文和富文本、直接属性、祖先属性及位置、同类兄弟序号、引用的 QuickStyleDef；排除选中状态及更新时间。用户改文字、仅改格式、移动、删除或修改相关祖先样式都会触发冲突。
- 先缓存所有目标的提交前指纹，再修改临时页面，避免同批父段落修改造成子段落的假冲突。
- 新的 QuickStyleDef 不覆盖原定义。重读后如果样式索引被占用，会重新分配；原生标题使用 h1/h2，正文与引用使用 p。
- `0x80042010` 最多重读重建两次，不重发过期 XML。其他 COM 异常不盲目重试，先回读确定实际结果。未知结果显示 `CommitOutcomeUnknown`。
- 原 `PageEditor.Submit` 同时移除了 `DateTime.MinValue` 无校验重试，以免旧功能在与 Agent 并行使用时覆盖新内容；冲突时提示重新执行。
- 回读比较目标的实际文字格式和段落格式，并检查正文、链接、段落及表格拓扑、代码与其他未指定段落的格式、图片引用等不变量。
- OneNote 会重写 HTML：包括无引号属性、本地化字体名称、span/T 合并以及冗余样式。使用按字符的有效样式比较，不依赖 HTML 字符串完全相同。链接默认颜色与段落基础色分别处理。
- OneNote 回存表格会重新生成纯表格外层 OE 的 ID。只允许这个包装节点用内部 Table ID 对应；表格本体、行、单元格和正文段落的 ID 与顺序仍必须保持。
- 未锁定列宽由 OneNote 随字体自动重算；不把这种变化当成结构损坏，也不替用户锁列。锁定列宽、列索引、表格边框和单元格底色仍严格检查。
- COM 调用串行保护。断开后拒绝新调用，已有调用结束后再释放捕获的 COM 引用，不在网络等待期间持有 COM 锁。

## 验证结果

在 Windows、OneNote `16.0.20326.20158` 上验证。生产构建输出到 `obj\agent-build\`，没有覆盖工作区原本已修改的 `bin\Release\net48` DLL/PDB。

| 验证 | 结果 |
|---|---|
| Release 构建 | 0 警告、0 错误 |
| Agent 离线回归 | 50/50 通过，覆盖格式保留、Unicode、工具约束、协议分片、多轮、冲突、表格回存、取消、撤销和模拟 HTTP |
| 原 AI 合并回归（不加 Live） | 69/69 通过 |
| 原高亮选区回归 | 11/11 通过 |
| 原语言识别回归 | 103/106；python/dict-config、python/print、python/print-only 与修改前 DLL 相同失败，未修改语言识别实现 |
| 真实 OneNote COM 合成页面 | 修改 4 段、撤销恢复 4 段，均 Verified；0 冲突、0 未验证，1 个代码段受保护 |
| WPF 窗口渲染 | 默认与最小宽度检查；修正最小高度使结果区可见 |

真实探针已覆盖一级/二级原生标题、字号、中文字体、居中/右对齐、段间距、局部取消加粗/改色/斜体/下划线、链接、emoji、嵌套段落、普通表格文字、自动/锁定列宽、图片与代码保留，以及格式撤销。
生成的 `.one` 分区和 before/expected/after/after-undo XML 保留在指定的 `obj\agent-probe\` 目录（Git 忽略），未编辑已有用户页面。

代码框转换（2026-09-26 追加）：Agent 离线回归 61/61（新增代码分类、范围校验、冲突、与格式合并提交、撤销和 Runner 全流程），高亮选区 11/11，AI 合并 69/69，语言识别 106/106。
探针已加入一段普通字体代码和一段整段 Consolas 代码的转换与撤销，但尚未在真实 OneNote 上重新运行；OneNote 回存代码行时的空白规范化仍需用探针确认。

列表、标记和表格样式（2026-09-28 追加，OneNote `16.0.20326.20158`）：Agent 离线回归 83/83（FakePage 模拟了「省略 List 不去掉列表」），
高亮选区 11/11，AI 合并 81/81，语言识别 106/106。真实 OneNote 探针新增项目符号、编号、重要标记、已完成待办、表格标题行和表头底色，
执行和撤销都是 Verified；撤销后列表和标记全部去掉，重建的父段落下级段落保持原 ID。图片文字未在真实页面上验收。

结构调整（2026-09-28 追加，OneNote `16.0.20326.20158`）：Agent 离线回归 100/100，高亮选区 11/11，AI 合并 81/81，语言识别 106/106。
真实 OneNote 探针新增第二个文本框，删空行 2 行、移动 1 段、缩进 2 段、插入 2 段、pipe 文字转表格（单元格含加粗和链接），再加一处段落样式，
一起整框提交，执行和撤销都是 Verified；撤销后文本框的文字顺序与执行前一致，移动和缩进过的段落保持原 ID。第一个文本框仍按段落逐项提交和撤销。

跨文本框与合并（2026-09-28 追加）：Agent 离线回归 109/109，高亮选区 11/11，AI 合并 81/81，语言识别 106/106。
真实 OneNote 探针新增第三个文本框（带下级段落的文字、2×2 表格、图片），先跨框移动一段，再把其余内容合并进第二个文本框：
执行 Verified，第三个文本框在同一次写入里被删掉，两张图片按二进制数据完整；撤销 Verified，第三个文本框按原内容重建（新 ID），第二个文本框恢复原样。

重现本次构建和离线测试：

```powershell
dotnet build OneNoteCodeHelper.sln -c Release -p:OutputPath=obj\agent-build\
powershell -ExecutionPolicy Bypass -File Tools\agent-test.ps1 -DllPath "$PWD\obj\agent-build\OneNoteCodeHelper.dll"
```

仅在需要真实 OneNote 验证时运行：

```powershell
powershell -ExecutionPolicy Bypass -File Tools\agent-format-probe.ps1 -OutputDirectory "$PWD\obj\agent-probe"
```

离屏窗口检查（不启动 OneNote）：

```powershell
.\Tests\bin\Release\net48\OneNoteCodeHelper.AgentTests.exe --render-ui "$PWD\obj\agent-ui"
```

### Agent 审查修复（2026-09-30）

修复转换撤销覆盖后续编辑、Unicode 代理项序列化、结构操作越过选区、继承格式修改误判无需修改，以及父段原生标题影响子段这五类问题。
8 个复现场景已迁入 `Tests/Program.cs`，加上引用样式、位置、自动列宽、组合结构操作、嵌套标题和失败原子性等边界，共新增 19 项正式测试。

本轮 Release 构建为 0 警告、0 错误；Agent 141/141、AI 合并 86/86、高亮选区 11/11、语言识别 130/130 全部通过，共 368 项（原有 349 项 + 新增 19 项）。并发编辑保护、整框撤销和跨框整组撤销回归均通过。
生产构建输出到 `%TEMP%\OneNoteAgentFix-20260929\app\`，四组脚本均通过 `-DllPath` 指向该 DLL，没有改动仓库跟踪的 DLL/PDB。
本轮仅使用离线回归和 FakePage 回存模拟，没有调用真实 AI、执行安装或注册脚本、修改真实笔记。新增修复在 Office 中的实际回存与撤销尚未验证，尤其是父段仅设置外观时的继承行为和转换结果指纹的回存稳定性；上文历史探针记录不代表这些新修复已经通过真实 Office 验收。

后续修正（同日）：上面的选区边界把整体选中的表格、图片也当成选区外对象，选区范围里合并含这些对象的文本框、缩进挂着它们的段落都被拒绝；
转换撤销指纹记录了同级序号，在转换结果前后增删段落或别处的文本框都会误判冲突。两处已修正，新增 2 项测试（选中表格和图片的合并、缩进及反例；代码框和转表格在前后增删段落、文本框后照常撤销），
两项在修正前均失败。Agent 143/143、AI 合并 86/86、高亮选区 11/11、语言识别 130/130 全部通过；同样构建到临时目录，没有改动仓库跟踪的 DLL/PDB，也未在真实 Office 中验证。

再次修正（同日）：父段仅设置外观时保留原有 QuickStyleDef，其中的加粗、斜体、下划线仍会生效（加粗的标题改成正文仍是粗体，斜体的引用改成标题仍是斜体），结果却报告已核验。
现在回退时由 `ParagraphStyles.PinEmphasis` 把目标定义的这三项写到文字上（文字上已有的显式值保留），并核对父段每个字的有效样式与原生样式版本一致，不一致时整次调用拒绝、没有副作用。
另外，页面上有重复 objectID 的段落时，`set_paragraph_style` 的继承检查和提交时的「未指定段落不变」核对都用 `SingleOrDefault` 查 ID，抛出的非 `AiException` 会让整次 Agent 中断；
后者在此前的版本里就存在。继承检查改为按两份草稿页的位置对应，提交核对改为比较同一 ID 全部段落按页面顺序的格式。新增 2 项测试（三种外观回退与无子段原生结果逐字比对；重复 ID 在目标下和不在目标下时设置样式、提交、撤销），
两项在修正前均失败。Agent 145/145、AI 合并 86/86、高亮选区 11/11、语言识别 130/130 全部通过；构建到临时目录，没有改动仓库跟踪的 DLL/PDB，也未在真实 Office 中验证。

第三次修正（同日）：`PinEmphasis` 只看文字 style，父段自身 style 上的加粗、斜体、下划线被定义的默认值盖掉，与原生版本比对不一致，整次调用被拒绝；
原本可用的「先设父段、再设子段」流程因此失败。现在段落 style 已有的项也不补，与原生样式「定义 → 段落 style → 文字 style」的优先级一致。
另外，转换撤销指纹直接包含 `quickStyleIndex` 和定义的 `index`，OneNote 给样式定义重新编号后内容不变也判为冲突。现在样式和标记引用换成所引用定义的内容再计算，找不到定义时保留原编号；定义内容变了仍算冲突。
新增 2 项测试（父段 style 带加粗、斜体、下划线时回退与原生外观一致并能先父后子提交、撤销；代码框和转表格在样式重新编号后照常撤销、改了定义内容仍冲突），两项在修正前均失败。
Agent 147/147、AI 合并 86/86、高亮选区 11/11、语言识别 130/130 全部通过；构建到临时目录，没有改动仓库跟踪的 DLL/PDB，也未在真实 Office 中验证。

## 仍需用户环境验收的部分

本次没有调用真实模型接口，也没有安装、注册或重启当前插件。离线协议测试验证了实际 HTTP 客户端的请求体和 SSE 解析，但不等同于验证用户的供应商、Key 和所选模型。

针对「Agent 响应超过大小限制」的现场报告，修正了流式响应把每个 SSE 分片重复的 JSON 包装计入 600,000 字符上限的问题。现在传输字数与有效输出分别计数，并用 `StringBuilder` 累积大量小片段；构造超过旧限制的流式数据已通过回归验证。DeepSeek 的[思考模式工具调用文档](https://api-docs.deepseek.com/guides/thinking_mode/)要求后续请求回传完整 `reasoning_content`，所以不会截断模型历史。
更新安装后，需要用本机配置运行一次示例需求，检查模型的标题识别和最终排版是否符合偏好；也应在实际加载项中验收开关窗、切页及关闭 OneNote 的生命周期。

墨迹、附件、媒体、未知 HTML、图片位置和任意布局调整不在本版能力内；遇到不支持的对象会保护相关段落或整个容器。对其他 Office 构建的回存兼容性应重新运行探针。
