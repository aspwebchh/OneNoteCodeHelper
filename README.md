# OneNote 代码高亮

OneNote 桌面版的 COM 外接程序，把笔记里的代码渲染成带底色的高亮代码框。
目前支持 **Java**、**C#**、**C/C++**、**JavaScript**、**TypeScript**、**Python**、**Go**、**Kotlin**、**Rust**、**PHP**、
**SQL**、**Lua**、**PowerShell**、**Bat**、**Bash**、**XML**、**HTML**、**CSS**、**JSON**、**YAML**，
以及不着色的 **纯文本**（只要等宽字体和代码框，适合放日志、命令输出）。

功能区「开始」选项卡上会多出一个「代码高亮」组：

| 控件 | 作用 |
|---|---|
| 高亮选中 | 选中页面上已有的代码文字，原地替换成高亮代码框 |
| 插入代码 | 打开窗口粘贴代码，预览确认后插入到当前页 |
| 语言 | 自动识别，或手动选上面列的任意一种。纯文本只能手动选，自动识别不会选它；JSX 按 JavaScript 识别，TSX 按 TypeScript 识别 |
| 深色主题 | 在浅色（GitHub）与深色（GitHub Dimmed）之间切换 |
| 显示边框 | 代码框默认只有底色、不带边框（OneNote 的表格边框颜色不可控）；按下后四周加一圈边框 |
| 字体（插入窗口内） | 默认 Consolas。代码里有中文时改选「NSimSun」新宋体，中英文才能对齐 |
| 诊断日志 | 打开日志文件 |

旁边还有一个「AI 助手」组，用大模型处理笔记文字和格式：

| 控件 | 作用 |
|---|---|
| Agent | 打开 AI 助手窗口，在窗口里选功能、模型、思考强度和范围（当前页 / 选中段落），处理中可以取消 |

Agent 窗口右上角的「AI 配置」打开配置窗口，分「接口」「模型」「文字功能」「Agent」四页设置，保存后写回配置文件（见下面「AI 配置文件」）。
关掉配置窗口回到 Agent 窗口，下拉随即刷新，下次执行使用新配置。处理中这个按钮不可用。

Agent 窗口里的三个下拉，选择会记住，下次打开沿用：

| 下拉 | 作用 |
|---|---|
| 功能 | 第一项「自定义排版（Agent）」按输入的需求整理格式，需求里提到时也修正错别字，支持结果核验和本次撤销（见下面「Agent 页面排版」）；后面是文字功能，默认有「智能校正」（同时校对和排版）「错别字修复」「排版优化」三项，选项和提示词都来自配置文件，可以自己加。初次使用选中「自定义排版（Agent）」 |
| 模型 | 直接显示发给接口的模型名，默认 `deepseek-v4-flash` / `deepseek-v4-pro`，来自配置文件 |
| 思考 | 思考强度 `none` / `low` / `medium` / `high` / `max`，参数和 opencode 配置里 deepseek 的 variants 一致：`none` 传 `"thinking":{"type":"disabled"}`；其余传 `"thinking":{"type":"enabled"}` 加 `"reasoning_effort":"<变体名>"` |

「自定义排版（Agent）」的默认需求在「AI 配置 → Agent」页顶部修改，支持 1–8000 字。打开窗口时自动填入需求框，执行前可以修改或替换；临时修改只在当前窗口保留，不会保存为默认需求。切换文字功能再切回 Agent，输入仍保留。
关闭并重新打开 Agent 窗口，需求框会填入最新保存的默认需求。配置页的「恢复内置默认」只恢复模板编辑框，点击「保存」后才生效。
保存配置后，需求框仍是旧模板时自动更新；已经修改或清空时保留本次输入。执行只使用需求框中的文字，模板不会重复追加。

窗口配色和 OneNote 一致（浅灰底、白卡片，OneNote 紫用在主按钮和选中状态上）。窗口大小不随内容变：处理中的状态、思考摘录、执行步骤和结果都在下方卡片里显示，内容多了在卡片里滚动，窗口不会随内容忽高忽低；可以拖动窗口边缘调整大小，拉高后下方卡片跟着变大。
处理 Agent 时思考框和执行步骤各占下方卡片的一块（约 3:2），文字功能没有步骤，思考框占满。

### 文字功能（智能校正等）

在 Agent 窗口的「功能」里选一个文字功能，需求框的位置换成这个功能的提示词（只读，在「AI 配置」的「文字功能」页改），点「执行」后**直接写回**。
范围选「当前页」处理整页（含标题）；打开窗口前选中了文字时可以选「选中段落」，只处理选中文字所在的段落。

- 只改段内文字，不合并、不拆分段落。改动按字符合并回原段落，**加粗、颜色、链接、列表和缩进都保留**；
  改错字时新字沿用被替换字的格式，排版时补的空格不会跟进加粗或链接。
- 代码框（段落字体是 Consolas、新宋体等等宽字体）不会交给 AI。
- 「智能校正」「排版优化」顺带删掉多余的空行：连续好几行空行只留一行，文本框（表格单元格）开头、结尾的空行删掉；
  Shift+回车打出的段内空行同样处理。删空行就是删段落，AI 按约定不能动段落，所以这一步由插件自己做，
  不经过 AI。带项目符号、待办标记的空行和代码框里的空行不算。选中了文字时只删选区里的空行。
- AI 返回什么就写回什么，不按改动大小把关。
- AI 处理期间你还可以继续编辑。写回前会重新读页面，处理期间被你改过的段落直接跳过，不会覆盖。
- 做完后窗口里显示 AI 自己写的改动清单（每段附的说明汇总去重，只算真正写回了的段落；插件删空行不在里面）。
  文字功能没有「撤销本次」。处理中点「取消」或关窗等于取消，已经开始写回时会等写完再显示结果。
- 请求走流式，窗口实时显示 AI 当前在做什么（思考中 / 输出中 / 已返回几段修改和最新一条改动说明）、
  思考摘录和已用时间。摘录只挑思考里最新的几句核心内容：只显示说完的中文整句，去掉 Markdown 符号、代码和 JSON 片段、
  语气词和整句英文或中英混杂的句子，最多约 16 行，说完新的一句才换；多了自动滚到最下面，往上翻时不会被拉回。
  分了几批并发时，思考摘录只跟一批，不来回跳。Agent 每一轮的摘录单独一段，前几轮的颜色淡一些，留在框里不清。
  思考原文只在窗口里显示，不写日志。
  连续 60 秒一点数据都没收到（接口或网络卡住）会直接报错，不会干等到 `TimeoutSeconds`。

### Agent 页面排版

点击「Agent」，先在窗口右上角「AI 配置」的「接口」页填写接口地址和 Key，然后在「功能」里选「自定义排版（Agent）」，「模型」选一个支持 Chat Completions `tools` 的模型。
窗口固定打开时的页面；之后切换 OneNote 页面，Agent 仍处理窗口里显示的目标页。
输入需求并点击「执行」，例如：

> 将该页面上的内容排版下，要美观。统一正文格式，突出标题，修正错别字。

也可以输入「一级标题居中，正文统一 11 磅，关键结论加粗」。模型通过一组受限工具读取段落、调整草稿、提交修改，
界面显示执行进度，最终结果以插件回读 OneNote 后核验的数据为准。
处理中显示当前轮次、模型正在准备的工具，以及思考（或回复）里最新的几句核心内容（规则同文字功能，另外把工具名、
段落 ID 换成中文说法，如「用「设置段落样式」把第 3 段设为二级标题」）；「执行步骤」逐条列出调用过的工具和要点
（例如「设置段落样式 · 二级标题 · 3 段」，失败的附上原因），做完后也保留。思考原文只在窗口里显示，不写日志。

- 范围默认「当前页」，包括原标题。打开窗口前选中文字，可以切换到「选中段落」；这个范围包含选中文字所在的**完整段落**，不只选中的几个字。选区在打开窗口时固定，重新选择需要重开窗口。
- 支持页面标题、一级/二级标题、正文、引用预设；字体、字号、四种主题文字颜色、左/中/右对齐、段前后间距，以及局部加粗、斜体、下划线和文字颜色。
  给父段设置标题等预设时，如果原生样式会连带改变未指定的下级段落，就只设置父段自身文字的目标外观（字体、字号、颜色和加粗、斜体、下划线与原生标题一致），保留原有标题层级和子段格式；执行步骤和已核验结果会说明「仅设置外观」。不影响子段时仍使用原生标题。
- 需求里明确提到时，可以把段落设为项目符号或编号列表、取消列表，加待办（可勾选完成）、重要、问题标记，或去掉这三种标记（其他自定义标记保留）。
  页面标题和代码段落不能加列表或标记。
- 整体排版时可以统一表格的边框、标题行和首行底色（浅蓝、浅灰、浅绿、浅黄或无底色），也可以去掉全部单元格的底色，不改单元格文字和行列。已有代码框不算表格。
- 需求里要求清除格式（比如「清除所有格式，包括代码外层的框」）时，Agent 用专门的工具一次处理整页或选中的段落，文字一字不改：
  去掉加粗、斜体、下划线、删除线、上下标、颜色、高亮、字体和字号，段落改成正文样式（字体取配置的默认字体，左对齐）；默认同时去掉列表符号、全部标记（含自定义标记）和链接（链接文字保留），
  需求里说要保留哪一项就保留，拆开代码框时也遵守保留链接的要求。行内代码、整段等宽的代码和空行也一起改成正文；页面标题只去掉文字格式，保留标题样式；emoji、符号字体保留，免得显示成方框。
  只清理父段时，在父段自身的文字上清除格式，保留未处理子段的外观；只清理子段时，也会去掉从父段或文本框继承的格式。关闭原生标题功能时，正文外观仍可清理。
  已有代码框会拆成正文段落：去掉代码框和语法颜色，每个代码行一段，文字、行首缩进和空行不变；普通表格恢复默认外观（显示边框、无标题行、无底色）。
  列表、标记的开关关闭时这两项照旧保留。结果里写「去掉链接 N 处」「拆开代码框 N 个」，链接数包含已核验拆框时删除的链接，撤销时链接、列表、标记和代码框一起还原。
  清除继承格式与拆框的真实 OneNote 回存、撤销仍待验收，目前这些修复只通过离线回归。
- OneNote 已识别出文字的图片（右键能「复制图片中的文本」的），Agent 可以读取图片文字来理解页面，但不修改图片。
- 需求里明确提到时，还可以调整段落结构：
  - 把段落连同下级段落移到别处，可以移到另一个文本框；表格单元格里的段落只能在同一个单元格里移动；
  - 把整个文本框（文字、表格、图片、空行）并进另一个文本框的指定位置，再删掉空了的源文本框；
  - 调整缩进层级，上下顺序不变；
  - 把代码框拆成正文段落（见上面的清除格式）：只拆外层段落上没有列表、标记和下级段落、格里只有平铺文字段落的代码框；拆出来的段落本次执行内不再移动或修改；
  - 在指定位置插入摘要、目录或小标题：只写纯文字、套用预设样式，每次最多 20 段，每次执行最多 50 段、5000 字；
    也可以插入空行（正文外观、零段间距），比如把几个文本框合并成一个、要求各部分之间空一行时，合并后在每个交界处补一行，已有空行的交界不再加。空行计入段数，结果里写「补空行 N 行」；
  - 把用制表符、`|` 或空格分隔的段落转成表格：每行文字一行（段内 Shift+Enter 换行的也各成一行），去掉 Markdown 分隔行，单元格保留加粗和链接，列宽由 OneNote 按内容自动计算。
    一条记录分成几行（如名称一行、地址一行，中间隔着空行）时，每几行合成表格的一行，名称行尾的冒号去掉。
    分隔方式不统一、键值对（如「IP：10.0.0.1，端口：80」）、有缺项等不规整的文本，由 Agent 逐格给出切分结果，插件按顺序在原文里逐字核验：
    原文的每个字都要放进单元格，只有空白、制表符、`|`、冒号、逗号、分号、顿号、等号这类分隔符，以及和该列表头相同的标签（挪进了表头）可以省略，
    改写、调换、重复或漏掉文字都会被拒绝。
    各行列数不一致时按最多的列数建表，缺的单元格留空；单元格里本身有空格的内容建议用制表符或 `|`。需求里要求加表头时，可以在首行前新增一行列名。
    标题和数据在同一段里（用 Shift+Enter 换行）时，标题会成为表格的第一行、其余单元格留空；想让标题留在表格外，先用回车把它分成单独的段落。
- 整理、美化排版或需求里提到删空行时，会删掉多余的空行：连续空行只留一行，文本框首尾的空行删掉。带列表、标记的空段落和代码里的空行不删。
  本次新插入的空行也能在同一次任务中清理；结果只统计实际保留的「补空行」和原页已有空行的删除。新空行即使随后删掉，仍占本次任务的插入配额。
- 需求只要求「代码块和文字之间保持一个空行，多删少补」时，用专用工具处理同一文本框内代码框与文字的交界，普通表格的单元格各自计算。
  空段落和文字首尾的 Shift+Enter 空行合计保留一行；等宽字体的框外空段落也参与计算，没有时补一行正文空段落。
  文字之间、代码框之间、文本框首尾和代码内部的空行保持原状，不用段间距模拟空行，不调整独立文本框的位置。
  删除受 `EnableBlankLineRemoval` 控制，补入受 `EnableInsert` 和现有插入配额控制；选区不完整、内容受保护或对应能力关闭时，整处跳过并说明原因。
  调整与代码转换、格式修改一起整框核验和撤销；处理中用户编辑过的文本框整体跳过空行调整。
- 页面标题、代码框和含墨迹等对象的文本框不调整结构；挂到带样式的段落或文本框下面会继承它的样式时，这次缩进、移动或合并不做。
  在「选中段落」范围内，结构调整还必须保留选区外段落和对象的存在、祖先关系及彼此顺序。例如减少缩进会把未选中的后续兄弟段落挂到自己下面时，整次工具调用会被拒绝；需要扩大选区、重开窗口后再执行。已全部选中的表格（有文字的格都选中）和图片属于选区，可以跟着选中段落一起缩进、移动或合并。当前页范围沿用上述结构规则。
- 移到另一个文本框的段落、表格和图片，OneNote 都按新对象建立：内容和格式不变，但会得到新的 objectID，指向它们的段落链接会失效。
- 需求里提到修正错别字时，Agent 用专门的工具逐处修正错别字、同音字、形近字和明显的标点误用：每处只换出错的几个字（原文和改后各不超过 30 字，不能跨换行），
  没改的字保留原有格式和链接，代码段落不改。结果里列出每处修正，例如「按装」→「安装」；这些内容只在窗口里显示，不写日志。
- 需求里要求去掉 Markdown 格式（比如把 AI 回复粘进来后整理）时，Agent 用专门的工具一次处理整页或选中的段落，只删 Markdown 标记字符，其余文字的格式和链接不变：
  行首的 `#` 标题、`>` 引用、`-` `*` `+` 和 `1.` 列表符号（含 `[ ]` 待办框），行内 `` `代码` ``、`**粗体**`、`*斜体*`、`~~删除线~~` 两边的符号。
  整段只有 ```` ``` ```` 围栏或 `---` 分隔线的段落直接删掉（和删空行一样受 `EnableBlankLineRemoval` 控制，关掉时只清空文字，文本框里至少留一段）。
  `markdown`/`md` 围栏里的内容照常处理；其他围栏（包括不写语言的）里是代码，不动，可以接着转成代码框。去掉的标记会告诉 Agent，
  它可以据此把原来的 `#` 设为标题、`-` / `1.` 设为列表、`[ ]` 设为待办。`snake_case`、`2 * 3`、`a*b*c` 这类不当强调；链接、图片和表格不处理（表格可以转成表格）。
  围栏保护按执行开始时的完整上下文保留，只选中围栏内部、分批或重复清理也不会改代码；无法解析上下文的文本框或单元格跳过清理。
  支持 OneNote 的不换行空格及嵌套粗体、斜体；清理后接着转表格会保留当前草稿里的文字修正、行内格式和链接，撤销时一起恢复原文。
  恢复编号列表时保留各组的原始起点；正文或标题隔开的列表分别编号，不会把后面的 `1.` 接成前一组的 `9.`。
  结果里只写「去除 Markdown 符号 N 处」，不逐条列出；撤销时一起还原。
- 除上面的错别字修正、去除 Markdown 标记、清除格式去掉的链接和结构调整外，保留文字、段落数量、顺序、嵌套关系、链接和表格结构；列表、标记和表格外观只按上面的规则修改。不合并、拆分段落，不润色改写；图片不单独移动，只随所在的段落或文本框一起移动。
- 取消列表时（包括撤销新加的列表），OneNote 不接受在原段落上去掉列表，插件把这一段重建为新段落：文字、格式和下级段落都不变，
  但这一段会得到新的 objectID，指向这一段的段落链接会失效。
- 排版时遇到还没高亮的代码，Agent 调用「高亮选中」同一套逻辑，把连续的代码段落（含中间空行）换成高亮代码框。
  主题、字体、字号、Tab 宽度、边框取打开窗口时功能区的设置；语言由模型指定，或自动识别。当普通正文输入的代码、整段都是等宽字体但还没放进代码框的代码（比如从 IDE 粘贴来的）都会转换。
  代码段落必须在同一个文本框或表格单元格里、前后连续，第一行不能比后面的行缩进更深，带项目符号或待办标记的段落不转换。
- 表格中可以调整普通文字的格式，锁定列宽保持不变；未锁定列仍由 OneNote 根据字体自动计算宽度。
- 已有的代码框、行内等宽代码所在段落、空段落和无法解析的 HTML 会受保护，代码段落也不会被设置正文样式（要求清除格式时除外）。含图片的文本框可以整理文字；含墨迹、附件、媒体或未知对象的整个文本框跳过。
- 修改先保存在内存草稿中，模型调用 `finish_edit` 后统一写回。写回前重新读取页面；处理期间用户改动的目标段落会跳过，其他段落继续提交。时间戳持续冲突时停止，不强制覆盖。
  所有 Agent 任务提交前都须完整读取所选范围内的正文和待高亮代码，页面概况按每页 100 段翻页，正文每批最多读取 100 段；局部需求也会读取整个范围，但只修改指定目标。
  正常提交时仍有未读段落会被拒绝，草稿保留并返回下一批段落 ID，供 Agent 继续处理；空行、受保护内容、已删除或已转换的段落以及本次新插入的段落不计入未读数量。
  有结构调整的文本框整框写回：处理期间这个文本框里有任何改动，结构调整整体跳过，框里的格式修改仍逐段提交。
  跨文本框移动、合并涉及的几个文本框一起写回、一起跳过，不会出现一边写进、一边没删的重复或丢失。
- 「取消」或关闭窗口会丢弃未提交草稿。如果 OneNote 已开始写入，则先回读确认实际结果；无法确认时会明确提示检查页面，不把模型的“完成”当作成功。
- 「撤销本次」只在当前窗口中保留最近一次执行的已核验修改，修正过的错别字、列表、标记和表格外观也一起还原。撤销会跳过后来又编辑过的段落。关闭窗口或启动下一次执行后不保留上一次撤销记录。
  代码框和转换得到的表格也一起撤销：之后没被改过时换回原来的段落（原段落会得到新的 objectID）。后来修改文字、格式、链接、列表、标记、图片、结构或相关继承样式，或者移到其他文本框、上级段落下、拖动所在文本框时，整处跳过，保留后续编辑；选中状态、修改时间、表格包装段落 ID、未锁定列的自动宽度和样式定义的重新编号不影响撤销，在它前后增删段落、页面上其他文本框的增删也不影响。
  结构调整按文本框整框撤销：文本框之后没被改过时，恢复原来的段落、顺序和层级，框里的格式修改也一起还原；改过的文本框整框跳过。
  拆开的代码框同样整框撤销，换回原来的代码框（底色、语法颜色都在），代码框会得到新的 objectID。
  移动和缩进过的段落保持原 objectID；删掉的空行、转成表格的段落恢复时会得到新的 objectID。
  跨文本框移动、合并涉及的几个文本框一起撤销，其中一个之后被改过就整组跳过；合并删掉的文本框按原来的位置和内容重建，得到新的 ID。
- Agent 使用窗口里选的模型、思考强度，不使用文字功能的提示词。执行文字功能也会让上一次 Agent 执行的撤销记录作废。

工具协议和实现位置见 [Agent 实施说明](docs/agent-implementation.md)，完整设计见 [技术设计](docs/agent-design.md)。

### AI 配置文件

`%APPDATA%\OneNoteCodeHelper\ai-settings.xml`。平时在 Agent 窗口右上角的「AI 配置」里改，配置窗口的每一项都对应这个文件里的一个字段（Agent 页的开关，鼠标停上去能看到对应的节点名）。
文件在窗口第一次保存时生成，里面每一项都有注释。想直接改 XML 时，点窗口左下角的「用文本编辑器打开」：
用系统默认的程序（.xml 关联的编辑器，没有就用记事本）打开，同时关掉窗口，免得两边同时改。

| 字段 | 说明 |
|---|---|
| `ApiUrl` | OpenAI 兼容接口的基础地址；插件会在后面加 `/chat/completions`，已写全该路径的地址也可直接使用 |
| `ApiKey` | 接口的 Key。**默认是空的**，不填在 Agent 窗口点「执行」会提示 |
| `TimeoutSeconds` | 单次请求从发出到收完的总时限，默认 300 秒。另有固定的 60 秒「无数据」判定，不在这里配 |
| `MaxTokens` | 单次请求最多输出多少 token（含思考过程），默认 16384，0 表示用接口默认值 |
| `Models/Model` | Agent 窗口「模型」下拉的选项，`id` 是接口的模型名，下拉里直接显示它 |
| `Functions/Function` | Agent 窗口「功能」下拉里的文字功能，`name` 显示名（不要和「自定义排版（Agent）」重名，重名的会被跳过）、`Prompt` 提示词。`removeExtraBlankLines="true"` 表示顺带删多余的空行；不写时和同名的内置功能一致（「智能校正」「排版优化」默认开，旧配置里的「错别字 + 排版」也继续默认开），写 `false` 关掉 |
| `Agent` | 可选的 Agent 配置；旧文件没有此节点也能使用默认值 |
| `Agent/DefaultRequest` | Agent 打开时预填的默认需求，1–8000 字；缺失、空白或超长时使用内置默认。工具协议、提交核验规则仍由插件提供 |

提示词只需写清楚要做什么。输入输出的 JSON 格式、只返回改动的段落、每段附一份改动说明、不许合并拆分段落、
代码网址保持原样这些约定由插件自动接在后面（见 `AiOptimizer.Protocol`），改提示词不会把格式弄坏。

窗口保存时在原文件上就地改，不整个重写：

- 手写的注释和窗口里没有的节点都保留。`Models`、`Functions` 两节按窗口里的列表重建，这两节里面的注释不保留；
  `removeExtraBlankLines` 一律写明 `true` 或 `false`。
- 原来没有 `Agent` 节点时，Agent 页的各项都是默认值就不加；改了其中一项才加上整个节点。
- 先写临时文件再替换，Agent 窗口这时去读也不会读到写了一半的文件。
- XML 写坏了读不了时，窗口顶部会提示并显示默认值；保存时先把原文件备份成 `ai-settings.xml.bak`，再重新写一份。
- 窗口开着时文件在别处被改过：窗口里没有改动就切回来时自动重新读；有改动的话，保存前先确认是否覆盖。
- 保存前会校验：接口地址要是 http(s) 地址，数值要在下面列的范围内，模型和功能至少各一个、名称不能重复，提示词不能为空。

这份文件和 `settings.xml` 分开放，是因为 `settings.xml` 在每次切功能区选项时都会被整体重写，
手改的 Key、提示词放在那里会被覆盖。插件自己只在文件不存在时写默认值，其余时候只有在窗口里点保存才会改它。
所以以后新增的内置功能（比如「智能校正」）不会自动进旧的配置文件：在「文字功能」页点「补回内置功能」，
或者自己在 `Functions` 里加一项。

Agent 可在 `AiConfig` 根节点内增加下列配置，也就是「AI 配置」窗口 Agent 页里的各项。改完保存后，回到 Agent 窗口点「执行」时生效：

```xml
<Agent>
  <DefaultRequest>将该页面上的内容排版下，要美观。统一正文格式，突出标题，修正错别字。</DefaultRequest>
  <MaxTurns>24</MaxTurns>
  <MaxToolCalls>96</MaxToolCalls>
  <TimeoutSeconds>600</TimeoutSeconds>
  <MaxPageChars>200000</MaxPageChars>
  <MaxRequestChars>1000000</MaxRequestChars>
  <FontFamily>Microsoft YaHei</FontFamily>
  <SendThinking>true</SendThinking>
  <ReplayReasoning>true</ReplayReasoning>
  <StreamUsage>true</StreamUsage>
  <EnableNativeHeadings>true</EnableNativeHeadings>
  <EnableParagraphSpacing>true</EnableParagraphSpacing>
  <EnableMixedOutlines>true</EnableMixedOutlines>
  <EnableCodeHighlight>true</EnableCodeHighlight>
  <EnableLists>true</EnableLists>
  <EnableTags>true</EnableTags>
  <EnableTableStyles>true</EnableTableStyles>
  <EnableMarkdownCleanup>true</EnableMarkdownCleanup>
  <EnableClearFormat>true</EnableClearFormat>
  <EnableBlankLineRemoval>true</EnableBlankLineRemoval>
  <EnableIndent>true</EnableIndent>
  <EnableMoves>true</EnableMoves>
  <EnableInsert>true</EnableInsert>
  <EnableTextTables>true</EnableTextTables>
  <EnableCodeUnwrap>true</EnableCodeUnwrap>
</Agent>
```

`EnableCodeHighlight` 设为 `false` 时 Agent 不再把代码转换为代码框，整段等宽的代码只保护、不处理。
`EnableLists`、`EnableTags`、`EnableTableStyles` 分别控制列表、标记和表格样式工具，设为 `false` 时这些只保护、不修改。
`EnableMarkdownCleanup` 设为 `false` 时不提供去除 Markdown 标记的工具，`EnableClearFormat` 设为 `false` 时不提供清除格式的工具。
`EnableBlankLineRemoval`、`EnableIndent`、`EnableMoves`、`EnableInsert`、`EnableTextTables`、`EnableCodeUnwrap` 分别控制删空行、调整缩进、移动段落（含跨文本框移动和合并文本框）、插入段落、转表格和拆开代码框，
设为 `false` 时不提供对应工具，不想让 Agent 改段落结构时全部关掉。
读取图片文字没有开关：页面上有 OneNote 识别出文字的图片时才提供，每张最多 4000 字。

`Agent/TimeoutSeconds` 是整个任务的总时限，根节点的 `TimeoutSeconds` 仍是每次 HTTP 请求时限。
`MaxPageChars` 限制处理范围内的可编辑文字（含待转换的等宽代码；放不下时这些代码只保护、不转换），可设 1000–1000000；
`MaxRequestChars` 限制每轮包含工具定义和历史的请求体字符数，可设 16000–3000000。
两者的默认值按 DeepSeek 的 1M token 上下文估算（约 0.6 token/汉字，100 万字最多约 60 万 token），用上下文较短的模型（如 128K）时应调小，比如 120000。
超限会停止并丢弃未提交草稿，报错里有实际字数和上限，可以缩小范围或调大预算再执行。另有固定的 1000 段上限。
`MaxTurns` 可设 2–60，`MaxToolCalls` 可设 6–200。轮数快用完时 Agent 先收尾：剩 2 轮时提醒模型，最后一轮只能提交已完成的草稿。
只有最后一轮允许未读完时提前提交；已核验写入改动但还有未读段落时显示「部分完成」，结果和日志单独列出未读数量，已核验的改动仍可撤销。没有写入时显示「没有写入修改」，无法确认写回时仍显示「写回未确认」。
没做完的部分可以缩小范围再执行一次。核验失败日志记录对象 ID、对象类别和固定原因，不记录笔记正文、工具参数或思考；撤销的结果不沿用执行时的未读数量。
某一轮输出达到根节点的 `MaxTokens`（含思考）被截断时，这一轮的工具不执行，Agent 丢掉这一轮、提醒模型分批后重试，已排的草稿保留；一个任务最多重试 2 次，仍被截断就停止并提示调大 `MaxTokens`（0 表示用接口默认值）或降低思考强度。每轮的结束原因和 token 用量记在日志里。

默认字体可选 `Microsoft YaHei`、`Calibri`、`Arial`；未安装时选择其中已安装的一种。
`Enable...` 开关默认开启，已有本机 Office16 回存验证；其他 Office 构建如有兼容问题，可分别关闭原生标题、段间距、图文混排、代码框转换、列表、标记、表格样式或清除格式支持。
`EnableMixedOutlines` 只允许图片和文字混排，不开启墨迹或附件编辑。
如果兼容接口不接受扩展参数，按其文档关闭 `SendThinking`（不发送 thinking/reasoning_effort）、
`ReplayReasoning`（不回传 reasoning_content）或 `StreamUsage`（不发送 stream_options）。模型本身仍必须支持工具调用。
Agent 请求失败时按原因提示：401/403 检查 Key 和权限，429 检查限流或额度，502/503 检查网关及模型服务；网关明确返回 `model_not_found` 时提示模型不存在或未开通。
HTTP 失败不会提交草稿；日志只记录状态和已识别的错误类别，不记录上游错误正文。

## 环境要求

- OneNote **桌面版**（Office16 的 `ONENOTE.EXE`）。UWP 版「OneNote for Windows 10」没有 COM 接口，用不了。
- .NET Framework 4.8 运行时（Windows 10/11 自带）。
- 构建需要 .NET SDK 或 VS2022。

## 构建与安装

安装和卸载脚本支持 **Windows PowerShell 5.1** 和 **Windows 上的 PowerShell 7.x**，最低版本为 5.1。
脚本保留 UTF-8 BOM，以便 Windows PowerShell 正确解析中文。

根目录的 `install.ps1` 一步搞定「关闭 OneNote → 构建 → 写注册表 → 重开 OneNote」，任选一种命令运行：

```powershell
# Windows PowerShell 5.1
powershell -ExecutionPolicy Bypass -File install.ps1

# PowerShell 7.x（需已安装）
pwsh -ExecutionPolicy Bypass -File install.ps1
```

**需要管理员权限**，脚本会自己弹 UAC 提权（原因见下面「踩过的坑」：COM 类必须注册到 HKLM）。
提权后会开一个新的 Windows PowerShell 窗口执行，窗口会留着让你看结果；从 `pwsh` 启动时也一样。
COM 自检同样使用 Windows PowerShell。

改了代码要重装，同一条命令再跑一遍即可。常用参数：

| 参数 | 作用 |
|---|---|
| `-Configuration Debug` | 装 Debug 版（默认 Release） |
| `-SkipBuild` | 跳过构建，直接注册已有输出 |
| `-Uninstall` | 卸载 |
| `-Force` | OneNote 不肯退出时强制结束进程。默认只礼貌请求，失败就停下来，免得丢掉未保存内容 |
| `-NoRestart` | 完成后不自动重开 OneNote |

卸载时运行根目录的 `uninstall.ps1`，会自动提权并关闭 OneNote，任选一种命令：

```powershell
# Windows PowerShell 5.1
powershell -ExecutionPolicy Bypass -File uninstall.ps1

# PowerShell 7.x（需已安装）
pwsh -ExecutionPolicy Bypass -File uninstall.ps1
```

卸载脚本也支持 `-Force` 和 `-NoRestart`。关闭 OneNote 后，脚本会等待本插件的
`dllhost.exe` 代理进程退出；5 秒后仍在运行时，只结束当前会话中 AppID 与本插件匹配的进程。
无法确认代理进程已退出时会停止卸载并报错。卸载保留设置和日志。

也可以只跑注册这一步（前提：**管理员身份的 PowerShell**、OneNote 已关闭、且已经构建过）：

```
powershell -ExecutionPolicy Bypass -File Tools\register.ps1 -Configuration Release
```

```
powershell -ExecutionPolicy Bypass -File Tools\unregister.ps1
```

装好后重新打开 OneNote，「开始」选项卡末尾就有「代码高亮」组了。

## 出问题时

1. **按钮没出现** — 先看日志 `%LOCALAPPDATA%\OneNoteCodeHelper\log.txt`。
   再看 OneNote 的「文件 / 选项 / 加载项」里本项的状态，以及注册表
   `HKCU\Software\Microsoft\Office\16.0\OneNote\Resiliency\DisabledItems`
   — 里面有东西就说明 `OnConnection` 抛过异常。
2. **插件停在「非活动应用程序加载项」** — 说明 `LoadBehavior` 是 2（不随启动加载）。
   OneNote 只要加载失败过一次就会把它从 3 降成 2，之后就再也不尝试了，
   所以后续启动连日志都不会生成，很容易误判成「代码没跑」。
   `register.ps1` 每次都会把它写回 3；也可以在「文件 / 选项 / 加载项 /
   管理: COM 加载项 / 转到」里直接勾上，那会立刻触发一次加载，有问题会当场看到。
3. **改了代码重新构建失败，提示 MSB3021 文件被占用** — 插件运行在系统代理进程
   `dllhost.exe /Processid:{441360A0-59D3-4969-9F91-7AF166C512BE}` 中。
   先关闭 OneNote，再等该代理进程退出。`install.ps1` 会在构建前检测 DLL 是否被占用。
4. **调试** — 用 VS2022 附加到上述 `dllhost.exe` 进程；OneNote 本身不会加载插件 DLL。

## 代码结构

```
AddIn.cs                    外接程序入口：IDTExtensibility2 + IRibbonExtensibility + 功能区回调
Ribbon.xml                  功能区定义（嵌入资源）
Interop/OfficeInterfaces.cs 手写的 Office COM 接口声明
Interop/NativeMethods.cs    Win32 调用（以 OneNote 窗口为属主的提示框）
Services/
  OneNoteApi.cs             OneNote COM 封装
  PageEditor.cs             选区解析、就地替换、插入
  CodeBlockBuilder.cs       生成 one:Table 代码框 XML
  OneNoteHtmlEncoder.cs     Token -> one:T 里的 span HTML
  AddInSettings.cs          设置与持久化
  AddInLog.cs               文件日志
  RenderDiagnostics.cs      不碰 OneNote 就能跑通渲染链路的诊断入口（AI 助手的测试也走这里）
  AiConfig.cs               ai-settings.xml 的模型与读写（保存时就地更新、保留注释），思考强度的五档及其请求参数
  AiClient.cs               Chat Completions 接口（HttpClient + JavaScriptSerializer，流式 SSE）
  AiOptimizer.cs            文字功能的编排：读段落 → 分批并发问 AI → 写回；固定的输入输出约定
  LiveText.cs               Agent 窗口的思考摘录（挑出最新的中文核心句），以及边收边数 AI 已返回几段修改的 JSON 扫描
  RichParagraph.cs          把 AI 改过的纯文本按字符合并回带格式的 one:T
  TextDiff.cs               逐字符 diff（公共前后缀 + LCS）
  BlankLines.cs             排版优化附带的删多余空行（空行段落、段内连续 <br>），不经过 AI
  PageEditCoordinator.cs    按页面串行提交，防止插件自身交叉写回
  Agent/                   工具协议、模型循环、格式编辑、快照、冲突核验和撤销
Highlighting/
  TokenKind / Token / ILanguage / LexerCursor / LanguageRegistry
  DetectionSample.cs        自动识别的样本：原文开头一段 + 去掉注释和字符串内容的同长文本
  LikelihoodPatterns.cs     自动识别打分用的正则：实例缓存 + 匹配超时
  Languages/                每种语言一个 ILanguage 实现；XML 与 HTML 共用 MarkupLexer，
                            HTML 的 <style>/<script> 分别交给 CSS/JavaScript 着色；
                            TypeScript 共用 JavaScript 的词法；PHP 标签外的内容交给 HTML 着色；
                            CommonScanners 放 C 系语言共用的注释（含可嵌套的块注释）、字符串、数字、插值字符串扫描；
                            CFamilyFeatures 放 Java/C#/C++/JS/TS 共有的识别特征
  Themes/CodeTheme.cs, CodeThemes.cs
Views/
  InsertCodeWindow.xaml     插入代码窗口
  CodePreviewRenderer.cs    用同一套 token 流渲染 WPF 预览
  AgentWindow.xaml          AI 助手窗口：功能、模型、思考强度、范围，Agent 需求或文字功能提示词，进度、结果和撤销
  AiSettingsWindow.xaml     AI 配置窗口：接口、模型、文字功能、Agent 四页，保存时写回 ai-settings.xml
  WindowIcons.cs            窗口图标，和功能区按钮共用嵌入资源里的图片
install.ps1                 一键构建 + 安装 / 卸载
uninstall.ps1               一键卸载入口
Tools/register.ps1          只做注册这一步
Tools/unregister.ps1        只做注销这一步
Tools/detect-test.ps1       自动识别回归测试，样本在 Tools/detect-samples/<语言 id>/ 下
Tools/ai-merge-test.ps1     AI 助手回归测试：格式合并、模型输出解析、配置保存；加 -Live 用本机配置真调一次接口
Tools/highlight-selection-test.ps1  「高亮选中」回归测试：认选区、缩进的段落、换成代码框
Tools/agent-test.ps1        Agent 离线回归，不调用真实接口或 OneNote
Tools/addin-surrogate-test.ps1  安装/卸载脚本的编码、语法和代理进程清理离线回归
Tools/agent-format-probe.ps1  显式创建专用测试分区和测试页，验证真实 OneNote 格式往返
Tests/                    独立签名的 net48 测试程序及页面/HTTP 模拟器
```

### 回归验证

```powershell
dotnet build OneNoteCodeHelper.sln -c Release
powershell -ExecutionPolicy Bypass -File Tools\detect-test.ps1
powershell -ExecutionPolicy Bypass -File Tools\ai-merge-test.ps1
powershell -ExecutionPolicy Bypass -File Tools\highlight-selection-test.ps1
powershell -ExecutionPolicy Bypass -File Tools\agent-test.ps1
```

这些回归不需要运行中的 OneNote，也不使用真实 API Key。脚本支持 `-DllPath` 指定待测 DLL。

安装/卸载脚本的编码、语法与代理进程清理逻辑可在两个版本下分别离线验证，无需构建：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File Tools\addin-surrogate-test.ps1
pwsh -NoProfile -ExecutionPolicy Bypass -File Tools\addin-surrogate-test.ps1
```

该检查只解析安装、卸载及辅助脚本，并使用模拟进程验证清理逻辑，不执行真实安装、卸载或 COM 注册。

需要验证 Office 回存时，先构建并运行 Agent 离线测试，再**单独**执行：

```powershell
powershell -ExecutionPolicy Bypass -File Tools\agent-format-probe.ps1 -OutputDirectory D:\temp\onenote-agent-probe
```

该探针会通过 COM 打开 OneNote，在指定目录新建专用 `.one` 测试分区及合成测试页，验证格式提交与撤销，
保存前后 XML，并保留测试分区供检查；不会编辑已有笔记。普通回归不运行这个探针或安装/注册脚本。

## 加一种语言

1. 在 `Highlighting/Languages/` 下实现 `ILanguage`：`Tokenize` 切 token，`ScoreLikelihood` 给自动识别打分。
   `Tokenize` 必须保证返回的 token 按序、不重叠、完整覆盖整个源码 — `RenderDiagnostics.VerifyCoverage`
   就是用来验这条契约的。
2. `ScoreLikelihood` 拿到的是 `DetectionSample`：
   - 关键字、结构类的特征在 `sample.Code` 上匹配。它把整行注释和字符串内容换成了空格，
     注释里的一句 `public class`、字符串里拼的 SQL 就不会给别的语言加分。
   - 要看注释标记或字符串内容的特征（C# 的 `///`、Bash 的 `"$1"`、JSON 的键）才用 `sample.Raw`。
   - 正则走 `LikelihoodPatterns`；多行模式下行首缩进写 `^[^\S\r\n]*`，
     不要写 `^\s*`（`\s` 会跨行，连续空行一多就是平方级回溯）。
   - C 系族里的语言先加上 `CFamilyFeatures.Score(sample)`，自己只写独有的特征。共有特征各语言分数相同、
     互相抵消，胜负才取决于独有写法。族外的语言哪怕语法像 C（比如 PHP）也不要调：族外打平会变成无法确定，
     只靠共有特征取胜的 Java/C# 片段就认不出来了。
   - 可以扣分：本语言里不可能出现的写法（比如 Python 里的 `) {`）是很强的反证。
3. 在 `LanguageRegistry.All` 里加一行。高亮效果和现有某一族几乎一样的，顺便加进 `LanguageRegistry` 的族表。
4. 在 `Tools/detect-samples/<语言 id>/` 下放几段样本（短片段、长文件都要有），然后跑一遍回归测试，
   确认新语言认得出、也没有把别的语言抢走：

   ```
   powershell -ExecutionPolicy Bypass -File Tools\detect-test.ps1
   ```

   加 `-ShowScores` 能看到每段样本各语言的得分，调权重时用得上。

功能区下拉、设置、预览都会自动跟上，不需要改别的地方。

## 自动识别怎么判定

各语言打分后，最高分至少 3 分，而且要是次高分的 2 倍以上、或者高出 4 分以上，才算认出来；
否则提示手动选，因为猜错语言比不猜更糟。有两个例外：

- 整段以标签开头、以 `>` 结尾的，只在 XML、HTML 和 PHP 之间选，里面的 `<script>` 再像 JS 也不算。
  PHP 文件和模板也以 `<?php` 或 HTML 标签开头；含 `<?php`、`<?=` 的页面把 HTML 的得分也记给 PHP，所以会认成 PHP。
- 高亮效果几乎一样的语言算一族（C 系：Java/C#/C++/JavaScript/TypeScript；标记：XML/HTML）。整族当成一个候选
  跟族外的最高分比，比得过就取族内分最高的那个。族内打平时，认错的代价只是个别关键字颜色不对。
  TypeScript 的分数是 JavaScript 的分数加上 TS 独有的写法，出现 TS 写法才认成 TS，否则两者打平、取 JavaScript。
  Go、Kotlin、Rust、PHP 的关键字差别大，不进族。

## 几个踩过的坑

这些结论都是在本机实测出来的，改代码时别踩回去：

- **必须早绑定调 OneNote COM**。这台机器上 OneNote 的类型库没有注册到 CLR 能找到的位置，
  `Type.InvokeMember` 抛 `TYPE_E_LIBNOTREGISTERED`、`dynamic` 抛 `E_FAIL`，
  而同一线程上 PowerShell 的原生晚绑定却是好的。所以项目引用了 `lib/` 下的 PIA。
- **`Windows.CurrentWindow.CurrentPageId` 返回空字符串**（OneNote 16.0.20326），
  拿它去 `GetPageContent` 会抛 `0x80042005`。只能从层级树里找 `isCurrentlyViewed="true"`。
- **`one:T` 虽然是 CDATA，内容仍按 HTML 解析**，`<` `>` `&` 必须转义，
  否则 Java 泛型 `List<String>` 会被整段吞掉。
- **行首缩进和空行会被折叠**，必须用 `&nbsp;` 顶住。
- **提交时省略 `one:List` 不会去掉列表**（16.0.20326 实测）。已有段落少了 `one:List`，OneNote 照样保留原来的列表；
  空的 `<one:List/>` 被架构拒绝，`bullet="0"` 只是换成另一种符号。要去掉列表只能去掉段落的 objectID 让 OneNote 重建它，
  下级段落保持原 ID。而省略 `one:Tag` 会删掉标记，没人引用的 `one:TagDef` 也会被 OneNote 一并清掉。
- **整框回传时，段落的增删和移动按 XML 走**（16.0.20326 实测）。同一文本框里调换顺序、挂到别的段落下、提到上一级的段落都保留 objectID；
  省略的段落被删除；没有 objectID 的段落和表格按新对象建立。新表格的 `one:Column` 必须带 `width`，但没锁定的列宽会被 OneNote 按内容重新计算。
- **跨文本框的对象一律新建**（16.0.20326 实测）。把段落连同原 objectID 放进另一个文本框，OneNote 也按新对象建立并分配新 ID；
  只写目标文本框、不写源文本框时，源文本框里的原段落照旧留着。图片在 piBasic 里只有 `one:CallbackID`，拿它新建会成为坏图
  （piBinaryData 读取时整张图消失），跨框时必须按 piBinaryData 读到的 `one:Data` 写入。
- **`one:Outline` 不能没有段落**：不带 `one:OEChildren` 或 `one:OEChildren` 为空都被架构拒绝。把文本框写成只剩一行空白时，
  OneNote 会直接删掉这个文本框（16.0.20326 实测）；`DeletePageContent` 也能删文本框。新建的文本框按位置排进页面 XML，不一定排在最后。
- **在 OneNote 里行首按 Tab 不是插入制表符**，而是把这一段挂到上一段的 `one:OEChildren` 下。
  直接在笔记里敲的代码，选区里的段落父节点各不相同，不能按父节点判断是不是「同一块」；
  「高亮选中」按所在的文本框 / 表格单元格判断，每深一层补一个制表符。
- **OneNote 不认 CSS 字体栈**。给 `font-family:Consolas,NSimSun` 它只取 `Consolas`；
  而且遇到中文时会把 `font-family` 整个丢掉、回退到自己的中文字体。
  所以代码里有中文注释又想对齐，只能整体换成中英文都等宽的字体（插入窗口里可以选「NSimSun」新宋体）。
- **WPF 窗口必须在 STA 线程上创建**。OneNote 对代理进程的功能区回调在 MTA 上运行；
  插入窗口使用独立 STA 线程和 `ShowDialog()` 消息循环。
- **功能区回调里不做耗时的事，也不同步弹框**。回调返回之前 OneNote 界面是停住的；
  在回调里弹的 `MessageBox` 没有属主，常被压在 OneNote 后面，看起来就是 OneNote 卡死了。
  所以「高亮选中」在后台线程里跑，提示框也在单独的线程里以 OneNote 窗口为属主弹出。
- **OneNote 以本地服务器方式激活 COM 加载项**。只注册 `InprocServer32=mscoree.dll`
  会返回 `0x80040154`，OneNote 随即把 `LoadBehavior` 降为 2。注册脚本为该 CLSID
  配置相同的 AppID 和空值 `DllSurrogate`，由 Windows 的 `dllhost.exe` 承载托管 DLL。
- **注册脚本必须存成 UTF-8 with BOM**。PowerShell 5.1 默认按 ANSI 读 `.ps1`，
  没有 BOM 的中文注释会让脚本直接解析失败。
- **COM 类必须注册到 HKLM，写 HKCU 没用**。`mscoree.dll` 的 `DllGetClassObject`
  只读 HKCR 的机器部分，不读每用户部分。本机实测：同一个 CLSID 写
  `HKCU\Software\Classes` 激活失败（`0x80070002`），写 `HKLM\SOFTWARE\Classes` 成功；
  而同一个 HKCU 键换成原生 DLL 则正常加载，证明不是 HKCU 注册本身的问题。
  这也是 RegAsm 只写 HKLM 的原因。所以注册必须管理员，只有 OneNote 的外接程序清单项
  （`Office\OneNote\AddIns`）留在 HKCU。
  失败时的表现很有迷惑性：OneNote 不报错，只是把 `LoadBehavior` 从 3 悄悄改成 2，
  功能区里什么都不出现，日志文件也不会被创建（因为 `OnConnection` 根本没被调用）。
- **程序集必须强名称签名**。Fusion 的规则是 `codeBase` 指向应用程序目录之外时只对
  强名称程序集生效。本外接程序的 DLL 在自己的目录里、宿主却是 `ONENOTE.EXE`，
  不签名的话注册表里的 `CodeBase` 会被直接忽略。
- **在 dllhost 里发 HTTPS 要显式开 TLS1.2**。插件的 AppDomain 由原生的 dllhost 创建，拿不到目标框架信息，
  `ServicePointManager` 会按老默认值只开 SSL3/TLS1.0，现在的 HTTPS 服务一律握手失败。
  `AiClient` 的静态构造里给 `SecurityProtocol` 加上了 `Tls12`。
- **AI 请求要走流式**。不走流式时整个结果生成完才有第一个字节：思考得久一点窗口就一动不动，
  接口卡住时只能干等到总超时。而网关（One API）的日志要等请求结束才记，那段时间在后台也查不到这个请求，
  看起来就像「插件根本没发请求」。排查时看 `log.txt` 里每次请求的「首包」时间，
  以及 Clash 之类代理的连接日志里有没有 `dllhost.exe → 接口域名` 的记录。
- **引用 `System.Web.Extensions` 要连 `System.Web` 一起引**。前者引用了后者，SDK 版的 XAML 编译器
  （`MarkupCompilePass1`）解析不到就报 `MC1000: Could not find assembly 'System.Web'`，看起来像 XAML 写错了。
- **`-replace` 的第一个参数是正则**。`$path -replace '\', '/'` 里的单个反斜杠是非法模式，
  会直接抛 `InvalidRegularExpression`。要替换路径分隔符用 `.Replace()`。
