# Agent 实施与验证记录

日期：2026-09-26。已实现第一版页面格式 Agent，保持 .NET Framework 4.8、WPF、强名称签名和现有 Office PIA 早绑定引用。没有新增第三方生产依赖。

## 实际执行链路

1. 「开始 → AI 助手 → Agent」通过 `AddIn.OnShowAgentWindow` 启动独立 STA 窗口，固定当前页及打开窗口时的选区。
2. 用户选择整页或选中段落并输入需求。后台读取固定页面，构建短 ID、保护范围、内容和格式指纹。
3. `AgentRunner` 把模型原生 `tool_calls` 映射到七个本地函数，完整读取后才能修改段落。模型的文字声明不能触发写回。
4. 格式工具只更新内存草稿。每个批量调用全部校验通过后才发布，返回修订号；参数错误不会留下半个工具调用的修改。
5. 模型独立调用 `finish_edit` 后，提交器在页面锁内重读、识别冲突、重建完整的受影响容器、带时间戳提交，并回读核验。
6. UI 展示实际核验结果，保存本次已确认段落的撤销记录。撤销再次校验指纹，避免覆盖后来编辑的内容。

页面正文作为数据传给模型，不作为系统指令。工具不接受任意 XML、HTML、外部页面 ID、文件路径或脚本。修改范围不能超出快照；图片二进制、代码和受保护对象不传入模型。模型能读取的正文和富文本格式会发送给用户配置的模型接口。

## 七个工具

本地校验和 JSON Schema 共用 `AgentSchema` 定义；所有对象拒绝未知字段，限制数组长度、数值、枚举和字符串长度。

| 工具 | 参数与执行规则 |
|---|---|
| `get_page_overview` | 可选 `offset`，每页最多 100 条；返回快照 ID、短段落 ID、最多 80 字摘要、嵌套深度、保护原因、可用预设与能力开关 |
| `read_blocks` | `snapshot_id`、`block_ids`；最多 100 段，返回完整文字、富文本 runs、当前草稿样式；读取受保护段落会拒绝 |
| `set_paragraph_style` | `snapshot_id`、`block_ids`、`preset_id`、可选 `overrides`；预设为 page_title/heading1/heading2/body/quote，page_title 只用于原标题 |
| `set_text_style` | `snapshot_id`、`targets`；每项指定 `block_id`、原文 `quote`、从 1 开始的 `occurrence`、`style`；支持 bold/italic/underline/color |
| `highlight_code` | `snapshot_id`、`block_ids`（最多 1000）、`language`（`auto` 或已支持的语言 id）；把连续的代码段落排入草稿，提交时换成高亮代码框。`Agent/EnableCodeHighlight=false` 时不注册 |
| `get_pending_changes` | `snapshot_id`；返回草稿修订号、修改段落 ID、未完整读取的可编辑段落 ID、已排入的代码框、尚未转换的等宽代码及保护计数 |
| `finish_edit` | `snapshot_id`、`draft_revision`；必须是当轮唯一工具，修订号匹配后冻结草稿，只提交一次 |

段落 overrides：`alignment` 为 left/center/right；`font_family` 为 Microsoft YaHei/Calibri/Arial 且需已安装；字号 8–32 pt；段前后间距 0–36 pt；颜色限于 `#1F4E79`、`#365F91`、`#222222`、`#666666`。
局部文字按精确匹配定位，重复短语的出现序号按不重叠匹配计数。切开代理对、组合字符或常见 emoji 序列的请求会拒绝。

## 代码框转换

代码段落分三类：单行单格表格里、格内全是等宽段落的视为已有代码框（`highlighted_code`），受保护；整段非空白字符都是等宽字体的是待转换代码（`unhighlighted_code`），可读取、可转换、不能设置样式；只有部分文字等宽的是行内代码（`protected_code`），受保护。普通字体输入的代码仍是可编辑正文，由模型判断后转换。

`highlight_code` 复用「高亮选中」的 `CodeSelection` 拼源码（每深一层补一个制表符）、`LanguageRegistry.Resolve` 定语言、`CodeBlockBuilder.BuildTable` 生成代码框，主题和字号等取打开窗口时的功能区设置。为保证能原样撤销，只接受同一文本块里连续的段落，第一段在最外层，每段的下级段落都在范围内，且不含列表或待办标记；代码中间的空段落可以一并传入。转换会丢弃这些段落已排的格式草稿。

提交时和格式修改在同一次重读、写入、回读中完成。任何源段落的指纹变化都会跳过整个代码框。回读核验中，写入前不存在的 objectID 视为新建对象；新代码行和还原段落只比较规范化的纯文字（硬空格按空格、去掉行尾空白），因为 OneNote 会改写代码行的 span。核验通过后记录代码框 Table 的 ID 和指纹（Table ID、底色、各行文字，不含会被 OneNote 改动的列宽和外层段落 ID）。

撤销时代码框指纹不变才换回原段落：去掉 objectID 和编辑记录让 OneNote 新建，`quickStyleIndex` 按当前页面的样式定义重新对应，再核验文字和格式。

## 代码位置

| 文件 | 职责 |
|---|---|
| `Ribbon.xml`、`AddIn.cs` | 按钮、STA 窗口启动、避免重复窗口、关闭时取消任务 |
| `Views/AgentWindow.xaml(.cs)` | 范围、需求、进度（轮次、思考摘录、执行步骤）、取消、结果、会话撤销 |
| `Services/Agent/AgentRunner.cs` | 模型循环、历史消息、工具分派、调用幂等和预算 |
| `Services/Agent/AgentChatClient.cs` | Chat Completions、HTTP/SSE、工具片段聚合、消息 DTO |
| `Services/Agent/AgentTools.cs` | 六个工具、Schema 与本地校验、草稿发布 |
| `Services/Agent/AgentFormatting.cs` | 富文本解析与局部样式、等宽判定、CSS 归一化、预设、QuickStyleDef 管理 |
| `Services/Agent/AgentCode.cs` | 代码框转换的范围校验、原段落记录、代码框指纹和撤销还原 |
| `Services/Agent/AgentPageSnapshot.cs` | 范围、短 ID、保护对象、指纹、配置和语义投影 |
| `Services/Agent/AgentCommitter.cs` | 时间戳提交、冲突重建、核验、撤销 |
| `Services/PageEditCoordinator.cs` | 有界的页面提交锁，与原有编辑器共用 |
| `Services/OneNoteApi.cs` | 显式 xs2013、非强制更新、COM 在途调用和断开协调 |
| `Tests/`、`Tools/agent-test.ps1` | 独立签名测试程序、Fake 页面、脚本模型、HTTP handler 模拟 |
| `Tools/agent-format-probe.ps1` | 显式创建并编辑专用合成页面的真实 COM 探针 |

与原设计的文件布局相比，协议 DTO 留在 `AgentChatClient.cs`，工具注册表留在 `AgentTools.cs`，两个样式编辑器合并为 `AgentFormatting.cs`。测试使用单独的 net48 EXE，没有向生产 `RenderDiagnostics` 增加测试接口。

## 协议、限额与异常

- 单轮请求发送 `tools`、`tool_choice=auto`、`stream=true`，不沿用原 AI 优化的 JSON 输出模式。支持 SSE 以及服务端返回的普通 JSON 消息。
- SSE 按 `tool_calls[index]` 拼接 arguments，保留工具 ID、名称和 `reasoning_content`；收到有效 `finish_reason` 并通过完整校验前不执行任何工具。
- `length`、无结束原因的截断、流内错误、重复工具 ID 携带不同参数等情况停止执行。重复 ID 且参数相同则复用上次结果。
- 默认最多 12 轮、48 次工具、总时限 600 秒、40,000 个可编辑字符、1000 段；单轮请求体最多 120,000 字符。限制可在允许范围内配置。
- HTTP 单次超时沿用 AI 配置；无新数据 60 秒终止。SSE 传输上限为 8 MiB（包括每个分片重复的 JSON 包装），本轮真正需要保存的思考、正文和工具参数合计上限为 500,000 字符；单条事件与非流式 JSON 上限 600,000 字符，单工具参数 64,000 字符，JSON 最大嵌套 40 层。超限时不执行不完整工具。
- 达到限额或在提交前取消，丢弃草稿。COM 提交已经开始后，即使取消也要先尝试核验结果。
- 不把供应商错误响应正文、工具参数或笔记正文写进新增日志；日志只记录状态、次数和异常类型。

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

## 仍需用户环境验收的部分

本次没有调用真实模型接口，也没有安装、注册或重启当前插件。离线协议测试验证了实际 HTTP 客户端的请求体和 SSE 解析，但不等同于验证用户的供应商、Key 和所选模型。

针对「Agent 响应超过大小限制」的现场报告，修正了流式响应把每个 SSE 分片重复的 JSON 包装计入 600,000 字符上限的问题。现在传输字数与有效输出分别计数，并用 `StringBuilder` 累积大量小片段；构造超过旧限制的流式数据已通过回归验证。DeepSeek 的[思考模式工具调用文档](https://api-docs.deepseek.com/guides/thinking_mode/)要求后续请求回传完整 `reasoning_content`，所以不会截断模型历史。
更新安装后，需要用本机配置运行一次示例需求，检查模型的标题识别和最终排版是否符合偏好；也应在实际加载项中验收开关窗、切页及关闭 OneNote 的生命周期。

墨迹、附件、媒体、未知 HTML、图片位置和任意布局调整不在本版能力内；遇到不支持的对象会保护相关段落或整个容器。对其他 Office 构建的回存兼容性应重新运行探针。
