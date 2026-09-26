# OneNote 页面排版 Agent 技术设计

设计日期：2026-09-26。第一版现已按本方案实现，原生标题、段间距、图文容器和撤销已通过专用合成页面的本机 OneNote 回存验证。
下文保留设计时的架构与验收计划；实际文件映射、配置默认值、测试结果和未完成的人工验收项以 [实施说明](agent-implementation.md) 为准。

## 1. 可行性与建议

可以实施。建议在现有 .NET Framework 4.8、WPF、早绑定 OneNote COM 架构内，增加一个小型工具调用循环：模型读取页面结构，调用格式工具形成修改草稿，插件校验并提交，最后回读页面核对结果。

第一版交付目标是：用户点击「Agent」，输入「将该页面上的内容排版下，要美观」，程序自动统一正文样式、识别已有段落中的标题、设置标题层级和段间距，并按需突出重点文字。排版后正文文字、段落数量、顺序、嵌套关系、链接目标不变。

这项需求有两个不同层面的确定性：

| 层面 | 判断 | 依据或实施要求 |
|---|---|---|
| 页面读取与写回 | 已具备基础 | 仓库已有 OneNoteApi、PageEditor 和正在使用的 AI 优化流程 |
| 模型工具调用 | 可实现，需要扩展客户端 | 现有接口为 Chat Completions；新增 tools、tool_calls 和多轮历史支持 |
| 字体、字号、颜色、加粗和段落对齐 | 技术路径明确 | 操作 OneNote 页面 XML 和 one:T 内受支持的 HTML 样式 |
| 原生标题、段前后间距、混合图文容器往返 | 实施前做本机探针 | XML 表达存在，但当前仓库未验证所有属性组合在目标 Office16 构建上的回存行为 |
| “美观” | 可通过统一样式和语义分类实现 | 第一版采用确定的样式预设，效果仍需用真实笔记验收 |
| 任意自由布局、图片移动、多栏重排 | 不列入第一版 | 涉及几何布局与对象完整性，需要另一组工具和验证 |

微软的 COM 文档支持通过 GetPageContent / UpdatePageContent 读取和修改页面；更新时可以仅提交发生变化的页面级对象，但提交的 Outline 必须包含该容器的完整内容。这个限制决定了提交器必须从最新页面组装完整的受影响容器。[OneNote Application 接口](https://learn.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote#updatepagecontent-method)

不需要引入新的 Agent 服务、MCP 服务端或大型编排框架。这里的“工具”就是插件注册的 C# 函数，模型返回工具名称和 JSON 参数，由插件调用对应实现。

## 2. 当前代码中可复用和必须新增的能力

| 代码入口 | 现状 | 设计处理 |
|---|---|---|
| [Ribbon.xml](D:/dev/OneNoteCodeHelper/Ribbon.xml) | 已有「AI 助手」组 | 加 Agent 按钮 |
| [AddIn.cs](D:/dev/OneNoteCodeHelper/AddIn.cs:460) | OnAiOptimize 创建独立 STA 窗口，后台运行任务 | 复用窗口启动模式，新增 OnShowAgentWindow |
| [AiClient.cs](D:/dev/OneNoteCodeHelper/Services/AiClient.cs:94) | 单轮请求，返回 string；SSE 只拼 content 并统计 reasoning 字数 | 抽取 HTTP/SSE 共用层，新增返回完整消息的 Agent 模型客户端 |
| [AiConfig.cs](D:/dev/OneNoteCodeHelper/Services/AiConfig.cs) | 地址、Key、模型、思考参数和提示词配置 | 复用连接配置，增加可选 Agent 限额与能力配置 |
| [OneNoteApi.cs](D:/dev/OneNoteCodeHelper/Services/OneNoteApi.cs:61) | 从层级树定位当前页，早绑定读取和写回 | 复用；新增接口抽象便于模拟测试，明确 xs2013 |
| [PageEditor.cs](D:/dev/OneNoteCodeHelper/Services/PageEditor.cs:316) | 读取段落、处理选区、生成局部页面更新 | 复用结构知识；新增包含样式、继承关系的页面快照 |
| [RichParagraph.cs](D:/dev/OneNoteCodeHelper/Services/RichParagraph.cs:16) | 纯文本修改合并回原 HTML，保持原格式 | 不能直接承担“设置格式”；新增格式补丁能力，保留原 Apply 行为 |
| [PageEditor.Submit](D:/dev/OneNoteCodeHelper/Services/PageEditor.cs:475) | 捕获带时间戳写回异常后，以 DateTime.MinValue 再提交 | Agent 禁用这个退路，走严格提交器 |

现有「排版优化」主要修改中英文空格、标点等段内文字并清理空行。新的 Agent 需要修改视觉样式，因此仅增加一段提示词无法实现需求。

保持 net48、强名称签名、lib 中 PIA 引用、原有功能区接口和注册方式。生产路径不使用 dynamic、反射调用 OneNote 或运行时类型库查找。

## 3. 第一版功能范围

### 3.1 对用户开放的能力

- 设置已有段落为页面标题、一级标题、二级标题、正文、引用的样式。
- 统一允许修改段落的字体、字号、基础颜色和段前后间距。
- 左对齐、居中、右对齐。
- 对段内指定文字加粗、斜体、下划线、设置文字颜色。
- 处理当前页，或者用户显式选择的“选中段落”范围。
- 取消任务、展示实际修改结果、在当前会话中撤销本次格式修改。

### 3.2 第一版边界

正文、标题文字均不改写，不合并、拆分、移动或删除段落；不更改列表层级、待办状态、表格行列、链接地址。代码段落及包含代码的受保护区域不参与自动格式统一。图片、墨迹、附件等对象只保留，不交给模型解释或修改。

表格中的普通文字可以使用相同的段落格式工具，但不得改变表格结构；混合图文 Outline 必须通过往返探针后才启用，其中遇到不能可靠保留的对象则跳过整个受影响 Outline，并报告原因。

不调用 BlankLines.RemoveFromPage：删空行会改变段落结构，不属于这个纯格式 MVP。将来如需清理空段落，应单独定义工具和保护规则。

这个 Agent 可以执行其工具集合支持的需求。对“把全文改写为三段”“删除页面”等请求，应明确说明当前工具不支持，不能声称已经完成。

## 4. 窗口与功能区

建议按钮定义：

```xml
<button id="OncAgent"
        label="Agent"
        size="large"
        image="AiOptimize"
        screentip="用自然语言整理页面格式"
        onAction="OnShowAgentWindow" />
```

第一版复用现有图标，后续可独立设计图标。窗口建议 620 × 560 DIP，包含：

| 区域 | 内容 |
|---|---|
| 目标信息 | 页面标题；固定目标页，防止用户切页后误写另一页 |
| 范围 | 当前页（默认）、选中段落（有选区时可用） |
| 输入框 | 多行输入，提示示例“统一正文，突出标题，使页面更清晰” |
| 模型 | 显示本次使用的模型和思考强度，复用功能区选择 |
| 控件 | 执行、取消；完成后显示撤销本次、关闭 |
| 进度 | 正在读取、分析结构、拟定格式、写回、验证；展示工具动作摘要 |
| 结果 | 实际修改数、冲突跳过数、不支持对象数、未验证数 |

点击执行即授权在所选范围内应用支持的格式操作，正常执行不再增加确认弹窗。可在后续版本增加“仅生成预览”的选项，第一版不依赖预览才能使用。

启动时在后台捕获页面 ID 和选中段落的 objectID，捕获完成前禁用执行。执行时重新读取固定页面，形成真正的基线快照；选择“选中段落”时以捕获的段落 ID 限定范围。窗口明确标明“选中段落”，因为选中半个段落时，段落级样式仍影响整段。

功能区回调只排队启动，不做页面读取、文件加载、网络请求或同步弹框。WPF 窗口沿用独立 STA 和 ShowDialog 消息循环，网络与 XML 处理在后台执行，通过 IProgress/Dispatcher 更新窗口。

## 5. 模块与执行模型

```mermaid
flowchart TD
    A[Agent 按钮与输入窗口] --> B[AgentRunner]
    B <--> C[AgentChatClient / 模型]
    B --> D[ToolRegistry 参数校验与分派]
    D --> E[页面快照与内存格式草稿]
    E --> F[格式补丁与不变量校验]
    F --> G[PageEditCommitter]
    G <--> H[OneNoteApi / COM]
    G --> I[回读验证与撤销记录]
    I --> A
```

工具产生的是有类型的格式操作，先应用到本地快照的克隆体。模型可以查看草稿、继续调整；finish_edit 通过校验后，宿主统一提交一次页面更新。

这样仍然是完整的“模型调用工具—收到结果—继续决策”的 Agent，同时避免每个加粗动作都触发一次 COM 写回。

建议核心接口如下，均为拟新增内部类型：

```csharp
interface IAgentChatClient
{
    Task<AgentAssistantMessage> CompleteAsync(
        IReadOnlyList<AgentMessage> messages,
        IReadOnlyList<AgentToolDefinition> tools,
        AgentRequestOptions options,
        IProgress<AgentStreamProgress> progress,
        CancellationToken cancellation);
}

interface IAgentTool
{
    AgentToolDefinition Definition { get; }
    Task<AgentToolResult> ExecuteAsync(
        AgentContext context,
        IReadOnlyDictionary<string, object> arguments,
        CancellationToken cancellation);
}

interface IOneNotePageAccess
{
    string GetPageContent(string pageId, PageInfo info);
    void UpdatePageContent(string xml, DateTime expectedLastModified);
}
```

AgentContext 保存：固定 PageId、执行范围、基线快照、短 ID 到 objectID 的映射、已读取段落集合、草稿、格式操作集合、DraftRevision、已执行 tool_call_id 缓存、取消令牌与预算。

## 6. 页面快照与模型可见数据

不能只使用现有 AiParagraph 的 ObjectId 和 Text。新快照需要同时表达文本、语义角色、容器、列表、层级、有效样式和编辑权限。

建议本地类型：

```text
PageSnapshot
  SnapshotId, PageId, PageTitle, LastModifiedUtc
  OriginalPageXml
  QuickStyleDefinitions
  Containers[]
  Blocks[]

PageBlock
  ShortId, ObjectId, Kind
  ContainerId, ParentBlockId, TableCellPath, Depth, Order
  Text, Runs[], EffectiveStyle
  Selected, Editable, ProtectedReason
  BaselineFingerprint
  OriginalElement
```

模型收到的是投影，例如：

```json
{
  "snapshot_id": "s1",
  "page_title": "部署说明",
  "scope": "page",
  "draft_revision": 0,
  "blocks": [
    {
      "id": "p12",
      "kind": "paragraph",
      "container_id": "o1",
      "parent_id": null,
      "depth": 0,
      "text": "部署前准备",
      "style": {"font_size_pt": 11, "bold": false},
      "editable": true
    },
    {
      "id": "p13",
      "kind": "code",
      "container_id": "o1",
      "editable": false,
      "reason": "protected_code"
    }
  ]
}
```

PageId 和真实 objectID 由宿主掌握，工具参数只接受本次快照产生的短 ID。ID 只在当前任务有效；刷新任务必须生成新的 SnapshotId。

快照保留标题、Outline、表格单元格和嵌套 OEChildren 的关系。模型可以将某段识别为标题，但不能因这种识别而把段落搬进另一个容器。

overview 返回段落摘要和对象占位；read_blocks 才返回完整正文及分段样式。只有完整读取且 editable=true 的段落才能修改。正文很长时返回完整的小批段落与游标，不默默截断后让模型误以为看完了。

选区模式只向模型提供允许范围内的正文，其他对象最多作为没有正文的边界占位。模型无权通过另传 pageId 扩大范围。

## 7. 工具协议

第一版注册六个工具。所有参数均校验类型、必填字段、未知字段、长度、枚举和目标权限，不依靠提示词约束执行权限。

| 工具 | 参数 | 返回与副作用 |
|---|---|---|
| get_page_overview | 无 | 返回快照 ID、范围、对象摘要、能力和样式预设；只读 |
| read_blocks | snapshot_id、block_ids | 返回完整文字、样式和层级，记录已读取对象；只读 |
| set_paragraph_style | snapshot_id、block_ids、preset_id、可选 overrides | 设置草稿中的段落预设及受限覆盖，返回 changed/noop 和 DraftRevision |
| set_text_style | snapshot_id、targets | targets 为段落 ID、精确文字、出现序号、允许的样式补丁；只改草稿 |
| get_pending_changes | snapshot_id | 返回草稿修订号、差异、未覆盖对象和校验结果；只读 |
| finish_edit | snapshot_id、draft_revision | 冻结草稿、校验、提交、回读；返回实际结果 |

set_paragraph_style 的最小完整工具定义：

```json
{
  "type": "function",
  "function": {
    "name": "set_paragraph_style",
    "description": "为已完整读取的可编辑段落设置样式；只修改草稿，不改正文和段落结构。",
    "parameters": {
      "type": "object",
      "properties": {
        "snapshot_id": {"type": "string"},
        "block_ids": {
          "type": "array",
          "items": {"type": "string"},
          "minItems": 1,
          "maxItems": 100,
          "uniqueItems": true
        },
        "preset_id": {
          "type": "string",
          "enum": ["page_title", "heading1", "heading2", "body", "quote"]
        },
        "overrides": {
          "type": "object",
          "properties": {
            "alignment": {"type": "string", "enum": ["left", "center", "right"]},
            "font_size_pt": {"type": "number", "minimum": 8, "maximum": 32},
            "font_family": {"type": "string", "enum": ["Microsoft YaHei", "Calibri", "Arial"]},
            "color": {"type": "string", "enum": ["#1F4E79", "#365F91", "#222222", "#666666"]},
            "space_before_pt": {"type": "number", "minimum": 0, "maximum": 36},
            "space_after_pt": {"type": "number", "minimum": 0, "maximum": 36}
          },
          "additionalProperties": false
        }
      },
      "required": ["snapshot_id", "block_ids", "preset_id"],
      "additionalProperties": false
    }
  }
}
```

page_title 只能用于实际 Title 下的段落。模型不能把普通正文转成页面标题对象。默认使用预设；用户明确要求居中或改变字号时才使用 overrides。宿主验证字体是否安装、数值是否有限，并按当前已验证 capability 删除不支持的覆盖字段；不支持的参数返回错误，不静默忽略。未提供覆盖字段沿用预设，对预设本身未管理的字段保留原值。

局部强调调用示例：

```json
{
  "snapshot_id": "s1",
  "targets": [
    {
      "block_id": "p18",
      "quote": "先备份配置",
      "occurrence": 1,
      "style": {"bold": true, "color": "#1F4E79"}
    }
  ]
}
```

quote 必须与快照中解码后的文字精确匹配；occurrence 从 1 开始。宿主定位实际字符范围，避免模型计算 UTF-16 下标。目标不存在、出现序号越界、匹配跨越受保护片段或切开 Unicode 字符边界时拒绝执行，不做模糊匹配。

style 第一版仅允许 bold、italic、underline、color；颜色限制为主题色表。取消某种样式必须显式传 false，未提供字段保持原值。

工具返回统一结构：

```json
{
  "ok": true,
  "draft_revision": 3,
  "changed": ["p12", "p15"],
  "noop": [],
  "errors": []
}
```

整个调用的参数必须先校验，再修改草稿；参数错误时该调用无副作用。正常无变化使用 noop。提交阶段才可能因用户编辑发生部分跳过，返回 skipped_conflict 和详细目标。

工具名通过显式 Dictionary<string, IAgentTool> 分派，不用模型提供的名字查找任意 C# 方法。不提供执行脚本、任意文件访问、任意 URL 请求、原始 XML 写入工具。

## 8. 模型接口与 SSE

继续使用用户配置的 /chat/completions。Agent 请求与旧 AI 优化请求分别建模，避免旧功能受新消息结构影响：

```text
model: 当前配置模型
messages: system + user + assistant/tool 完整历史
tools: 本轮允许的工具定义
tool_choice: auto
stream: true
max_tokens、thinking、reasoning_effort: 由供应商能力适配器生成
```

Agent 请求不沿用旧流程固定的 response_format=json_object。工具参数本身为 JSON，assistant.content 可以为空，也可以是自然语言。

多轮历史必须依次包含：assistant 的完整 tool_calls；每个调用对应一个 role=tool、tool_call_id 匹配的结果；然后才发下一次模型请求。工具失败也作为结构化结果返回，不能丢失对应的 tool 消息。

SSE 按 choice.index 和 tool_call.index 聚合参数片段；保留调用 ID、函数名、原始 arguments 字符串。完整收齐并校验结束原因后才解析参数和执行工具。length、aborted、流中错误、没有结束信号等情况均不能执行该轮工具。工具调用的增量结构与参数校验要求见 [DeepSeek Chat Completions API](https://api-docs.deepseek.com/api/create-chat-completion/)。

AgentAssistantMessage 至少保存 Content、ReasoningContent、ToolCalls、FinishReason、Usage。现有 AiReply 仅保存 reasoning 字数，不能直接承担这个角色。

对 DeepSeek，当前文档要求带 tools 的后续请求回传历史 reasoning_content；只在任务内存中保留用于协议续接；界面显示进度摘要和思考原文的最后三行（只在窗口内），日志不记录原文。[DeepSeek Thinking Mode](https://api-docs.deepseek.com/guides/thinking_mode/)

供应商差异使用 AgentModelCapabilities 配置，包括 SupportsTools、SupportsThinkingWithTools、ReplayReasoningContent、SupportsStrictSchema、SupportsStreamUsage。默认不发送 strict 或其他未经当前接口确认的可选参数。初次接入用无页面数据的只读工具做一次联调；“OpenAI 兼容”不能替代工具能力验证。DeepSeek 思考模式下使用 auto，不能假设 required 或指定函数的 tool_choice 可用。

若模型只返回普通回答却没有提交草稿，不能将“我已排版”显示为成功。Runner 可提醒一次必须调用工具完成操作；仍不调用则结束为“未应用修改”。

## 9. 样式预设与 XML 编辑

建议使用一套内置样式，以下数值是产品初始方案，最终通过本机测试页校准：

| preset_id | 字号 | 基础颜色 | 字重处理 | 段前/段后 |
|---|---:|---|---|---|
| page_title | 20 pt | #1F4E79 | 加粗 | 0 / 12 pt |
| heading1 | 16 pt | #1F4E79 | 加粗 | 12 / 6 pt |
| heading2 | 13.5 pt | #365F91 | 加粗 | 8 / 4 pt |
| body | 11 pt | #222222 | 保留原有局部强调 | 0 / 4 pt |
| quote | 10.5 pt | #666666 | 保留原有局部强调 | 4 / 4 pt |

普通段落默认左对齐，字体采用单个已安装的 Microsoft YaHei；缺少字体时由宿主选用测试过的系统字体，不能给 OneNote 写 CSS 字体栈。字号、颜色和间距均由代码生成，不允许模型直接拼接 CSS。

XML 映射：

| 格式 | 写入位置 |
|---|---|
| 段落字体、字号、基础文字色 | one:OE/@style，以及必要的局部覆盖清理 |
| 局部加粗、斜体、下划线、文字色 | one:T CDATA 中的 span 样式 |
| 左/中/右对齐 | one:OE/@alignment |
| 段前、段后 | one:OE/@spaceBefore、@spaceAfter |
| 原生标题语义 | 页面 one:QuickStyleDef，及段落 quickStyleIndex 引用 |

这些属性在 [OneMore 保存的 OneNote 2013 XSD](https://github.com/stevencohn/OneMore/blob/main/Reference/0336.OneNoteApplication_2013.xsd) 中可查；它是外部参考架构，实际启用仍以目标 Office 的往返结果为准。首版不暴露两端对齐、CSS margin、line-height、多栏等未经验证的能力。

格式解析须计算样式继承：QuickStyleDef、祖先的可继承属性、OE、T 和内部 span。不能只改 OE 的字号，因为更深层的 span 可能覆盖它。OneMore 对样式覆盖及 OneNote 回存归一化的说明可作为实现参考。[OneMore 样式技术说明](https://onemoreaddin.com/developers/TechNote%20-%20Styles.htm)

ParagraphStyleEditor 采用字段级修改：字体、字号可以统一清理其对应的 T/span 覆盖；段落基础色保留链接色和原有局部语义强调，除非用户显式要求统一该局部颜色。body 不清除原有加粗、斜体、下划线。标题加粗作为明确的格式操作处理。

原生标题的 QuickStyleDef 按样式签名查找或新建，index 根据当前页面分配；不得把 heading1 硬编码为 index=2，也不得原地改掉被其他段落引用的定义。提交时在最新页面重新解析 index，防止并发增加样式导致冲突。新增定义在 XML 架构要求的位置输出。

RichTextStyleEditor 对已解析的字符/run 进行分割和合并，只编辑允许的样式属性；保留 a.href、HTML 实体、nbsp、br、未修改的标签属性及多个 one:T 的映射。只将新文字编码为 HTML 的逻辑不能直接拿来做格式统一；更不能把正文整体重新生成一段 HTML。

格式修改前后必须满足：解码后文本逐字符相同，硬空格与换行语义相同，链接目标相同，段落 ID、顺序、嵌套与列表/待办元数据相同。无法保证往返的 HTML 或混合墨迹文字段落应标记为不可编辑。

若本机探针发现原生标题定义无法可靠写回，可以降级为显式字体样式，但报告只能称为“标题外观”，不能称为已经设置原生标题层级。

## 10. Agent 状态机和示例过程

```text
Idle → Capturing → Thinking ⇄ ExecutingTools
                         → Validating → Committing → Verifying
                         → Completed / PartiallyCompleted

提交前取消：Cancelled，丢弃草稿
提交开始后取消：先等待 COM 返回并验证，报告实际结果
协议、接口或校验失败：Failed，不自动提交残留草稿
```

以「将该页面上的内容排版下，要美观」为例：

1. 宿主固定目标页、创建基线快照，提示词要求只修改格式。
2. 模型调用 get_page_overview，获知页面对象、可编辑范围和预设。
3. 模型分批调用 read_blocks，读取正文及原有层级；代码区域只呈现受保护占位。
4. 模型把既有段落分类，批量调用 set_paragraph_style；需要强调的短语用 set_text_style。
5. 模型调用 get_pending_changes，检查草稿改动和遗漏。
6. 模型单独调用 finish_edit，携带最新 draft_revision。
7. 宿主重新读页、检查冲突、构造补丁、写回并回读。
8. 工具返回真实计数，例如“修改 18 段、跳过 2 段冲突、保留 1 个代码区”。可再发一次禁止写工具的摘要请求；即使摘要请求失败，界面仍展示宿主生成的实际结果。

finish_edit 必须是所在 assistant 消息中唯一的工具调用，且修订号必须与当前草稿一致。收到后冻结草稿，不能再执行新的编辑工具。一次任务最多提交一次；用户追加需求开始新的任务快照。

普通工具调用在同一任务中顺序执行。相同 taskId + tool_call_id + 参数摘要命中缓存时返回旧结果，不重复执行；同一个调用 ID 携带不同参数视为协议错误。赋值型格式操作本身也应幂等。

建议默认限额：模型请求 12 轮，工具调用 48 次，单次批量目标 100 个，累计编辑目标 1000 个，整页正文 40000 字符。单次 HTTP 总超时复用现有配置，60 秒无数据超时继续生效，整个任务另设 600 秒上限。

请求预算须包括工具定义、历史工具结果和回传的 reasoning_content。字符上限只做应用保护，不能当作 token 的精确换算；按供应商上下文上限另留余量。超限要求缩小范围或分成新任务，不静默丢弃历史工具消息，也不提交未完成草稿。上述数值均为初始默认值，需要配置和样本校准。

## 11. 并发冲突与严格提交

现有 ApplyParagraphEdits 主要比较纯文本，能发现用户改字，但不能发现“文字没变，只调了字号或加粗”。格式 Agent 必须比较格式状态。

每段基线指纹使用 SHA-256，输入为稳定规范化表示：正文与行内 HTML 语义、OE/T 样式、列表和标签信息、父容器及祖先路径、影响有效样式的继承属性、引用的 QuickStyleDef 内容。仅排除已确认的瞬态字段，如 selected 和编辑时间；保留未知字段用于保守检测。空格、nbsp、br 不能被规范化丢失。

提交算法：

```text
取得页面写入锁
重新 GetPageContent(固定 PageId)
逐段比较基线指纹
    被用户改变、删除或移动 → 跳过该段
    未变 → 在最新页面副本上重放该段格式操作
验证文字、结构、保护对象不变量
从最新副本取出完整的受影响 Title/Outline 及必要样式定义
UpdatePageContent(changes, 最新 lastModifiedTime, xs2013, false)
GetPageContent 再读回验证
释放页面写入锁
```

如果发生 0x80042010（hrLastModifiedDateDidNotMatch），重新读页、重新比对、重新生成 XML，最多重试 2 次。不能重发旧 XML，不能改成 DateTime.MinValue，也不能用 force=true 绕过。[OneNote 错误码](https://learn.microsoft.com/en-us/office/client-developer/onenote/error-codes-onenote)

如果不能解析页面时间戳，严格模式下不写入，提示无法完成并发校验。只有确认为时间戳冲突的错误才进入上述重试；只读页、无效 XML、被锁定的分区等错误直接报告。

指纹基线固定为模型实际读取的那一版。发生冲突时不能自动把用户新改的段落当作新授权基线并继续覆盖。无冲突段落可继续应用，冲突段落明确列出。

即使只改一个段落，也必须保留它所在 Outline 的其他段落、表格、图片引用等内容。图片/墨迹的 CallbackID 不能变成空对象；若当前读取方式不能充分保留该容器，提交器拒绝它。不能把单个 OE 当成独立页面级补丁提交。

进程内增加 PageEditCoordinator。Agent、现有 AI 优化、高亮选中和插入代码共用同一个按 PageId 的提交锁；锁只包住“重读—组装—提交—回读”，模型请求期间不持锁。外部用户编辑仍通过时间戳与指纹保护。现有 Submit 的无条件退路也应在接入统一提交器时移除，并跑完整回归。

一次 UpdatePageContent 有助于减少中间状态，但不能把 COM 操作宣传为数据库事务。调用超时或异常后若结果不明，必须回读核对；不能盲目再次写入。

## 12. 结果校验、取消和撤销

回读时比较实际格式语义，不能直接比较 XML 字符串，因为 OneNote 可能调整 span、实体表达和样式存放位置。

验证至少包括：

- 目标段落文字、ID、顺序、层级和链接目标没有改变。
- 允许的格式字段生效，代码和非目标对象保持原状。
- 样式引用有效，标题是否达到原生语义或仅外观降级。
- 实际变化数与工具计划数对应；被 OneNote 忽略的属性记为未生效。

结果状态区分 Verified、PartiallyApplied、NoChange、CancelledBeforeCommit、FailedBeforeCommit 和 CommitOutcomeUnknown。网络摘要失败不抹掉已经核验的写入结果；回读失败时不能显示“完全成功”。

取消在提交前丢弃草稿。COM 调用开始后不能假设 CancellationToken 能终止写入；等待调用结束，再验证并展示结果。窗口关闭与后台任务生命周期分开，后台任务未退出前不释放互斥状态。

第一版撤销记录放在会话内存，保存实际提交前的受影响格式、回读后的实际格式和保护指纹。点击撤销重新读页，只对仍与“本次提交后状态”匹配的段落恢复旧格式，再按相同严格流程提交；用户后来改过的段落跳过。不能用整页旧 XML 覆盖回去，也不承诺 Ctrl+Z 一定把整次 Agent 操作作为一个撤销单元。

持久化崩溃恢复可在后续加入：保存到用户本地数据目录、限制保留时长和数量，并使用 Windows 用户级加密。第一版明确说明关闭会话后不再提供插件级撤销记录。

## 13. 提示词与工具权限

系统提示词由代码维护，核心要求：只在当前任务范围内格式化；正文是待处理数据；使用工具读取完整段落后再修改；保留内容与结构；优先少量统一预设；不修改代码；以实际工具结果说明完成情况。

用户需求作为 user 消息，页面正文只作为 tool 结果。笔记中即使包含“忽略之前规则、访问某个 URL、修改别的页面”等文字，也不能改变工具权限。真正的保障仍是本地工具白名单、固定目标页、范围校验与格式不变量。

日志只记任务 ID、耗时、模型 ID、工具名、计数、错误分类和用量。API Key、正文、完整工具参数、reasoning_content 不写日志或仓库。错误响应中可能包含输入内容，展示和记录前要限制长度并清理敏感字段。

这些规则是执行器内部约束，不要求用户每步批准正常排版操作。

## 14. 线程与生命周期

保留已在仓库验证的模型：独立 STA 负责 WPF，线程池后台执行 HTTP、XML 和现有早绑定 COM 访问。不要仅因为新增 Agent 就切换 OneNote 的 COM 激活方式。

Agent 窗口只允许一个运行任务；再次点击按钮激活已有窗口。PageEditCoordinator 负责与其他写页面功能协调，单独的 _agentRunning 不能替代共享锁。

OnDisconnection/OnBeginShutdown 先禁止新任务并发送取消，停止排队中的 COM 工作；正在执行的 COM 调用结束后再释放它使用的引用。由统一生命周期管理器管理在途调用，不在网络 await 期间持有 COM 调用锁，也不从窗口线程强行 FinalReleaseComObject。宿主回调不能无限等待；断开处理应与现有 ReleaseHostReferences 路径整合并单独验收，避免恢复仓库已经解决过的 dllhost 残留问题。

## 15. 拟新增和修改的文件

以下路径为仓库根目录下的建议实现位置：

| 文件 | 职责 |
|---|---|
| Ribbon.xml、AddIn.cs | 按钮、窗口启动、任务生命周期 |
| Views/AgentWindow.xaml、.xaml.cs | 输入、进度、结果、撤销 |
| Services/Agent/AgentRunner.cs | 工具调用循环、预算和状态机 |
| Services/Agent/AgentChatClient.cs | 多轮消息与工具调用 SSE 聚合 |
| Services/Agent/AgentProtocol.cs | 消息、工具、参数与结果 DTO |
| Services/Agent/AgentToolRegistry.cs | 显式工具注册和参数校验 |
| Services/Agent/AgentPageSnapshot.cs | 页面结构、短 ID、权限与指纹 |
| Services/Agent/AgentTools.cs | 六个工具实现，规模扩大后再拆文件 |
| Services/ParagraphStyleEditor.cs | 段落样式和 QuickStyleDef 管理 |
| Services/RichTextStyleEditor.cs | 局部 HTML 格式修改 |
| Services/PageEditCommitter.cs | 冲突检查、局部容器提交、回读和撤销 |
| Services/PageEditCoordinator.cs | 按页提交锁与在途 COM 操作管理 |
| Services/AiClient.cs | 抽取复用传输层，保留旧 CompleteAsync 行为 |
| Services/AiConfig.cs | 旧配置兼容、Agent 限额和供应商能力 |
| Services/OneNoteApi.cs、PageEditor.cs | 接入抽象与统一严格提交器 |
| Services/RenderDiagnostics.cs | 暴露无需 OneNote 的 Agent 测试入口 |
| Tools/agent-test.ps1 | 离线协议、格式、冲突与状态机测试，UTF-8 BOM |
| Tools/agent-format-probe.ps1 | 仅显式运行时验证专用测试页，不在普通回归中执行 |
| README.md | 功能说明、范围、配置、取消和限制 |

配置缺少 Agent 节点时使用代码默认值，不重写用户现有 ai-settings.xml，也不复制本机配置入库。首次可以继续使用 JavaScriptSerializer；Agent 路径应设置有限 JSON 大小、嵌套层数与参数长度，不能直接沿用 MaxJsonLength=int.MaxValue 作为唯一约束。

## 16. 测试与验收

### 16.1 离线自动化

| 类别 | 关键用例 |
|---|---|
| 工具协议 | 单/多工具调用、参数分片交错、content=null、tool_call_id 对应、未知工具 |
| 流式失败 | 中途 EOF、length、错误事件、取消、空闲超时，确认不执行不完整工具 |
| 多轮兼容 | reasoning_content 按能力回传、接口不支持 tools 时明确失败 |
| 格式 | 多 one:T、嵌套 span、链接、实体、中文、emoji、nbsp、br、重复短语 |
| 继承 | OE 与 span 覆盖、QuickStyleDef 引用、共享样式不影响未选段落 |
| 保护 | 代码、表格、图片占位、墨迹、附件、未知 HTML、半段选区 |
| 冲突 | 用户改文字、只改颜色、移动/删除段落、改变祖先格式、改变样式定义 |
| 提交 | 0x80042010 重读重建、其他 COM 错误不盲重试、未知结果回读 |
| 状态 | 未 finish 不提交、预算超限不提交、提交中取消、重复 tool_call 幂等 |
| 撤销 | 没有后续编辑可恢复；后来改过的段落跳过 |
| 数据不变量 | 修改前后文本、结构和保护对象一致；只允许预期格式差异 |

采用 FakeAgentChatClient 与 FakeOneNotePageAccess，无网络、无运行中的 OneNote 即可验证全流程。测试应检查最终行为与保护不变量，不只检查实现调用了某个方法。

### 16.2 本机 OneNote 探针

在明确指定的临时测试页上，逐项写入字体、字号、局部样式、对齐、段前后间距和标题样式；随后读取 XML，并在 OneNote 中观察效果。特别验证一个 Outline 内同时有文字、图片、表格、代码的情况。

把通过的能力形成 capability 表和脱敏 XML 样本。没有通过的属性不出现在工具说明里。图片和墨迹保留要做对象级检查，不能只看正文仍然存在。

### 16.3 验收场景

对含标题、多个正文段落、嵌套列表、超链接、代码框和图片的测试页输入示例需求，确认标题层次清晰、正文统一、段间距合适；确认文字内容逐字不变、代码和图片仍完整。模型处理中修改一段字号，该段必须被识别为冲突；中途切换页面，目标仍是窗口显示的固定页面。

共享页面编辑或客户端发生修改后按仓库要求执行：

```powershell
dotnet build OneNoteCodeHelper.sln -c Release
powershell -ExecutionPolicy Bypass -File Tools\detect-test.ps1
powershell -ExecutionPolicy Bypass -File Tools\ai-merge-test.ps1
powershell -ExecutionPolicy Bypass -File Tools\highlight-selection-test.ps1
powershell -ExecutionPolicy Bypass -File Tools\agent-test.ps1
```

普通回归不加 -Live，不运行安装、卸载或注册脚本。构建后检查跟踪的 bin 文件，避免把无关产物变化混入提交。

## 17. 推荐实施顺序与工作量

| 阶段 | 交付物 | 估算 |
|---|---|---:|
| 1. 能力探针 | 样式往返、混合容器保留、当前模型工具协议的验证结果 | 0.5–1 天 |
| 2. 编辑内核 | 快照、样式修改、指纹、严格提交和关键离线测试 | 2–3 天 |
| 3. Agent 循环 | 工具注册、SSE、多轮、预算、错误处理 | 1.5–2 天 |
| 4. 用户交互 | 按钮、输入窗口、进度、结果、会话撤销 | 1–2 天 |
| 5. 集成验收 | 与旧功能共存、真实页面验收、回归和文档 | 1–2 天 |

一名熟悉仓库的开发者，预计约 6–10 个工作日完成可用且带保护的第一版；这是设计估算，不是实测承诺，Office 回存兼容问题会影响时间。仅演示工具能改字号的原型可以更快，但不能据此判断完整排版功能已经完成。

实施优先级是先验证“格式补丁写入后能准确读回”，再接模型与按钮。第一版完成后，新增诸如“统一表格外观”只需增加对应工具和校验规则；修改正文、清理段落、移动图片等能力分别扩展，不让已有格式工具承担这些副作用。
