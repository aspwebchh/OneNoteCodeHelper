# OneNote 代码高亮

OneNote 桌面版的 COM 外接程序，把笔记里的代码渲染成带底色的高亮代码框。
目前支持 **Java**、**C#**、**C/C++**、**JavaScript**、**Python**、**SQL**、**Lua**、**PowerShell**、
**Bat**、**Bash**、**XML**、**HTML**、**CSS**、**JSON**、**YAML**，以及不着色的 **纯文本**（只要等宽字体和代码框，适合放日志、命令输出）。

功能区「开始」选项卡上会多出一个「代码高亮」组：

| 控件 | 作用 |
|---|---|
| 高亮选中 | 选中页面上已有的代码文字，原地替换成高亮代码框 |
| 插入代码 | 打开窗口粘贴代码，预览确认后插入到当前页 |
| 语言 | 自动识别，或手动选上面列的任意一种。纯文本只能手动选，自动识别不会选它；TypeScript、JSX 按 JavaScript 识别 |
| 深色主题 | 在浅色（类 IntelliJ）与深色（类 VS Code Dark+）之间切换 |
| 字体（插入窗口内） | 默认 Consolas。代码里有中文时改选「NSimSun」新宋体，中英文才能对齐 |
| 诊断日志 | 打开日志文件 |

旁边还有一个「AI 助手」组，用大模型处理笔记文字和格式：

| 控件 | 作用 |
|---|---|
| Agent | 打开需求输入窗口，通过模型工具调用整理当前页或选中段落的格式；支持取消、结果核验和本次撤销 |
| AI 优化 | 选中文字后点它，按「功能」里选的方式修改并**直接写回**；什么都不选则处理整页（含标题）。弹一个进度小窗，可以取消 |
| 功能 | 默认有「智能校正」（同时校对和排版）「错别字修复」「排版优化」三项，初次使用选中「智能校正」；选项和提示词都来自配置文件，可以自己加 |
| 模型 | 直接显示发给接口的模型名，默认 `deepseek-v4-flash` / `deepseek-v4-pro`，来自配置文件 |
| 思考 | 思考强度 `none` / `low` / `medium` / `high` / `max`，参数和 opencode 配置里 deepseek 的 variants 一致：`none` 传 `"thinking":{"type":"disabled"}`；其余传 `"thinking":{"type":"enabled"}` 加 `"reasoning_effort":"<变体名>"` |
| AI 配置 | 用系统默认的程序（.xml 关联的编辑器）打开配置文件，保存后自动重新读取，下拉随即刷新 |

「AI 优化」的行为：

- 只改段内文字，不合并、不拆分段落。改动按字符合并回原段落，**加粗、颜色、链接、列表和缩进都保留**；
  改错字时新字沿用被替换字的格式，排版时补的空格不会跟进加粗或链接。
- 代码框（段落字体是 Consolas、新宋体等等宽字体）不会交给 AI。
- 「智能校正」「排版优化」顺带删掉多余的空行：连续好几行空行只留一行，文本框（表格单元格）开头、结尾的空行删掉；
  Shift+回车打出的段内空行同样处理。删空行就是删段落，AI 按约定不能动段落，所以这一步由插件自己做，
  不经过 AI。带项目符号、待办标记的空行和代码框里的空行不算。选中了文字时只删选区里的空行。
- AI 返回什么就写回什么，不按改动大小把关。
- AI 处理期间你还可以继续编辑。写回前会重新读页面，处理期间被你改过的段落直接跳过，不会覆盖。
- 进度窗做完不会自己关：结果里只有 AI 自己写的改动清单（每段附的说明汇总去重，只算真正写回了的段落；
  插件删空行不在里面），看完点右上角的叉（或按 Esc）关。处理中关窗等于取消。
- 请求走流式，进度窗实时显示「已思考 / 已输出多少字」和已用时间。连续 60 秒一点数据都没收到
  （接口或网络卡住）会直接报错，不会干等到 `TimeoutSeconds`。

### Agent 页面排版

在「AI 配置」中填写接口地址和 Key，选一个支持 Chat Completions `tools` 的模型，然后点击「Agent」。
窗口固定打开时的页面；之后切换 OneNote 页面，Agent 仍处理窗口里显示的目标页。
输入需求并点击「执行」，例如：

> 将该页面上的内容排版下，要美观。统一正文格式，突出标题和重点。

也可以输入「一级标题居中，正文统一 11 磅，关键结论加粗」。模型通过七个受限工具读取段落、调整草稿、提交修改，
界面显示执行进度，最终结果以插件回读 OneNote 后核验的数据为准。

- 范围默认「当前页」，包括原标题。打开窗口前选中文字，可以切换到「选中段落」；这个范围包含选中文字所在的**完整段落**，不只选中的几个字。选区在打开窗口时固定，重新选择需要重开窗口。
- 支持页面标题、一级/二级标题、正文、引用预设；字体、字号、四种主题文字颜色、左/中/右对齐、段前后间距，以及局部加粗、斜体、下划线和文字颜色。
- 保留文字、段落数量、顺序、嵌套关系、列表、待办状态、链接和表格结构。不改写文字、不删空行、不移动图片；这些不属于本版 Agent 的工具范围。
- 唯一会改变段落结构的是代码：排版时遇到还没高亮的代码，Agent 调用「高亮选中」同一套逻辑，把连续的代码段落（含中间空行）换成高亮代码框。
  主题、字体、字号、Tab 宽度、边框取打开窗口时功能区的设置；语言由模型指定，或自动识别。当普通正文输入的代码、整段都是等宽字体但还没放进代码框的代码（比如从 IDE 粘贴来的）都会转换。
  代码段落必须在同一个文本框或表格单元格里、前后连续，第一行不能比后面的行缩进更深，带项目符号或待办标记的段落不转换。
- 表格中可以调整普通文字的格式，锁定列宽保持不变；未锁定列仍由 OneNote 根据字体自动计算宽度。
- 已有的代码框、行内等宽代码所在段落、空段落和无法解析的 HTML 会受保护，代码段落也不会被设置正文样式。含图片的文本框可以整理文字；含墨迹、附件、媒体或未知对象的整个文本框跳过。
- 修改先保存在内存草稿中，模型调用 `finish_edit` 后统一写回。写回前重新读取页面；处理期间用户改动的目标段落会跳过，其他段落继续提交。时间戳持续冲突时停止，不强制覆盖。
- 「取消」或关闭窗口会丢弃未提交草稿。如果 OneNote 已开始写入，则先回读确认实际结果；无法确认时会明确提示检查页面，不把模型的“完成”当作成功。
- 「撤销本次」只在当前窗口中保留最近一次执行的已核验修改。撤销会跳过后来又编辑过的段落。关闭窗口或启动下一次执行后不保留上一次撤销记录。
  代码框也一起撤销：代码框之后没被改过时换回原来的段落（原段落会得到新的 objectID），改过的代码框跳过。
- Agent 使用功能区选中的模型、思考强度以及打开窗口时读取的 AI 配置，不使用「功能」下拉里的文字改写提示词。

工具协议和实现位置见 [Agent 实施说明](docs/agent-implementation.md)，完整设计见 [技术设计](docs/agent-design.md)。

### AI 配置文件

`%APPDATA%\OneNoteCodeHelper\ai-settings.xml`，第一次点「AI 配置」时生成，里面每一项都有注释：

| 字段 | 说明 |
|---|---|
| `ApiUrl` | OpenAI 兼容接口的基础地址；插件会在后面加 `/chat/completions`，已写全该路径的地址也可直接使用 |
| `ApiKey` | 接口的 Key。**默认是空的**，不填点「AI 优化」会提示 |
| `TimeoutSeconds` | 单次请求从发出到收完的总时限，默认 300 秒。另有固定的 60 秒「无数据」判定，不在这里配 |
| `MaxTokens` | 单次请求最多输出多少 token（含思考过程），默认 16384，0 表示用接口默认值 |
| `Models/Model` | 「模型」下拉的选项，`id` 是接口的模型名，下拉里直接显示它 |
| `Functions/Function` | 「功能」下拉的选项，`name` 显示名、`Prompt` 提示词。`removeExtraBlankLines="true"` 表示顺带删多余的空行；不写时和同名的内置功能一致（「智能校正」「排版优化」默认开，旧配置里的「错别字 + 排版」也继续默认开），写 `false` 关掉 |
| `Agent` | 可选的 Agent 配置；旧文件没有此节点也能使用默认值，不会自动重写已有配置 |

提示词只需写清楚要做什么。输入输出的 JSON 格式、只返回改动的段落、每段附一份改动说明、不许合并拆分段落、
代码网址保持原样这些约定由插件自动接在后面（见 `AiOptimizer.Protocol`），改提示词不会把格式弄坏。

这份文件和 `settings.xml` 分开放，是因为 `settings.xml` 在每次切功能区选项时都会被整体重写，
手改的 Key、提示词放在那里会被覆盖。插件只在文件不存在时写一次默认值，之后只读。
所以以后新增的内置功能（比如「智能校正」）不会自动进旧的配置文件：自己在 `Functions` 里加一项，
或者把文件删掉让插件重新生成（Key 要重填）。

Agent 可在 `AiConfig` 根节点内增加下列配置。改完后重开 Agent 窗口生效：

```xml
<Agent>
  <MaxTurns>12</MaxTurns>
  <MaxToolCalls>48</MaxToolCalls>
  <TimeoutSeconds>600</TimeoutSeconds>
  <MaxPageChars>40000</MaxPageChars>
  <MaxRequestChars>120000</MaxRequestChars>
  <FontFamily>Microsoft YaHei</FontFamily>
  <SendThinking>true</SendThinking>
  <ReplayReasoning>true</ReplayReasoning>
  <StreamUsage>true</StreamUsage>
  <EnableNativeHeadings>true</EnableNativeHeadings>
  <EnableParagraphSpacing>true</EnableParagraphSpacing>
  <EnableMixedOutlines>true</EnableMixedOutlines>
  <EnableCodeHighlight>true</EnableCodeHighlight>
</Agent>
```

`EnableCodeHighlight` 设为 `false` 时 Agent 不再把代码转换为代码框，整段等宽的代码只保护、不处理。

`Agent/TimeoutSeconds` 是整个任务的总时限，根节点的 `TimeoutSeconds` 仍是每次 HTTP 请求时限。
`MaxPageChars` 限制处理范围内的可编辑文字（含待转换的等宽代码；放不下时这些代码只保护、不转换）；`MaxRequestChars` 限制每轮包含工具定义和历史的请求体字符数。
超限会停止并丢弃未提交草稿，可以缩小范围再执行。另有固定的 1000 段上限。

默认字体可选 `Microsoft YaHei`、`Calibri`、`Arial`；未安装时选择其中已安装的一种。
三个 `Enable...` 开关默认开启，已有本机 Office16 回存验证；其他 Office 构建如有兼容问题，可分别关闭原生标题、段间距或图文混排支持。
`EnableMixedOutlines` 只允许图片和文字混排，不开启墨迹或附件编辑。
如果兼容接口不接受扩展参数，按其文档关闭 `SendThinking`（不发送 thinking/reasoning_effort）、
`ReplayReasoning`（不回传 reasoning_content）或 `StreamUsage`（不发送 stream_options）。模型本身仍必须支持工具调用。

## 环境要求

- OneNote **桌面版**（Office16 的 `ONENOTE.EXE`）。UWP 版「OneNote for Windows 10」没有 COM 接口，用不了。
- .NET Framework 4.8 运行时（Windows 10/11 自带）。
- 构建需要 .NET SDK 或 VS2022。

## 构建与安装

根目录的 `install.ps1` 一步搞定「关闭 OneNote → 构建 → 写注册表 → 重开 OneNote」：

```
powershell -ExecutionPolicy Bypass -File install.ps1
```

**需要管理员权限**，脚本会自己弹 UAC 提权（原因见下面「踩过的坑」：COM 类必须注册到 HKLM）。
提权后会开一个新的 PowerShell 窗口执行，窗口会留着让你看结果。

改了代码要重装，同一条命令再跑一遍即可。常用参数：

| 参数 | 作用 |
|---|---|
| `-Configuration Debug` | 装 Debug 版（默认 Release） |
| `-SkipBuild` | 跳过构建，直接注册已有输出 |
| `-Uninstall` | 卸载 |
| `-Force` | OneNote 不肯退出时强制结束进程。默认只礼貌请求，失败就停下来，免得丢掉未保存内容 |
| `-NoRestart` | 完成后不自动重开 OneNote |

卸载时运行根目录的 `uninstall.ps1`，会自动提权并关闭 OneNote：

```
powershell -ExecutionPolicy Bypass -File uninstall.ps1
```

卸载脚本也支持 `-Force` 和 `-NoRestart`。卸载只清理注册表项，保留设置和日志。

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
  AiConfig.cs               ai-settings.xml 的模型与读写，思考强度的五档及其请求参数
  AiClient.cs               Chat Completions 接口（HttpClient + JavaScriptSerializer，流式 SSE）
  AiOptimizer.cs            AI 优化的编排：读段落 → 分批并发问 AI → 写回；固定的输入输出约定
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
                            CommonScanners 放 C 系语言共用的注释、字符串、数字、插值字符串扫描；
                            CFamilyFeatures 放 Java/C#/C++/JS 共有的识别特征
  Themes/CodeTheme.cs, CodeThemes.cs
Views/
  InsertCodeWindow.xaml     插入代码窗口
  CodePreviewRenderer.cs    用同一套 token 流渲染 WPF 预览
  AiProgressWindow.xaml     AI 优化的进度小窗
  AgentWindow.xaml          Agent 需求、范围、进度和撤销窗口
install.ps1                 一键构建 + 安装 / 卸载
uninstall.ps1               一键卸载入口
Tools/register.ps1          只做注册这一步
Tools/unregister.ps1        只做注销这一步
Tools/detect-test.ps1       自动识别回归测试，样本在 Tools/detect-samples/<语言 id>/ 下
Tools/ai-merge-test.ps1     AI 助手回归测试：格式合并、模型输出解析；加 -Live 用本机配置真调一次接口
Tools/highlight-selection-test.ps1  「高亮选中」回归测试：认选区、缩进的段落、换成代码框
Tools/agent-test.ps1        Agent 离线回归，不调用真实接口或 OneNote
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
   - C 系语言先加上 `CFamilyFeatures.Score(sample)`，自己只写独有的特征。共有特征四种语言分数相同、
     互相抵消，胜负才取决于独有写法。
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

- 整段以标签开头、以 `>` 结尾的，只在 XML 和 HTML 之间选，里面的 `<script>` 再像 JS 也不算。
- 高亮效果几乎一样的语言算一族（C 系：Java/C#/C++/JavaScript；标记：XML/HTML）。整族当成一个候选
  跟族外的最高分比，比得过就取族内分最高的那个。族内打平时，认错的代价只是个别关键字颜色不对。

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
- **AI 请求要走流式**。不走流式时整个结果生成完才有第一个字节：思考得久一点进度窗就一动不动，
  接口卡住时只能干等到总超时。而网关（One API）的日志要等请求结束才记，那段时间在后台也查不到这个请求，
  看起来就像「插件根本没发请求」。排查时看 `log.txt` 里每次请求的「首包」时间，
  以及 Clash 之类代理的连接日志里有没有 `dllhost.exe → 接口域名` 的记录。
- **引用 `System.Web.Extensions` 要连 `System.Web` 一起引**。前者引用了后者，SDK 版的 XAML 编译器
  （`MarkupCompilePass1`）解析不到就报 `MC1000: Could not find assembly 'System.Web'`，看起来像 XAML 写错了。
- **`-replace` 的第一个参数是正则**。`$path -replace '\', '/'` 里的单个反斜杠是非法模式，
  会直接抛 `InvalidRegularExpression`。要替换路径分隔符用 `.Replace()`。
