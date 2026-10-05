# 仓库工作说明

本文件适用于整个仓库。项目是 OneNote 桌面版（Office16）的 C# COM 外接程序，目标框架为 .NET Framework 4.8，使用 WPF 窗口；另有一个 .NET 10 的本地 MCP 服务端，供外部 Agent 经命名管道调用插件。

动手前先查对应文档：

- `README.md`：用户功能、AI 配置文件、安装与发布、代码结构，以及「几个踩过的坑」里已经实测过的 OneNote、COM 和 PowerShell 限制。
- `docs/agent-implementation.md`：Agent 的执行链路、工具协议、限额、写回保护和代码位置；`docs/agent-design.md` 是最初的完整设计。
- `docs/onenote-mcp.md`：MCP 架构、IPC、工具、离线回归和真实 OneNote 验收步骤。

## 代码入口

- `AddIn.cs` 和 `Ribbon.xml`：COM 外接程序生命周期、功能区及其回调、窗口线程和以 OneNote 为属主的提示框。
- `Interop/`：Office COM 接口与 Win32 声明。`Services/OneNoteApi.cs` 封装 OneNote 调用（串行、断开协调），`Services/PageEditor.cs` 处理页面 XML、选区和写回，`Services/PageEditCoordinator.cs` 按页面串行化所有写回。
- `Highlighting/`：语言词法分析、自动识别和主题；`Services/CodeBlockBuilder.cs`、`OneNoteHtmlEncoder.cs` 生成 OneNote 代码框；`Views/` 提供插入、预览、Agent 和 AI 配置窗口。
- 文字功能：`Services/AiConfig.cs`（`ai-settings.xml` 读写）、`AiClient.cs`（流式请求）、`AiOptimizer.cs`（编排与固定协议）、`RichParagraph.cs` 和 `TextDiff.cs`（保留格式的逐字合并）、`BlankLines.cs`（删多余空行）、`LiveText.cs`（思考摘录）。
- `Services/Markdown/`：「插入代码」窗口的「Markdown」选项（解析、行内格式、生成段落、插到页面末尾），预览在 `Views/MarkdownPreviewRenderer.cs`。段落预设、列表和标记复用 `ParagraphStyles`、`AgentMarks`、`AgentLayout`，强调配对复用 `AgentMarkdown`。Markdown 不是 `LanguageRegistry` 里的语言，不进功能区语言下拉、Agent 和 MCP 的语言列表。
- `Services/Agent/`：Agent 的模型循环（`AgentRunner`）、工具注册与校验（`AgentTools`）、页面快照（`AgentPageSnapshot`）、格式与结构草稿、提交核验和撤销（`AgentCommitter`）。各文件职责见 `docs/agent-implementation.md` 的「代码位置」。
- MCP：`Services/Mcp/` 是插件内的管道服务、读取和工作区草稿；`Mcp/Server/` 是独立的 net10.0-windows EXE；`Mcp/Shared/` 的协议代码同时编进插件和 EXE。
- `Services/AddInSettings.cs`、`AddInLog.cs`：功能区设置和文件日志。`Services/RenderDiagnostics.cs`：不经过 OneNote 跑通渲染链路的入口，供回归脚本反射调用。
- `Tests/`：独立签名的 net48 测试程序（不在解决方案里，由回归脚本单独构建），引用已构建的插件 DLL，靠 `Properties/AgentTestAccess.cs` 的 `InternalsVisibleTo` 访问内部类型。
- `Tools/`：不依赖正在运行的 OneNote 的回归脚本、真实 OneNote 探针，以及需要管理员权限的注册脚本。

## 修改时保持的约束

### 构建与依赖

- 保持 `OneNoteCodeHelper.csproj` 的 `net48` 目标框架、强名称签名，以及 `lib/` 中 Office PIA 的早绑定引用。此项目不能依赖 `dynamic` 或运行时类型库查找来调用 OneNote。
- 插件只使用 .NET Framework 自带的程序集，不引入 NuGet 包（JSON 用 `JavaScriptSerializer`，HTTP 用 `HttpClient`）。MCP 项目的包版本固定，不顺手升级。
- 不要更换 `OneNoteCodeHelper.snk`：COM 注册和 `InternalsVisibleTo` 的公钥都依赖它。
- `Mcp/Shared/` 同时以 net48 和 net10 编译，只能用两边都有的 API，不引用 WPF、Office、MCP SDK 或 JSON 库。主项目通过 `Compile Remove` 排除 `Mcp\**` 和 `Tests\**`，新增文件时检查它落在哪个项目里。

### 页面写回

- 生成 `one:T` 内容时按 HTML 转义并保留行首缩进、空行；改动页面 XML 时检查选区、嵌套段落和原有格式。
- 文字功能的 AI 写回不能合并或拆分段落。Agent 的结构工具（删空行、缩进、移动、合并文本框、插入、转表格、拆代码框）必须走期望页面核验并能整框撤销，跨文本框的改动整组提交、整组撤销。两者都不能覆盖处理期间用户已经改动的段落。
- 页面写回在 `PageEditCoordinator` 的页面锁内进行，始终带读取时的 `lastModifiedTime` 提交；冲突（`0x80042010`）时重读重建或提示重新执行，不重发旧 XML，也不用 `DateTime.MinValue` 强制写入。提交开始后即使取消，也要先回读确认结果。
- OneNote 回存会改写 XML 写法，核验按语义比较，不依赖字符串完全相同。新发现的回存行为先实测，再补进 README「几个踩过的坑」，注明 OneNote 版本。
- 新增或修改 Agent 工具时，JSON Schema 和本地校验共用 `AgentTools.cs` 里的 `AgentSchema` 定义，系统提示词按实际注册的工具拼接；同步更新 Agent 离线测试、`docs/agent-implementation.md` 和 README 的用户说明。MCP 开放的工具沿用同一套保护范围。

### 高亮

- `ILanguage.Tokenize` 返回的 token 必须按顺序、无重叠、完整覆盖源码（`RenderDiagnostics.VerifyCoverage` 检查这一条）。新语言同时注册到 `LanguageRegistry.All`，按需更新语言族，并在 `Tools/detect-samples/<语言 id>/` 添加识别样本。打分规则见 README「加一种语言」。

### 线程与 COM

- OneNote 的功能区回调不能执行耗时操作或同步弹框；WPF 窗口须在 STA 线程上创建。涉及 COM、线程或注册机制的改动先核对 README「几个踩过的坑」。

### 脚本、日志与隐私

- `.ps1` 文件（包括新建的）保存为 UTF-8 with BOM，并同时兼容 Windows PowerShell 5.1 和 PowerShell 7.x。
- 日志只记录状态、次数、耗时、对象 ID 和异常类型，不写笔记正文、工具参数、模型思考或上游错误响应正文。
- 不要把本机的 API Key、`%APPDATA%\OneNoteCodeHelper\ai-settings.xml`、`settings.xml` 或 `%LOCALAPPDATA%\OneNoteCodeHelper\log.txt` 的内容写入仓库。

### 文档

- 代码注释、界面文字、日志和文档使用中文，沿用现有措辞风格。
- 用户可见的行为变化同步更新 README 对应章节；新增、删除或移动源文件时更新 README「代码结构」；Agent 和 MCP 的协议变化分别更新 `docs/agent-implementation.md`、`docs/onenote-mcp.md`。

## 构建与验证

在 Windows PowerShell 中，从仓库根目录执行：

```powershell
dotnet build OneNoteCodeHelper.sln -c Release
powershell -ExecutionPolicy Bypass -File Tools\detect-test.ps1
powershell -ExecutionPolicy Bypass -File Tools\ai-merge-test.ps1
powershell -ExecutionPolicy Bypass -File Tools\highlight-selection-test.ps1
powershell -ExecutionPolicy Bypass -File Tools\agent-test.ps1
powershell -ExecutionPolicy Bypass -File Tools\addin-surrogate-test.ps1
powershell -ExecutionPolicy Bypass -File Tools\mcp-test.ps1
```

构建整个解决方案需要 .NET 10 SDK（含 MCP 项目）。这些回归不需要运行中的 OneNote，也不调用真实接口。按改动范围选择脚本：

| 改动范围 | 至少运行 |
|---|---|
| `Highlighting/`、自动识别、识别样本 | `detect-test.ps1` |
| 「高亮选中」、`PageEditor`、`CodeBlockBuilder`、`OneNoteHtmlEncoder` | `highlight-selection-test.ps1` |
| 文字功能、`AiConfig`、`AiClient`、`RichParagraph`、`BlankLines`、AI 配置窗口 | `ai-merge-test.ps1` |
| `Services/Agent/`、`PageEditCoordinator`、`OneNoteApi`、`Tests/` | `agent-test.ps1` |
| `Services/Markdown/`、插入窗口的 Markdown 模式 | `agent-test.ps1`（用例在 `Tests/MarkdownRegression.cs`）；改了界面再跑 `--render-ui` |
| `Mcp/`、`Services/Mcp/` | `mcp-test.ps1`；改发布方式时再用 `publish-mcp.ps1` 的输出跑一次 `mcp-test.ps1 -McpExe <发布的 EXE>` |
| `install.ps1`、`uninstall.ps1`、`Tools/addin-surrogate.ps1` | `addin-surrogate-test.ps1`，无需构建，分别用 `powershell` 和 `pwsh` 运行 |
| 共享渲染、页面编辑、构建配置 | 全部脚本 |

- 回归脚本默认读取 `bin\Release\net48\OneNoteCodeHelper.dll`，都支持 `-DllPath`。`agent-test.ps1` 和 `mcp-test.ps1` 会先以该 DLL 构建 `Tests/`。
- OneNote 运行时插件 DLL 被 `dllhost.exe` 代理进程占用，构建会报 MSB3021。不要为此关闭 OneNote 或结束代理进程：把插件构建到其他目录（如 `dotnet build OneNoteCodeHelper.csproj -c Release -o <临时目录>`），再用 `-DllPath` 指向它；`mcp-test.ps1` 另有 `-McpExe`。
- 修改 `Views/` 后可运行 `Tests\bin\Release\net48\OneNoteCodeHelper.AgentTests.exe --render-ui <目录>`，离线把各窗口渲染成图片检查布局（先运行 `agent-test.ps1` 生成测试程序）。

## 不在普通验证中运行

- `install.ps1`、`uninstall.ps1`、`Tools/register.ps1`、`Tools/unregister.ps1`：会关闭 OneNote 或修改 COM 注册表。
- 真实 OneNote 探针：`Tools/agent-format-probe.ps1`、`Tools/probe-com.ps1` 和测试程序的 `--probe*` 参数会通过 COM 启动 OneNote，并创建测试分区或笔记本。只在用户明确要求时运行，输出放在专用目录，不碰已有笔记。
- `ai-merge-test.ps1 -Live`：会用本机配置真实调用接口。
- `publish.ps1`：只构建和打包到 `bin\publish\`，不影响 OneNote，但不属于日常验证。

编译产物（`bin/`、`obj/`）、发布包和 PDF 文件由 `.gitignore` 排除，不要重新加入版本控制。
