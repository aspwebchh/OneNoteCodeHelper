# OneNote 本地 MCP

## 架构和运行条件

```text
外部 Agent（例如 Codex）
  │ MCP / stdio
  ▼
OneNoteCodeHelper.Mcp.exe（独立进程，.NET 10）
  │ 本机命名管道 / IPC v1
  ▼
OneNoteCodeHelper 插件（net48，现有 dllhost.exe COM 代理）
  │ McpReadService / McpWorkspaceDraft / AgentTools → 提交核验与撤销 → Office PIA
  ▼
OneNote 桌面版 Office16
```

真实读取、编辑、冲突检查、写回核验和撤销都在插件中执行。MCP EXE 不引用插件 DLL、不创建 OneNote COM 对象，也不启动内置 Agent 模型循环。外部调用不要求插件配置 AI API Key。

先打开 OneNote 并加载插件，随后由 MCP 客户端启动 EXE。插件连接宿主后自动在后台监听，无须打开 Agent 窗口。MCP 服务版本为 1.1.0，IPC 帧仍为 v1。支持已打开笔记本的浏览、搜索、独立读取、编辑、新建、复制、移动与导出；仅支持 Windows 本机，不提供 HTTP、向量搜索、附件导入、Resources 或 Prompts。

管道名称是 `OneNoteCodeHelper.Mcp.<Windows 用户 SID>.<登录会话编号>`。ACL 只允许当前用户，并拒绝网络身份；原生管道使用 `PIPE_REJECT_REMOTE_CLIENTS`。MCP 客户端和 OneNote 必须使用同一用户、同一登录会话。客户端恢复凭据只存在于 MCP 进程内存，不能使用另一个 MCP 进程认领快照。

每个运行中的 MCP 进程占用一条管道连接。插件同时最多保持 64 条连接；满额时新的 MCP 进程等待空位，3 秒内连不上就报 `plugin_unavailable`，已有连接和草稿不受影响，有客户端退出后即可连接。连上后 3 秒内没有完成握手的连接会被断开，不占名额。

## 构建和发布

开发机需要 .NET 10 SDK，以及原插件构建所需的 .NET Framework 4.8 参考程序集。官方 C# MCP SDK 固定为 `ModelContextProtocol 2.2.0`，宿主依赖固定为 `Microsoft.Extensions.Hosting 10.0.0`。

在仓库根目录执行：

```powershell
powershell -ExecutionPolicy Bypass -File publish.ps1
```

此命令构建完整解决方案，并在 `bin\publish\` 生成包含插件、MCP 自包含运行时、安装脚本和文档的发布目录及 ZIP。支持 Windows PowerShell 5.1 和 Windows 上的 PowerShell 7.x。
每次输出独立目录；目录内有 `INSTALL.md`、版本及入口清单 `release-manifest.json` 和 `SHA256SUMS.txt`，ZIP 另有 `.sha256` 校验文件。
可用 `-OutputDirectory D:\Releases` 更改发布包父目录，`-NoZip` 省略压缩；`-SkipBuild` 使用已有插件构建和 MCP 自包含输出，不调用 `dotnet`。
默认 DLL 被正在运行的插件占用时，可以构建到独立目录并打包已验证的输出，无须关闭 OneNote：

```powershell
dotnet build OneNoteCodeHelper.sln -c Release -p:OutputPath=bin\mcp-extension\
powershell -ExecutionPolicy Bypass -File Tools\publish-mcp.ps1 -OutputDirectory bin\mcp-extension-publish
powershell -ExecutionPolicy Bypass -File publish.ps1 -SkipBuild -PluginDirectory bin\mcp-extension -McpDirectory bin\mcp-extension-publish
```

发布包解压到固定目录后，在包根目录运行：

```powershell
powershell -ExecutionPolicy Bypass -File install.ps1 -Configuration Release -SkipBuild
```

包内没有源码，安装必须加 `-SkipBuild`。目标机需要 OneNote 和 .NET Framework 4.8，无须 .NET 10 SDK 或运行时。
安装后使用包内的 `Mcp\Server\bin\Release\net10.0-windows\win-x64\publish\OneNoteCodeHelper.Mcp.exe`，把下文 Codex 配置示例中的路径改成解压目录的实际绝对路径。

只构建插件并单独发布 MCP 时执行：

```powershell
dotnet build OneNoteCodeHelper.sln -c Release
powershell -ExecutionPolicy Bypass -File Tools\publish-mcp.ps1
```

等价的发布命令：

```powershell
dotnet publish Mcp\Server\OneNoteCodeHelper.Mcp.csproj -c Release -r win-x64 --self-contained true -p:PublishTrimmed=false -p:PublishSingleFile=false
```

发布入口：

```text
Mcp\Server\bin\Release\net10.0-windows\win-x64\publish\OneNoteCodeHelper.Mcp.exe
```

交付时复制整个 `publish` 文件夹。它自带 Windows x64 运行时，目标机器不需要另装 .NET 10；不能只复制 EXE。插件 DLL 仍位于 `bin\Release\net48\OneNoteCodeHelper.dll`，注册路径和 Office PIA 机制保持原样。发布产物由 `.gitignore` 排除。

`install.ps1` 的正常构建流程先构建插件项目，再发布 MCP，然后沿用原来的插件注册步骤。MCP 是可选组件：缺少 .NET 10 SDK、NuGet 源不可用、或发布目录里的 EXE 正被客户端（如 Codex 会话）占用时，只给警告并跳过 MCP，插件照常注册；退出客户端后重新运行即可更新 MCP。`-SkipMcp` 只构建插件、不发布 MCP；`-SkipBuild` 会同时跳过构建和发布。安装脚本会关闭 OneNote、修改 COM 注册，需由用户在方便时手动运行；上面的构建、发布和离线测试命令不会执行安装。

更新前先停止连接本服务的 MCP 客户端会话，以释放发布目录中正在使用的 EXE 和 DLL；更新后重新启动客户端以刷新工具目录。

## 诊断和启动

```powershell
& '.\Mcp\Server\bin\Release\net10.0-windows\win-x64\publish\OneNoteCodeHelper.Mcp.exe' --doctor
```

成功时输出 JSON，包含 `connected`、`instance_id`、`ipc_version` 和记录数量，退出码为 `0`。启动最多等待管道连接和握手 3 秒。插件未就绪返回 `plugin_unavailable`，协议不一致返回 `protocol_mismatch`，退出码为 `1`；参数错误为 `2`，其他启动或传输错误为 `3`。失败说明写入 stderr。

连接失败时检查：OneNote 是否打开、功能区是否有插件按钮、安装的插件 DLL 和 MCP EXE 是否来自同一次发布、Windows 用户和登录会话是否一致。可查看插件日志中的 `MCP ... state=...`；MCP 不记录笔记正文、工具参数或密钥。日志路径沿用 `%LOCALAPPDATA%\OneNoteCodeHelper\log.txt`。

无参数启动采用 stdio。正常模式的 stdout 只输出 MCP 消息，stdin 关闭后进程退出；无需手工双击并常驻 EXE。`--pipe-name <name>` 用于隔离管道测试和诊断，通常不要在客户端配置中使用。

## Codex 配置示例

以下仅为示例，不会自动修改用户配置。将 `command` 改成实际发布入口的绝对路径，加入 `%USERPROFILE%\.codex\config.toml`：

```toml
[mcp_servers.onenote_local]
command = 'G:\WPF\OneNoteCodeHelper\Mcp\Server\bin\Release\net10.0-windows\win-x64\publish\OneNoteCodeHelper.Mcp.exe'
startup_timeout_sec = 10
tool_timeout_sec = 120
```

也可以使用 Codex CLI 注册 stdio 服务：

```powershell
codex mcp add onenote_local -- 'G:\WPF\OneNoteCodeHelper\Mcp\Server\bin\Release\net10.0-windows\win-x64\publish\OneNoteCodeHelper.Mcp.exe'
```

配置字段和 CLI 用法见 [Codex MCP 官方文档](https://learn.chatgpt.com/docs/extend/mcp?surface=cli)。先运行 `--doctor` 成功，再让 Codex 连接。其他支持 stdio 的本机 MCP 客户端使用相同 EXE 即可。

## 工具和完整调用流程

目录包含原有 23 个 Agent 工具、6 个管理工具和 16 个新增 MCP 工具，共 45 个。新增工具只在外部 MCP 可用，内置 Agent 目录及模型行为保持原样。能力开关、选区保护、代码及图片保护和插入配额仍由插件检查。目录里的工具不一定在当前草稿可用，`begin_edit.available_tools` 和页面概况返回实际能力；使用被关闭的能力会明确报错。

| 管理工具 | 作用 |
|---|---|
| `get_status` | 连接状态、插件实例、IPC 版本、记录数量 |
| `get_current_page` | 当前页 ID、标题、选中段落数 |
| `begin_edit` | `scope=page` 或 `selection`；可指定 `page_id`，省略保留当前页行为。选区只允许当前页 |
| `get_edit_status` | 草稿状态、修订号、调用次数、提交及撤销结果 |
| `abort_edit` | 取消未提交草稿；已经开始写入则继续核验 |
| `undo_edit` | 撤销本客户端该快照的已核验修改，检查冲突并保留整组撤销 |

现有工具：`get_page_overview`、`read_blocks`、`read_image_text`、`set_page_title`、`set_paragraph_style`、`set_text_style`、`fix_text`、`strip_markdown`、`clear_format`、`set_list`、`set_tag`、`set_table_style`、`highlight_code`、`normalize_code_spacing`、`remove_blank_lines`、`set_indent`、`move_blocks`、`merge_outlines`、`insert_blocks`、`text_to_table`、`unwrap_code`、`get_pending_changes`、`finish_edit`。

原工具的参数和结果沿用内置 Agent。MCP 的 `get_page_overview` 额外要求 `snapshot_id`；所有草稿工具都必须绑定同一快照。同一快照的调用串行执行，独立快照可并发，但页面提交和内置功能共用页面锁。

示例：把合成页面的第一段设为一级标题，其他段落不动。

1. 调用 `get_current_page {}`，确认目标页。
2. 调用 `begin_edit {"scope":"page"}`，保存返回的 `snapshot_id`（以下写作 `S`）。选中范围改用 `selection`，和 Agent 窗口一样包括夹在选中段落之间的空行；没有选中段落会报 `selection_empty`。
3. 调用 `get_page_overview {"snapshot_id":"S","offset":0}`，按 `next_offset` 翻页到末尾。页面和选区从开始时固定，切换页面或重新选择不会改变这个快照。
4. 调用 `read_blocks {"snapshot_id":"S","block_ids":["p1","p2"]}`，每批最多 100 段，读完范围内所有可读段落。ID 来自概况和 `next_read_block_ids`，不能猜。保护段落会返回跳过原因。
5. 调用 `set_paragraph_style {"snapshot_id":"S","block_ids":["p1"],"preset_id":"heading1"}`。此时只改草稿。
6. 调用 `get_pending_changes {"snapshot_id":"S"}`，确认改动、`unread_count=0`，读取最新 `draft_revision`（以下写作 `R`）。
7. 单独调用 `finish_edit {"snapshot_id":"S","draft_revision":R}`。未读完返回 `ok=false`，旧修订号被拒绝；两种情况都未提交。
8. 以结果的 `status`、冲突及未核验数量判断完成情况。需要撤销时调用 `undo_edit {"snapshot_id":"S"}`。

## 日常扩展接口

| 工具 | 输入与行为 |
|---|---|
| `list_nodes` | `parent_id?`, `cursor?`, `limit?`。逐层返回 `id/type/name/parent_id/path/last_modified`，不枚举回收站 |
| `search_pages` | `query`, `root_id?`, `cursor?`, `limit?`。原生 FindPages 搜索，不改变界面；返回路径、摘要与链接 |
| `read_page` | `page_id`, `format?`, `cursor?`, `limit?`。默认 `blocks`，可用 `markdown`；无需编辑草稿 |
| `read_selection` | `format?`, `cursor?`, `limit?`。固定当前页完整选中段落，无选区报错 |
| `find_tasks` | `root_id?`, `completed?`, `cursor?`, `limit?`。默认未完成；按页可取消扫描，返回来源、对象 ID、链接和 `skipped` |
| `append_content` | `snapshot_id`, `content`, `format`。整页草稿下方新增独立文本框，可向空页追加，不移动原有对象 |
| `insert_content` | `snapshot_id`, `target_id`, `position=before/after`, `content`, `format`。目标必须已读取并位于可调整结构的范围 |
| `replace_text` | `snapshot_id`, `block_id`, `quote`, `occurrence`, `replacement`。区分大小写精确匹配，第几处从 1 开始；只改普通正文 |
| `begin_workspace_edit` | 无参数。建立一项工作区操作草稿，返回 `snapshot_id` 与修订号 |
| `create_page` | `snapshot_id`, `section_id`, `title`, `content?`, `format?`。正文可省略，格式默认 `plain` |
| `copy_page` / `move_page` | `snapshot_id`, `page_id`, `section_id`。单页复制／移动；移动拒绝含子页的父页面 |
| `create_section` | `snapshot_id`, `parent_id`, `name`。父节点为笔记本或分区组，提交返回真实分区 ID |
| `get_link` / `navigate_to` | `page_id`, `object_id?`。使用真实层级或对象 ID；定位会改变当前视图 |
| `export_page` | `page_id`, `format=markdown/pdf`, `output_path`。绝对路径、目录须已存在、不覆盖文件 |

列表默认 50 项、最多 100 项。只读结果不占编辑草稿名额，最多同时保留 64 个，5 分钟后游标报 `cursor_expired`，需要重新读取。下一页应使用相同查询参数与返回的 `next_cursor`，可以调整 `limit`。页面读取固定完整 XML，后续编辑不会混入旧读取；搜索固定候选 ID、路径及版本，只为本批生成摘要。候选页后来变化、无法读取或没有可核验版本时，分别返回 `page_changed`、`page_unreadable` 或 `version_unavailable`，不会返回新正文冒充旧摘要。

统一读取结构含 `type/object_id/parent_id/container_id/depth/text/markdown/readable/editable`，以及链接、列表和标记。表格含 `rows` 和逐单元格的 `cells`，代码框可读但写入保护；图片返回 OCR 与占位，未知对象返回原因。顶层 `complete/issues` 明示完整性，不将占位内容称为保真导出。真实 `object_id` 与编辑快照的 `p1/n1/t1` 是不同的身份；编辑须重新 `begin_edit → get_page_overview → read_blocks`，按概况中的真实对象 ID 对应短 ID。

`plain` 保留空行及行首缩进。受限 Markdown 支持一、二级标题、引用、嵌套项目符号与编号列表、待办、粗斜体、链接、围栏代码和管道表格；代码复用插件语言、主题、字体和 HTML 编码。其他语法保留为文字并返回 `warnings`，原始 HTML 不执行、图片不下载。每个草稿插入上限 5000 字、50 段，表格单元格和代码行也计入段数。OneNote 不保留全空文本框，追加全空内容明确拒绝；空正文新建页面省略文本框。

`replace_text` 不跨换行，不修改标题、代码、混合二进制对象或不能无损解析的 HTML。保留段落数量、层级以及未修改字符的格式和链接，提交仍核对原文；插入和追加通过结构核验及整框撤销。指定页整页编辑也必须读完可读取范围。

工作区草稿最多规划一项操作，相同参数重复规划返回原计划，参数不同须另开草稿；同一修订号重复提交只返回原结果。创建／复制沿用 `EnableInsert`，移动还要求 `EnableMoves`，导入列表、标记、代码和表格服从各自开关。新页面及副本仅在内容和位置未被后续修改时可撤销，新分区仅在仍为空且名称、位置未变时可撤销。

复制取得图片二进制数据，核验文字、格式、层级、表格、标记、链接和图片；墨迹、附件、音视频、未知对象或坏图在写入前拒绝。新页面和副本都以 OneNote 新建的空白页为底写入，沿用其样式编号和原生标题；副本带上源页的页面设置和日期，子页面复制或移动后成为一级页面。写入核验按 OneNote 的回存规则归一：作者与语言、页面日期与层级、未由用户设定的文本框宽度、OCR 及数值写法由 OneNote 维护，不参与比较。移动依次创建目标副本、完整核验、再次检查源页、将源页放入回收站；返回 `page_id/link/object_id_map/source_recycled/source_outcome_known`，不保证旧链接有效。源页变化或删除失败时保留源页及已核验目标副本并报告实际结果；结果未知禁止自动重发。撤销成功移动先重建并核验源分区副本，再移除未被编辑的目标副本，返回恢复页的新 ID 和链接。源页状态无法确认时不允许补偿撤销。

删除 API 仅使用回收站模式，提交前和实际删除前都检查笔记本回收站；缺少可确认回收站的笔记本明确拒绝可撤销工作区操作。新建的笔记本在第一次删除内容之前，层级里没有回收站分区组（16.0.20326 实测），此时先在其中删除一页即可。旧 OneNote 2007 格式即使使用非永久删除参数也会永久删除，因此不支持这类工作区写入，依据 [Microsoft OneNote Application 接口说明](https://learn.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote#deletehierarchy-method)。

导出只读取已保存页：Markdown 使用统一模型，返回占位或未解析内容的 `complete/issues`；PDF 使用原生 Publish。先生成同目录临时文件、检查源页版本和文件内容，再移动到目标路径；失败清理临时文件且不覆盖已有目标。返回字节数、绝对路径和源页修改时间。新增工具提供输入／输出 Schema、工具注解；请求附 `_meta.progressToken` 时，stdio 每两秒发送进度通知，不改变 IPC 帧协议。

## 完整调用示例

以下 `P/SECTION/ROOT/S/R` 是前一步工具返回的真实值或草稿值，不要原样照抄；`finish_edit` 必须使用最后一次 `get_pending_changes` 返回的修订号。

**搜索问答**：

```text
list_nodes {}
list_nodes {"parent_id":"ROOT"}
search_pages {"query":"预算 AND 项目","root_id":"ROOT","limit":50}
search_pages {"query":"预算 AND 项目","root_id":"ROOT","cursor":"返回的 next_cursor"}
read_page {"page_id":"P","format":"blocks","limit":50}
read_page {"page_id":"P","format":"blocks","cursor":"返回的 next_cursor"}
get_link {"page_id":"P","object_id":"引用段落的真实 object_id"}
```

读到 `next_cursor=null` 后，根据已读原文回答并附原文链接；遇到不完整对象明确说明，不需要建立编辑草稿。

**会议记录新建和后续追加**：

```text
list_nodes {"parent_id":"ROOT"}
begin_workspace_edit {}
create_page {"snapshot_id":"S","section_id":"SECTION","title":"2026-10-03 项目例会","format":"markdown","content":"# 结论\n确定试点范围。\n## 行动\n- [ ] 整理试点名单\n  - 核对覆盖部门"}
get_pending_changes {"snapshot_id":"S"}
finish_edit {"snapshot_id":"S","draft_revision":R}
begin_edit {"scope":"page","page_id":"返回的新 page_id"}
get_page_overview {"snapshot_id":"新 S","offset":0}
read_blocks {"snapshot_id":"新 S","block_ids":["按概况分批读取所有可读短 ID"]}
append_content {"snapshot_id":"新 S","format":"plain","content":"会后补充\n  试点反馈下周汇总。"}
get_pending_changes {"snapshot_id":"新 S"}
finish_edit {"snapshot_id":"新 S","draft_revision":R}
```

检查概况的 `next_offset` 和待提交结果的 `next_read_block_ids`，翻页并读完后再提交。需要更新正文时，在已读取的普通正文上调用 `replace_text {"snapshot_id":"S","block_id":"p2","quote":"下周","occurrence":1,"replacement":"下周三"}`，然后再次检查与提交；在指定段落附近导入用 `insert_content`。

**待办汇总与完成**：

```text
find_tasks {"root_id":"ROOT","completed":false,"limit":50}
find_tasks {"root_id":"ROOT","completed":false,"cursor":"返回的 next_cursor"}
begin_edit {"scope":"page","page_id":"待办来源 page_id"}
get_page_overview {"snapshot_id":"S","offset":0}
read_blocks {"snapshot_id":"S","block_ids":["范围内所有可读短 ID"]}
set_tag {"snapshot_id":"S","block_ids":["对应待办真实 object_id 的短 ID"],"tag":"todo","completed":true}
get_pending_changes {"snapshot_id":"S"}
finish_edit {"snapshot_id":"S","draft_revision":R}
```

汇总按来源路径显示事项并附链接，检查 `skipped`；本轮没有截止日期或负责人字段。每个来源页面单独建立草稿；读取调用按 100 段分批。

**归档**：

```text
begin_workspace_edit {}
create_section {"snapshot_id":"S","parent_id":"ROOT","name":"2026 归档"}
get_pending_changes {"snapshot_id":"S"}
finish_edit {"snapshot_id":"S","draft_revision":R}
begin_workspace_edit {}
move_page {"snapshot_id":"新 S","page_id":"P","section_id":"刚返回的真实 section_id"}
get_pending_changes {"snapshot_id":"新 S"}
finish_edit {"snapshot_id":"新 S","draft_revision":R}
get_edit_status {"snapshot_id":"新 S"}
undo_edit {"snapshot_id":"新 S"}
```

希望保留源页时将 `move_page` 换成 `copy_page`。归档后使用结果的新链接；检查移动结果后才决定撤销，撤销记录受 30 分钟／16 条限制。

**导出**：

```text
read_page {"page_id":"P","format":"markdown"}
export_page {"page_id":"P","format":"markdown","output_path":"D:\\笔记导出\\会议记录.md"}
export_page {"page_id":"P","format":"pdf","output_path":"D:\\笔记导出\\会议记录.pdf"}
```

目标目录须预先存在；读取 Markdown 按游标读完，导出不依赖这个游标，直接重新读取已保存页并返回对应版本。路径已存在时更换文件名，不自动覆盖。

## 状态、取消和重连

草稿遵守开始时 Agent 配置中的 `TimeoutSeconds`、`MaxToolCalls` 等限制。管理工具不消耗草稿调用配额；读取、草稿修改和首次提交消耗配额，重复提交原结果不再次消耗。最多同时保留 8 个未完成草稿。

`finish_edit` 成功结束后冻结草稿。同一快照和同一提交修订号再次调用只返回原结果，不再写入；不同修订号拒绝。`undo_edit` 完成后也只返回原撤销结果，包括发生冲突或写回未确认的结果。一次结果不明确时先查 `get_edit_status`，不能自动重发写入。

常见提交状态为 `Verified`、`PartiallyApplied`、`NoChange` 和 `CommitOutcomeUnknown`。后者表示“写回未确认”，需要人工查看目标页；不能把它视为没有写入。状态查询另含 `Draft`、`Committing`、`CancelledBeforeCommit`、`Undoing` 等生命周期状态。

MCP 请求取消、`abort_edit`、客户端断线或超时会取消未提交草稿。写入已经开始时，提交器仍继续回读、核验实际结果并记录；取消工具请求的响应本身可能已经丢失，此时通过 `get_edit_status` 查询。取消前已完成的修改通过 `undo_edit` 撤销。

连接建立后 EXE 缓存工具目录；命名管道断开后的下一次调用可重新连接。EXE 不会自动重发失败的工具调用。同一个仍在运行的 MCP 进程可恢复原插件实例的已完成记录；进程退出后恢复凭据丢失，新的 MCP 进程不能认领旧快照。

完成结果和撤销记录在插件内存保留 30 分钟，最多 16 条，按最早完成顺序淘汰。插件重启后全部失效，旧快照不可跨实例使用。操作期间用户修改过的段落、文本框或表格仍由现有指纹及期望页面检查保护，冲突不会被覆盖；跨文本框结构修改整组提交、整组撤销。

## 离线回归

```powershell
dotnet build OneNoteCodeHelper.sln -c Release
powershell -ExecutionPolicy Bypass -File Tools\detect-test.ps1
powershell -ExecutionPolicy Bypass -File Tools\ai-merge-test.ps1
powershell -ExecutionPolicy Bypass -File Tools\highlight-selection-test.ps1
powershell -ExecutionPolicy Bypass -File Tools\agent-test.ps1
powershell -ExecutionPolicy Bypass -File Tools\addin-surrogate-test.ps1
powershell -ExecutionPolicy Bypass -File Tools\mcp-test.ps1
powershell -ExecutionPolicy Bypass -File Tools\publish-mcp.ps1
powershell -ExecutionPolicy Bypass -File Tools\mcp-test.ps1 -McpExe '.\Mcp\Server\bin\Release\net10.0-windows\win-x64\publish\OneNoteCodeHelper.Mcp.exe'
```

MCP 测试使用模拟页面和随机命名管道，启动真实 EXE 验证初始化、目录、调用、中文、错误、取消、重连和 EOF 退出。普通测试不调用真实 AI、不创建真实页面、不运行注册或安装脚本、不关闭用户 OneNote。
页面读取同时覆盖带缩进和无缩进的 XML，验证整页、选区排版与跨文本框合并可以提交、撤销，并保留文字、链接、空白和未选中的嵌套段落；真实的正文、链接、格式及整组结构冲突仍会阻止覆盖。
`Tools\mcp-test.ps1` 可用 `-DllPath` 和 `-McpExe` 指定独立输出目录中的插件 DLL 与 MCP EXE，便于运行中的插件占用默认 DLL 时验证新构建。
扩展回归另覆盖嵌套路径、中文搜索与摘要按批读取、跨客户端游标及过期、独立读取超过编辑段数配额、表格内对象完整性、空页追加、Markdown／emoji、富文本替换、图片复制、源页变化、移动补偿撤销、丢失提交响应、在途取消、导出失败及真实 stdio 进度通知。

## 专用合成页面的真实 OneNote 验收

以下为需要用户手动执行的真实环境验收，离线测试不能代替它。安装更新的插件后，在专门的测试分区手工建立两页，页面标题写明“MCP 验收测试”，只填写虚构内容：几段中文、空行、一段待高亮代码，以及第二个独立文本框。

**真实 COM 往返的自动探针**：`Tests\bin\Release\net48\OneNoteCodeHelper.AgentTests.exe --probe-workspace <目录>` 在该目录新建独立测试笔记本，核对创建分区、带 Markdown 正文的新页面、含图片页的复制、移动／回收站／补偿撤销、页面草稿的 `append_content`/`insert_content`、含待办文本框的内置格式与 AI 写回，以及 Markdown 导出；结束后关闭测试笔记本，不读写其他笔记本。核验失败时打印第一处差异。目录层级与 FindPages、OCR、原生 PDF Publish 仍需下面的专用合成数据验证。测试通过不代表已在用户真实笔记上验收，不自动安装、注册或关闭 OneNote。

1. OneNote 打开并加载插件，运行发布 EXE 的 `--doctor`，应返回版本和连接成功。
2. 让外部 Agent 按完整流程读取并设置第一段样式，观察提交前真实页面不变，`finish_edit` 后样式变化、中文和链接保留。再以相同快照和修订号提交，页面不重复改动。
3. 在第一页选择一段，开始 `selection` 草稿；随后选择其他段落并切换到第二页。概况和提交应仍绑定开始时第一页的完整选中段落。没有选区时 `begin_edit` 应报错。
4. 开始新草稿、读取并改样式后，手工修改同一目标段落文字，再提交。应报告冲突、保留手工修改。跨文本框移动或合并时，在组内任一框手工改动，检查整组跳过。
5. 提交一次合并或跨框移动，调用 `undo_edit`，检查整组恢复；再次撤销只返回原结果。另一轮提交后手工修改组内内容，检查撤销不会覆盖这些修改。
6. 提交前用 `abort_edit` 或退出 MCP 客户端，确认草稿不写回。写入已经开始时取消请求，随后先查询状态；如出现“写回未确认”，人工检查页面，不重发提交。
7. 保持 MCP 客户端开启，在保存测试页后正常关闭 OneNote。检查不再接受新请求，MCP 调用报告连接中断；现有插件代理应按原机制释放 COM 引用，不阻塞宿主关闭。重新打开 OneNote，确认旧快照失效，新 `begin_edit` 正常。
8. 逐层列出测试笔记本和嵌套分区；创建两张同名测试页，用中文搜索并分页，确认路径与真实 ID 可区分，搜索不改变当前视图。扫描待办，锁定一页后检查未读取原因。
9. 用工作区草稿创建分区，再用返回 ID 创建仅标题页与 Markdown 页。正文包含标题、引用、嵌套列表、待办、链接、表格、emoji、带缩进及首尾空行的围栏代码；读取及原生界面核对，追加到空页并撤销。
10. 给测试页加入一张图片并等待 OCR，再复制到测试归档分区，核对图片数据和文字格式；含附件或墨迹的测试页应在提交前拒绝。移动普通页后核对新链接与源页回收站状态，撤销后核对恢复副本。移动带子页面的父页应拒绝。
11. 移动规划后改源页，以及提交后改目标副本，再测试提交／撤销冲突，确认不会覆盖用户改动。用中文绝对路径导出 Markdown 与 PDF，核对占位完整性、PDF 内容以及已存在文件不会被覆盖。

验收完成后仅删除专用测试页。不要在真实笔记上执行破坏性验收。
