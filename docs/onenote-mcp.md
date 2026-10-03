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
  │ 现有 AgentTools → AgentCommitter → 页面锁 → Office PIA
  ▼
OneNote 桌面版 Office16
```

真实读取、编辑、冲突检查、写回核验和撤销都在插件中执行。MCP EXE 不引用插件 DLL、不创建 OneNote COM 对象，也不启动内置 Agent 模型循环。外部调用不要求插件配置 AI API Key。

先打开 OneNote 并加载插件，随后由 MCP 客户端启动 EXE。插件连接宿主后自动在后台监听，无须打开 Agent 窗口。v1 仅支持 Windows 本机；没有 HTTP、云端连接、跨笔记本搜索和新建页面工具。

管道名称是 `OneNoteCodeHelper.Mcp.<Windows 用户 SID>.<登录会话编号>`。ACL 只允许当前用户，并拒绝网络身份；原生管道使用 `PIPE_REJECT_REMOTE_CLIENTS`。MCP 客户端和 OneNote 必须使用同一用户、同一登录会话。客户端恢复凭据只存在于 MCP 进程内存，不能使用另一个 MCP 进程认领快照。

## 构建和发布

开发机需要 .NET 10 SDK，以及原插件构建所需的 .NET Framework 4.8 参考程序集。官方 C# MCP SDK 固定为 `ModelContextProtocol 2.2.0`，宿主依赖固定为 `Microsoft.Extensions.Hosting 10.0.0`。

在仓库根目录执行：

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

`install.ps1` 的正常构建流程会构建完整解决方案并发布 MCP，然后沿用原来的插件注册步骤。`-SkipBuild` 会同时跳过构建和发布。安装脚本会关闭 OneNote、修改 COM 注册，需由用户在方便时手动运行；上面的构建、发布和离线测试命令不会执行安装。

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

目录从插件的 `AgentTools` 生成，包含全部现有工具。能力开关、选区保护、代码及图片保护和插入配额仍由插件检查。目录里的工具不一定在当前草稿可用，`begin_edit.available_tools` 和页面概况返回实际能力；使用被关闭的能力会明确报错。

| 管理工具 | 作用 |
|---|---|
| `get_status` | 连接状态、插件实例、IPC 版本、记录数量 |
| `get_current_page` | 当前页 ID、标题、选中段落数 |
| `begin_edit` | `scope=page` 或 `selection`，固定页面、选区和设置；返回快照和配额 |
| `get_edit_status` | 草稿状态、修订号、调用次数、提交及撤销结果 |
| `abort_edit` | 取消未提交草稿；已经开始写入则继续核验 |
| `undo_edit` | 撤销本客户端该快照的已核验修改，检查冲突并保留整组撤销 |

现有工具：`get_page_overview`、`read_blocks`、`read_image_text`、`set_page_title`、`set_paragraph_style`、`set_text_style`、`fix_text`、`strip_markdown`、`clear_format`、`set_list`、`set_tag`、`set_table_style`、`highlight_code`、`normalize_code_spacing`、`remove_blank_lines`、`set_indent`、`move_blocks`、`merge_outlines`、`insert_blocks`、`text_to_table`、`unwrap_code`、`get_pending_changes`、`finish_edit`。

原工具的参数和结果沿用内置 Agent。MCP 的 `get_page_overview` 额外要求 `snapshot_id`；所有草稿工具都必须绑定同一快照。同一快照的调用串行执行，独立快照可并发，但页面提交和内置功能共用页面锁。

示例：把合成页面的第一段设为一级标题，其他段落不动。

1. 调用 `get_current_page {}`，确认目标页。
2. 调用 `begin_edit {"scope":"page"}`，保存返回的 `snapshot_id`（以下写作 `S`）。选中范围改用 `selection`；没有选中段落会报 `selection_empty`。
3. 调用 `get_page_overview {"snapshot_id":"S","offset":0}`，按 `next_offset` 翻页到末尾。页面和选区从开始时固定，切换页面或重新选择不会改变这个快照。
4. 调用 `read_blocks {"snapshot_id":"S","block_ids":["p1","p2"]}`，每批最多 100 段，读完范围内所有可读段落。ID 来自概况和 `next_read_block_ids`，不能猜。保护段落会返回跳过原因。
5. 调用 `set_paragraph_style {"snapshot_id":"S","block_ids":["p1"],"preset_id":"heading1"}`。此时只改草稿。
6. 调用 `get_pending_changes {"snapshot_id":"S"}`，确认改动、`unread_count=0`，读取最新 `draft_revision`（以下写作 `R`）。
7. 单独调用 `finish_edit {"snapshot_id":"S","draft_revision":R}`。未读完返回 `ok=false`，旧修订号被拒绝；两种情况都未提交。
8. 以结果的 `status`、冲突及未核验数量判断完成情况。需要撤销时调用 `undo_edit {"snapshot_id":"S"}`。

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

## 专用合成页面的真实 OneNote 验收

以下为需要用户手动执行的真实环境验收，离线测试不能代替它。安装更新的插件后，在专门的测试分区手工建立两页，页面标题写明“MCP 验收测试”，只填写虚构内容：几段中文、空行、一段待高亮代码，以及第二个独立文本框。

1. OneNote 打开并加载插件，运行发布 EXE 的 `--doctor`，应返回版本和连接成功。
2. 让外部 Agent 按完整流程读取并设置第一段样式，观察提交前真实页面不变，`finish_edit` 后样式变化、中文和链接保留。再以相同快照和修订号提交，页面不重复改动。
3. 在第一页选择一段，开始 `selection` 草稿；随后选择其他段落并切换到第二页。概况和提交应仍绑定开始时第一页的完整选中段落。没有选区时 `begin_edit` 应报错。
4. 开始新草稿、读取并改样式后，手工修改同一目标段落文字，再提交。应报告冲突、保留手工修改。跨文本框移动或合并时，在组内任一框手工改动，检查整组跳过。
5. 提交一次合并或跨框移动，调用 `undo_edit`，检查整组恢复；再次撤销只返回原结果。另一轮提交后手工修改组内内容，检查撤销不会覆盖这些修改。
6. 提交前用 `abort_edit` 或退出 MCP 客户端，确认草稿不写回。写入已经开始时取消请求，随后先查询状态；如出现“写回未确认”，人工检查页面，不重发提交。
7. 保持 MCP 客户端开启，在保存测试页后正常关闭 OneNote。检查不再接受新请求，MCP 调用报告连接中断；现有插件代理应按原机制释放 COM 引用，不阻塞宿主关闭。重新打开 OneNote，确认旧快照失效，新 `begin_edit` 正常。

验收完成后仅删除专用测试页。不要在真实笔记上执行破坏性验收。
