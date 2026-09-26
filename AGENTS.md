# 仓库工作说明

本文件适用于整个仓库。项目是 OneNote 桌面版（Office16）的 C# COM 外接程序，目标框架为 .NET Framework 4.8，使用 WPF 窗口。用户功能、安装步骤和已验证的 OneNote 限制见 `README.md`；修改相关代码前先查对应章节。

## 代码入口

- `AddIn.cs` 和 `Ribbon.xml`：COM 外接程序生命周期、功能区及其回调。
- `Interop/`：Office COM 接口与 Win32 声明。`Services/OneNoteApi.cs` 封装 OneNote 调用，`Services/PageEditor.cs` 处理页面 XML、选区和写回。
- `Highlighting/`：语言词法分析、自动识别和主题；`Services/CodeBlockBuilder.cs`、`OneNoteHtmlEncoder.cs` 生成 OneNote 代码框；`Views/` 提供插入、预览和 Agent 窗口。
- `Services/AiConfig.cs`、`AiClient.cs`、`AiOptimizer.cs`、`RichParagraph.cs`、`BlankLines.cs`：AI 配置、请求、编排及保留格式的写回。
- `Tools/`：不依赖正在运行的 OneNote 的回归脚本，以及需要管理员权限的注册脚本。

## 修改时保持的约束

- 保持 `OneNoteCodeHelper.csproj` 的 `net48` 目标框架、强名称签名，以及 `lib/` 中 Office PIA 的早绑定引用。此项目不能依赖 `dynamic` 或运行时类型库查找来调用 OneNote。
- 生成 `one:T` 内容时按 HTML 转义并保留行首缩进、空行；改动页面 XML 时检查选区、嵌套段落和原有格式。AI 写回不能合并或拆分段落，也不能覆盖处理期间用户已经改动的段落。
- `ILanguage.Tokenize` 返回的 token 必须按顺序、无重叠、完整覆盖源码。新语言同时注册到 `LanguageRegistry.All`，按需更新语言族，并在 `Tools/detect-samples/<语言 id>/` 添加识别样本。
- OneNote 的功能区回调不能执行耗时操作或同步弹框；WPF 窗口须在 STA 线程上创建。涉及 COM、线程或注册机制的改动先核对 `README.md` 的“几个踩过的坑”。
- 修改 `.ps1` 文件时保留 UTF-8 BOM，以便 Windows PowerShell 5.1 正确解析中文内容。不要把本机的 API Key、`%APPDATA%\OneNoteCodeHelper\ai-settings.xml` 或日志写入仓库。

## 构建与验证

在 Windows PowerShell 中，从仓库根目录执行：

```powershell
dotnet build OneNoteCodeHelper.sln -c Release
powershell -ExecutionPolicy Bypass -File Tools\detect-test.ps1
powershell -ExecutionPolicy Bypass -File Tools\ai-merge-test.ps1
powershell -ExecutionPolicy Bypass -File Tools\highlight-selection-test.ps1
```

回归脚本默认读取 `bin\Release\net48\OneNoteCodeHelper.dll`。改动范围较小时至少运行对应脚本；修改共享渲染、页面编辑或构建配置时运行全部脚本。`ai-merge-test.ps1` 不加 `-Live`，避免真实接口调用。

`install.ps1`、`uninstall.ps1` 和 `Tools/register.ps1`、`Tools/unregister.ps1` 会关闭 OneNote 或修改 COM 注册表；普通构建和验证不要运行它们。仓库跟踪 `bin/` 下的构建产物，提交前检查工作区，避免把无关的 DLL/PDB 变化带入改动。
