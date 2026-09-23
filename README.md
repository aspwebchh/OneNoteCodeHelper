# OneNote 代码高亮

OneNote 桌面版的 COM 外接程序，把笔记里的代码渲染成带底色的高亮代码框。
目前支持 **Java**、**Lua**、**PowerShell**、**Bat**、**Bash**、**XML**、**HTML**、**CSS**，
以及不着色的 **纯文本**（只要等宽字体和代码框，适合放日志、命令输出）。

功能区「开始」选项卡上会多出一个「代码高亮」组：

| 控件 | 作用 |
|---|---|
| 高亮选中 | 选中页面上已有的代码文字，原地替换成高亮代码框 |
| 插入代码 | 打开窗口粘贴代码，预览确认后插入到当前页 |
| 语言 | 自动识别 / Java / Lua / PowerShell / Bat / Bash / XML / HTML / CSS / 纯文本。纯文本只能手动选，自动识别不会选它 |
| 深色主题 | 在浅色（类 IntelliJ）与深色（类 VS Code Dark+）之间切换 |
| 字体（插入窗口内） | 默认 Consolas。代码里有中文时改选「NSimSun」新宋体，中英文才能对齐 |
| 诊断日志 | 打开日志文件 |

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
Services/
  OneNoteApi.cs             OneNote COM 封装
  PageEditor.cs             选区解析、就地替换、插入
  CodeBlockBuilder.cs       生成 one:Table 代码框 XML
  OneNoteHtmlEncoder.cs     Token -> one:T 里的 span HTML
  AddInSettings.cs          设置与持久化
  AddInLog.cs               文件日志
  RenderDiagnostics.cs      不碰 OneNote 就能跑通渲染链路的诊断入口
Highlighting/
  TokenKind / Token / ILanguage / LexerCursor / LanguageRegistry
  Languages/                每种语言一个 ILanguage 实现；XML 与 HTML 共用 MarkupLexer，
                            HTML 的 <style> 内容交给 CssLanguage 着色
  Themes/CodeTheme.cs, CodeThemes.cs
Views/
  InsertCodeWindow.xaml     插入代码窗口
  CodePreviewRenderer.cs    用同一套 token 流渲染 WPF 预览
install.ps1                 一键构建 + 安装 / 卸载
Tools/register.ps1          只做注册这一步
Tools/unregister.ps1        只做注销这一步
```

## 加一种语言

1. 在 `Highlighting/Languages/` 下实现 `ILanguage`：`Tokenize` 切 token，`ScoreLikelihood` 给自动识别打分。
   `Tokenize` 必须保证返回的 token 按序、不重叠、完整覆盖整个源码 — `RenderDiagnostics.VerifyCoverage`
   就是用来验这条契约的。
2. 在 `LanguageRegistry.All` 里加一行。

功能区下拉、设置、预览都会自动跟上，不需要改别的地方。

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
- **OneNote 不认 CSS 字体栈**。给 `font-family:Consolas,NSimSun` 它只取 `Consolas`；
  而且遇到中文时会把 `font-family` 整个丢掉、回退到自己的中文字体。
  所以代码里有中文注释又想对齐，只能整体换成中英文都等宽的字体（插入窗口里可以选「NSimSun」新宋体）。
- **WPF 窗口必须在 STA 线程上创建**。OneNote 对代理进程的功能区回调在 MTA 上运行；
  插入窗口使用独立 STA 线程和 `ShowDialog()` 消息循环。
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
- **`-replace` 的第一个参数是正则**。`$path -replace '\', '/'` 里的单个反斜杠是非法模式，
  会直接抛 `InvalidRegularExpression`。要替换路径分隔符用 `.Replace()`。
