using System;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using OneNoteCodeHelper.Services;
using OneNoteCodeHelper.Services.Agent;
using OneNoteCodeHelper.Views;

internal static class WindowPreview
{
    // 只渲染内存中的窗口内容，不启动 OneNote、不调用接口、不显示原生窗口。
    // 出几张图：Agent 自定义排版、选了第一个文字功能；模拟处理中：Agent 第 3 轮（前两轮摘录留着）、Agent 刚开始还没有输出、
    // 文字功能（思考框占满）、思考内容超出思考框（滚到最下面、顶上渐隐）；Agent 做完（步骤和结果）。都按窗口默认大小渲染，几张应一样高。
    // 另外是 AI 配置窗口的四页，以及配置文件读不了、保存前校验不过时的样子；最后是插入代码窗口：空的、填好代码的，
    // Markdown 的预览，以及自动识别时提示改选 Markdown 的各一张。
    internal static int Render(string directory)
    {
        Directory.CreateDirectory(directory);
        var xml = "<one:Page xmlns:one='" + OneNoteApi.OneNs + "' ID='preview' name='项目记录与实施计划' lastModifiedTime='2026-09-26T00:00:00Z'>" +
            "<one:Outline><one:OEChildren><one:OE objectID='p1' selected='all'><one:T><![CDATA[示例段落]]></one:T></one:OE></one:OEChildren></one:Outline></one:Page>";
        RenderOne(directory, "agent-window.png", xml, null);
        RenderOne(directory, "agent-window-text.png", xml, window => window.FunctionPicker.SelectedIndex = 1);
        RenderOne(directory, "agent-window-running.png", xml, window => SimulateAgent(window, true,
            "先读取页面概况，看看有哪些段落。",
            "正文统一为 11 磅微软雅黑，段后 6 磅。\n二级标题目前只是加粗的正文，需要改成原生二级标题。",
            "接下来先读取第 4 到第 12 段，确认列表缩进不受影响。\n第 7 段是代码，保持等宽字体不动。"));
        RenderOne(directory, "agent-window-running-empty.png", xml, window => SimulateAgent(window, false, (string)null));
        RenderOne(directory, "agent-window-running-text.png", xml, SimulateText);
        RenderOne(directory, "agent-window-running-overflow.png", xml, window => SimulateAgent(window, true,
            "先读取页面概况，看看有哪些段落。\n页面一共 18 段，前两段是标题。",
            "正文统一为 11 磅微软雅黑，段后 6 磅。\n二级标题目前只是加粗的正文，需要改成原生二级标题。\n第 3 到第 5 段是列表，缩进保持不变。",
            "第 9 段的引用块字号偏大，改成和正文一致。\n表格里的文字不动，只调整表格前后的间距。\n第 12 段有错别字「帐号」，改成「账号」。",
            "接下来先读取第 4 到第 12 段，确认列表缩进不受影响。\n第 7 段是代码，保持等宽字体不动。\n最后核对一遍标题层级，再提交草稿。"));
        RenderOne(directory, "agent-window-done.png", xml, SimulateDone);
        for (var page = 0; page < 4; page++)
        {
            var index = page;
            RenderSettings(directory, $"ai-settings-{index}.png", AiConfigStore.Default, null, window => window.ShowPage(index));
        }

        // Agent 页滚到底：三组开关。
        RenderSettings(directory, "ai-settings-3-switches.png", AiConfigStore.Default, null, window =>
        {
            window.ShowPage(3);
            window.AgentPage.ScrollToEnd();
        });

        // 配置文件读不了时的提示条，以及保存前校验不过时底部的红字（配置里有和 Agent 同名的功能）。
        var invalid = AiConfigStore.Parse(XElement.Parse("<AiConfig><Functions><Function name='智能校正'><Prompt>p</Prompt></Function>" +
            "<Function name='" + AiConfigStore.AgentFunctionName + "'><Prompt>p</Prompt></Function></Functions></AiConfig>"));
        RenderSettings(directory, "ai-settings-errors.png", invalid, new XmlException("根级别上的数据无效。 第 1 行，位置 1。"),
            window => window.TryBuildConfig(out _));

        RenderInsert(directory, "insert-code-window.png", new AddInSettings(), null);
        // 有一行很长、行数超出一屏：代码框和预览两个方向都出滚动条。
        RenderInsert(directory, "insert-code-window-filled.png", new AddInSettings { LanguageId = "python" }, window =>
        {
            window.CodeBox.Text = InsertSample;
            window.UpdatePreview();
        });
        // Markdown：标题、列表、待办、引用、表格和代码框都在预览里。
        RenderInsert(directory, "insert-code-window-markdown.png", new AddInSettings { InsertMarkdown = true }, window =>
        {
            window.CodeBox.Text = MarkdownSample;
            window.UpdatePreview(true);
        });
        // 自动识别时粘了 Markdown：提示可以改选 Markdown。
        RenderInsert(directory, "insert-code-window-markdown-hint.png", new AddInSettings(), window =>
        {
            window.CodeBox.Text = MarkdownSample;
            window.UpdatePreview();
            System.Threading.Thread.Sleep(500);
            DoEvents();
        });
        return 0;
    }

    private const string InsertSample =
        "import json\n" +
        "from pathlib import Path\n" +
        "\n" +
        "\n" +
        "# 读取配置文件，缺省时用内置的默认值\n" +
        "DEFAULTS = {\"theme\": \"light\", \"font\": \"Consolas\", \"size\": 10.5, \"borders\": True, \"language\": \"auto\", \"tab_width\": 4}\n" +
        "\n" +
        "\n" +
        "def load_settings(path: Path) -> dict:\n" +
        "    \"\"\"读出设置；文件不存在或格式不对时返回默认值。\"\"\"\n" +
        "    if not path.exists():\n" +
        "        return dict(DEFAULTS)\n" +
        "    try:\n" +
        "        data = json.loads(path.read_text(encoding=\"utf-8\"))\n" +
        "    except (OSError, ValueError):\n" +
        "        return dict(DEFAULTS)\n" +
        "    return {**DEFAULTS, **data}\n" +
        "\n" +
        "\n" +
        "class Highlighter:\n" +
        "    def __init__(self, settings: dict):\n" +
        "        self.settings = settings\n" +
        "        self.count = 0\n" +
        "\n" +
        "    def render(self, code: str) -> str:\n" +
        "        self.count += 1\n" +
        "        lines = code.splitlines()\n" +
        "        return \"\\n\".join(f\"{i + 1:>4}  {line}\" for i, line in enumerate(lines))\n" +
        "\n" +
        "\n" +
        "if __name__ == \"__main__\":\n" +
        "    settings = load_settings(Path(\"settings.json\"))\n" +
        "    print(Highlighter(settings).render(\"print('你好')\"))\n";

    private const string MarkdownSample =
        "# 部署说明\n" +
        "\n" +
        "按下面的步骤在测试机上安装，**先备份配置**，详见 [安装文档](https://example.com/install)。\n" +
        "\n" +
        "## 准备\n" +
        "\n" +
        "1. 安装 .NET Framework 4.8\n" +
        "2. 关闭 OneNote\n" +
        "   - 等待 `dllhost.exe` 退出\n" +
        "3. 运行安装脚本\n" +
        "\n" +
        "- [x] 已备份 `settings.xml`\n" +
        "- [ ] 通知同事 ~~周五~~ 周四升级\n" +
        "\n" +
        "> 注意：安装需要管理员权限。\n" +
        "\n" +
        "```powershell\n" +
        "powershell -ExecutionPolicy Bypass -File install.ps1\n" +
        "Get-Process dllhost | Where-Object { $_.Id -gt 0 }\n" +
        "```\n" +
        "\n" +
        "| 参数 | 说明 | 默认 |\n" +
        "| --- | --- | :---: |\n" +
        "| `-SkipBuild` | 跳过构建 | 否 |\n" +
        "| `-Force` | 强制结束进程 | 否 |\n";

    /// <summary>插入代码窗口。插入之前不会用到 PageEditor，传一个连不上 OneNote 的就行。</summary>
    private static void RenderInsert(string directory, string fileName, AddInSettings settings, Action<InsertCodeWindow> setup)
    {
        // 正文字体用默认的 Agent 设置，不读本机的 ai-settings.xml。
        var window = new InsertCodeWindow(new PageEditor(null), settings, IntPtr.Zero, new AgentOptions());
        setup?.Invoke(window);
        Snap(window, directory, fileName);
        window.Close();
    }

    /// <summary>
    /// AI 配置窗口的一页。不关窗：界面改动过时关窗会弹框问要不要保存。
    /// </summary>
    private static void RenderSettings(string directory, string fileName, AiConfig config, Exception loadError, Action<AiSettingsWindow> setup)
    {
        var window = new AiSettingsWindow(config, loadError, IntPtr.Zero);
        setup?.Invoke(window);
        Snap(window, directory, fileName);
    }

    /// <summary>Agent 做完：思考框收起，步骤留着，下面是结果。</summary>
    private static void SimulateDone(AgentWindow window)
    {
        SimulateAgent(window, true, "先读取页面概况，看看有哪些段落。");
        window.ShowAgentProgress(new AgentProgress { Step = new AgentStep { Id = 5, Text = "设置段落样式 · 二级标题 · 3 段", State = AgentStepState.Done } }, false);
        window.ShowAgentProgress(new AgentProgress { Step = new AgentStep { Id = 6, Text = "提交草稿 · 回读核验", State = AgentStepState.Done } }, false);
        window.ShowProgressLayout(false, false);
        window.SpinnerIcon.Visibility = Visibility.Collapsed;
        window.SuccessIcon.Visibility = Visibility.Visible;
        window.StatusText.Text = "已完成";
        window.DetailText.Text = "用时 58 秒";
        window.ResultText.Text = "已核验写入 16 段：1 段一级标题、3 段二级标题、12 段正文。\n修正的文字：\n• 第 12 段「帐号」→「账号」\n• 第 15 段「登陆」→「登录」";
        window.ResultPanel.Visibility = Visibility.Visible;
    }

    /// <summary>
    /// 用窗口自己的进度处理摆出 Agent 处理中的样子，不真的执行。turns 是每一轮的思考摘录，null 表示这一轮还没有输出；
    /// withSteps 为 false 时步骤区只有占位。
    /// </summary>
    private static void SimulateAgent(AgentWindow window, bool withSteps, params string[] turns)
    {
        ShowSpinner(window);
        window.ShowProgressLayout(true, true);
        for (var i = 0; i < turns.Length; i++)
        {
            window.ShowAgentProgress(new AgentProgress { Turn = i + 1, Status = "模型正在分析页面…", Thinking = "" }, false);
            if (turns[i] != null) window.ShowAgentProgress(new AgentProgress { Thinking = turns[i] }, false);
        }

        if (withSteps)
        {
            var steps = new[]
            {
                ("读取段落 · 18 段", AgentStepState.Done),
                ("设置段落样式 · 一级标题 · 1 段", AgentStepState.Done),
                ("设置段落样式 · 正文 · 12 段", AgentStepState.Done),
                ("设置段落间距 · 段后 6 磅 · 12 段", AgentStepState.Done),
                ("设置段落样式", AgentStepState.Running)
            };
            for (var i = 0; i < steps.Length; i++)
                window.ShowAgentProgress(new AgentProgress { Step = new AgentStep { Id = i + 1, Text = steps[i].Item1, State = steps[i].Item2 } }, false);
            window.ShowAgentProgress(new AgentProgress { Status = "模型正在准备：设置段落样式" }, false);
        }

        window.DetailText.Text = $"第 {Math.Max(1, turns.Length)} 轮 · 已用时 42 秒";
    }

    /// <summary>文字功能处理中：没有步骤，思考框占满下方卡片。</summary>
    private static void SimulateText(AgentWindow window)
    {
        window.FunctionPicker.SelectedIndex = 1;
        ShowSpinner(window);
        window.ShowProgressLayout(true, false);
        window.ShowAiProgress(new AiProgress("正在请 AI 智能校正：当前页 42 段…", 0, 3, "AI 正在思考",
            "第 3 段「登陆」应为「登录」。\n第 8 段中英文之间补空格，数字和单位之间不加。\n第 15 段的句号重复了，删掉一个。"), false);
        window.DetailText.Text = "AI 正在思考 · 已用时 18 秒";
    }

    private static void ShowSpinner(AgentWindow window)
    {
        window.InfoIcon.Visibility = Visibility.Collapsed;
        window.SpinnerIcon.Visibility = Visibility.Visible;
        window.IntroText.Visibility = Visibility.Collapsed;
    }

    private static void RenderOne(string directory, string fileName, string xml, Action<AgentWindow> setup)
    {
        var settings = new AddInSettings { AiModel = "example-model", AiEffort = "medium" };
        var window = new AgentWindow(new NoAccess(), "preview", xml, AiConfigStore.Default, settings, IntPtr.Zero);
        setup?.Invoke(window);
        Snap(window, directory, fileName);
        window.Close();
    }

    /// <summary>处理完界面线程上排着的回调，比如后台识别完成后回来刷新预览。</summary>
    private static void DoEvents()
    {
        var frame = new System.Windows.Threading.DispatcherFrame();
        System.Windows.Threading.Dispatcher.CurrentDispatcher.BeginInvoke(
            System.Windows.Threading.DispatcherPriority.Background, new Action(() => frame.Continue = false));
        System.Windows.Threading.Dispatcher.PushFrame(frame);
    }

    private static void Snap(System.Windows.Window window, string directory, string fileName)
    {
        // 按默认大小渲染：客户区约为窗口宽度减去 16px 边框、高度减去 39px 标题栏和边框。
        var width = (int)window.Width - 16;
        var height = (int)window.Height - 39;
        var content = (FrameworkElement)window.Content;
        window.Content = null;
        var host = new System.Windows.Controls.Border { Background = window.Background, Resources = window.Resources, Child = content };
        host.SetValue(System.Windows.Documents.TextElement.FontFamilyProperty, window.FontFamily);
        host.SetValue(System.Windows.Documents.TextElement.FontSizeProperty, window.FontSize);
        host.SetValue(System.Windows.Documents.TextElement.ForegroundProperty, window.Foreground);
        host.Measure(new Size(width, height));
        host.Arrange(new Rect(0, 0, width, height));
        // 思考框在布局完成后才滚到最下面（ScrollChanged 里排的滚动要再过几遍布局才生效）。
        for (var i = 0; i < 3; i++) host.UpdateLayout();
        var bitmap = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
        bitmap.Render(host);
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(bitmap));
        var path = Path.Combine(directory, fileName);
        using (var file = File.Create(path)) encoder.Save(file);
        Console.WriteLine(path);
    }

    private sealed class NoAccess : IOneNotePageAccess
    {
        public string GetPageContent(string pageId, PageInfo info) => throw new InvalidOperationException("Preview cannot access OneNote.");
        public void UpdatePageContent(string xml, DateTime expectedLastModified) => throw new InvalidOperationException("Preview cannot write OneNote.");
        public void DeletePageContent(string pageId, string objectId, DateTime expectedLastModified) => throw new InvalidOperationException("Preview cannot write OneNote.");
    }
}
