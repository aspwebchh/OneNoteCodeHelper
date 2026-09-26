using System;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Office.Interop.OneNote;
using OneNoteCodeHelper.Services;
using OneNoteCodeHelper.Services.Agent;
using OneNoteCodeHelper.Views;

internal static class WindowPreview
{
    // 只渲染内存中的窗口内容，不启动 OneNote、不调用接口、不显示原生窗口。
    // 出几张图：Agent 自定义排版、选了第一个文字功能、模拟处理中（摘录框超过三行、只有一行、还没有输出三种）；
    // 窗口高度固定，几张应一样高。
    internal static int Render(string directory)
    {
        Directory.CreateDirectory(directory);
        var xml = "<one:Page xmlns:one='" + OneNoteApi.OneNs + "' ID='preview' name='项目记录与实施计划' lastModifiedTime='2026-09-26T00:00:00Z'>" +
            "<one:Outline><one:OEChildren><one:OE objectID='p1' selected='all'><one:T><![CDATA[示例段落]]></one:T></one:OE></one:OEChildren></one:Outline></one:Page>";
        RenderOne(directory, "agent-window.png", xml, null);
        RenderOne(directory, "agent-window-text.png", xml, window => window.FunctionPicker.SelectedIndex = 1);
        RenderOne(directory, "agent-window-running.png", xml, window => SimulateRunning(window,
            "…正文统一为 11 磅微软雅黑，段后 6 磅。二级标题目前只是加粗的正文，需要改成原生二级标题，\n" +
            "接下来先读取第 4 到第 12 段的格式，确认列表缩进不受影响。最后检查草稿再提交。"));
        RenderOne(directory, "agent-window-running-short.png", xml, window => SimulateRunning(window, "先读取页面概况，看看有哪些段落。"));
        RenderOne(directory, "agent-window-running-empty.png", xml, window => SimulateRunning(window, null));
        return 0;
    }

    /// <summary>只摆出处理中各块的样子（思考摘录、步骤），不真的执行。thinking 为 null 时摆出占位文字。</summary>
    private static void SimulateRunning(AgentWindow window, string thinking)
    {
        window.InfoIcon.Visibility = Visibility.Collapsed;
        window.SpinnerIcon.Visibility = Visibility.Visible;
        window.StatusText.Text = "模型正在准备：设置段落样式";
        window.DetailText.Text = "第 3 轮 · 已用时 42 秒";
        window.IntroText.Visibility = Visibility.Collapsed;
        window.ThinkingBox.Visibility = Visibility.Visible;
        window.ThinkingText.Text = thinking ?? "等待模型输出…";
        if (thinking == null) window.ThinkingText.Foreground = (Brush)window.FindResource("Faint");
        var done = new SolidColorBrush(Color.FromRgb(0x16, 0xA3, 0x4A));
        var running = new SolidColorBrush(Color.FromRgb(0x25, 0x63, 0xEB));
        window.StepsList.ItemsSource = new[]
        {
            new { Glyph = "✓", Brush = (Brush)done, Text = "读取段落 · 18 段" },
            new { Glyph = "✓", Brush = (Brush)done, Text = "设置段落样式 · 一级标题 · 1 段" },
            new { Glyph = "✓", Brush = (Brush)done, Text = "设置段落样式 · 正文 · 12 段" },
            new { Glyph = "✓", Brush = (Brush)done, Text = "设置段落间距 · 段后 6 磅 · 12 段" },
            new { Glyph = "…", Brush = (Brush)running, Text = "设置段落样式" }
        };
        window.StepsPanel.Visibility = Visibility.Visible;
    }

    private static void RenderOne(string directory, string fileName, string xml, Action<AgentWindow> setup)
    {
        var settings = new AddInSettings { AiModel = "example-model", AiEffort = "medium" };
        var window = new AgentWindow(new NoAccess(), "preview", xml, AiConfigStore.Default, settings, IntPtr.Zero);
        setup?.Invoke(window);
        // 窗口宽高固定：客户区约为窗口宽度减去 16px 边框、高度减去 39px 标题栏和边框。
        var width = (int)window.Width - 16;
        var height = (int)window.Height - 39;
        var content = (FrameworkElement)window.Content;
        window.Content = null;
        var host = new System.Windows.Controls.Border { Background = window.Background, Resources = window.Resources, Child = content };
        host.SetValue(System.Windows.Documents.TextElement.FontFamilyProperty, window.FontFamily);
        host.SetValue(System.Windows.Documents.TextElement.FontSizeProperty, window.FontSize);
        host.Measure(new Size(width, height));
        host.Arrange(new Rect(0, 0, width, height));
        host.UpdateLayout();
        var bitmap = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
        bitmap.Render(host);
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(bitmap));
        var path = Path.Combine(directory, fileName);
        using (var file = File.Create(path)) encoder.Save(file);
        Console.WriteLine(path);
        window.Close();
    }

    private sealed class NoAccess : IOneNotePageAccess
    {
        public string GetPageContent(string pageId, PageInfo info) => throw new InvalidOperationException("Preview cannot access OneNote.");
        public void UpdatePageContent(string xml, DateTime expectedLastModified) => throw new InvalidOperationException("Preview cannot write OneNote.");
    }
}
