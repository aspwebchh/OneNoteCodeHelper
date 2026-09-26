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
    internal static int Render(string directory)
    {
        Directory.CreateDirectory(directory);
        var xml = "<one:Page xmlns:one='" + OneNoteApi.OneNs + "' ID='preview' name='项目记录与实施计划' lastModifiedTime='2026-09-26T00:00:00Z'>" +
            "<one:Outline><one:OEChildren><one:OE objectID='p1' selected='all'><one:T><![CDATA[示例段落]]></one:T></one:OE></one:OEChildren></one:Outline></one:Page>";
        var window = new AgentWindow(new NoAccess(), "preview", xml, AiConfigStore.Default, "example-model", "medium", IntPtr.Zero);
        // 窗口固定宽度、高度随内容：客户区宽度约为窗口宽度减去 16px 边框，高度按内容量出来。
        var width = (int)window.Width - 16;
        var content = (FrameworkElement)window.Content;
        window.Content = null;
        var host = new System.Windows.Controls.Border { Background = window.Background, Resources = window.Resources, Child = content };
        host.SetValue(System.Windows.Documents.TextElement.FontFamilyProperty, window.FontFamily);
        host.SetValue(System.Windows.Documents.TextElement.FontSizeProperty, window.FontSize);
        host.Measure(new Size(width, double.PositiveInfinity));
        var height = (int)Math.Ceiling(host.DesiredSize.Height);
        host.Arrange(new Rect(0, 0, width, height));
        host.UpdateLayout();
        var bitmap = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
        bitmap.Render(host);
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(bitmap));
        var path = Path.Combine(directory, "agent-window.png");
        using (var file = File.Create(path)) encoder.Save(file);
        Console.WriteLine(path);
        window.Close();
        return 0;
    }
    private sealed class NoAccess : IOneNotePageAccess
    {
        public string GetPageContent(string pageId, PageInfo info) => throw new InvalidOperationException("Preview cannot access OneNote.");
        public void UpdatePageContent(string xml, DateTime expectedLastModified) => throw new InvalidOperationException("Preview cannot write OneNote.");
    }
}
