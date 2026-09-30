using System;
using System.Linq;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using OneNoteCodeHelper.Services;

namespace OneNoteCodeHelper.Views
{
    /// <summary>窗口左上角和标题栏的图标，和功能区按钮用的是同一张嵌入资源 Resources\xxx.png。</summary>
    internal static class WindowIcons
    {
        /// <summary>读嵌入资源里的图标，找不到就不显示。</summary>
        internal static ImageSource Load(string name)
        {
            try
            {
                var assembly = typeof(WindowIcons).Assembly;
                var resource = assembly.GetManifestResourceNames()
                    .FirstOrDefault(n => n.EndsWith("." + name + ".png", StringComparison.OrdinalIgnoreCase));
                if (resource == null) return null;

                using (var stream = assembly.GetManifestResourceStream(resource))
                {
                    var image = new BitmapImage();
                    image.BeginInit();
                    image.CacheOption = BitmapCacheOption.OnLoad;
                    image.StreamSource = stream;
                    image.EndInit();
                    image.Freeze();
                    return image;
                }
            }
            catch (Exception ex)
            {
                AddInLog.Warn("读取窗口图标失败：" + name, ex);
                return null;
            }
        }
    }
}
