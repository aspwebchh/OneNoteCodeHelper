using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace OneNoteCodeHelper.Services
{
    /// <summary>
    /// 文件日志。COM 外接程序跑在 OneNote 进程里，异常没有任何可见出口，
    /// 出问题时这个日志基本是唯一线索，所以写日志本身绝不能再抛异常。
    /// </summary>
    internal static class AddInLog
    {
        private static readonly object Gate = new object();

        internal static string LogPath { get; } = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OneNoteCodeHelper",
            "log.txt");

        internal static void Info(string message) => Write("INFO ", message, null);

        internal static void Warn(string message, Exception ex = null) => Write("WARN ", message, ex);

        internal static void Error(string message, Exception ex = null) => Write("ERROR", message, ex);

        private static void Write(string level, string message, Exception ex)
        {
            try
            {
                var builder = new StringBuilder();
                builder.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture))
                       .Append(" [").Append(level).Append("] ").Append(message);

                if (ex != null)
                {
                    builder.AppendLine().Append(ex);
                }

                lock (Gate)
                {
                    var directory = Path.GetDirectoryName(LogPath);
                    if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    // 单个日志超过 1MB 就轮转，避免长期驻留把磁盘写满。
                    var file = new FileInfo(LogPath);
                    if (file.Exists && file.Length > 1024 * 1024)
                    {
                        var backup = LogPath + ".bak";
                        File.Delete(backup);
                        File.Move(LogPath, backup);
                    }

                    File.AppendAllText(LogPath, builder.AppendLine().ToString(), Encoding.UTF8);
                }
            }
            catch
            {
                // 日志写不进去也只能放弃，不能影响宿主。
            }
        }
    }
}
