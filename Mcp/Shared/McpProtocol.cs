#nullable disable
using System;
using System.Diagnostics;
using System.IO;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OneNoteCodeHelper.Mcp
{
    // 仅共享协议和帧格式，不引用 WPF、Office、MCP SDK 或 JSON 库。
    internal static class McpProtocol
    {
        internal const int Version = 1;
        internal const int ConnectTimeoutMs = 3000;
        internal const int MaxFrameBytes = 8 * 1024 * 1024;
        internal static string DefaultPipeName =>
            "OneNoteCodeHelper.Mcp." + WindowsIdentity.GetCurrent().User.Value + "." +
            Process.GetCurrentProcess().SessionId;

        internal static async Task<string> ReadAsync(Stream stream, CancellationToken cancellation)
        {
            var header = new byte[4];
            var first = await stream.ReadAsync(header, 0, 4, cancellation).ConfigureAwait(false);
            if (first == 0) return null;
            await ReadRestAsync(stream, header, first, cancellation).ConfigureAwait(false);
            var length = header[0] | header[1] << 8 | header[2] << 16 | header[3] << 24;
            if (length <= 0 || length > MaxFrameBytes) throw new InvalidDataException("IPC 消息长度无效。");
            var data = new byte[length];
            await ReadRestAsync(stream, data, 0, cancellation).ConfigureAwait(false);
            return new UTF8Encoding(false, true).GetString(data);
        }

        private static async Task ReadRestAsync(Stream stream, byte[] buffer, int offset, CancellationToken cancellation)
        {
            while (offset < buffer.Length)
            {
                var count = await stream.ReadAsync(buffer, offset, buffer.Length - offset, cancellation).ConfigureAwait(false);
                if (count == 0) throw new EndOfStreamException("IPC 消息未完整接收。");
                offset += count;
            }
        }

        // 同一流的发送由调用方串行化。
        internal static async Task WriteAsync(Stream stream, string json, CancellationToken cancellation)
        {
            var data = new UTF8Encoding(false, true).GetBytes(json);
            if (data.Length == 0 || data.Length > MaxFrameBytes) throw new InvalidDataException("IPC 消息过大。");
            var length = data.Length;
            var header = new[] { (byte)length, (byte)(length >> 8), (byte)(length >> 16), (byte)(length >> 24) };
            await stream.WriteAsync(header, 0, header.Length, cancellation).ConfigureAwait(false);
            await stream.WriteAsync(data, 0, data.Length, cancellation).ConfigureAwait(false);
            await stream.FlushAsync(cancellation).ConfigureAwait(false);
        }
    }

    internal sealed class McpRequest
    {
        public int ProtocolVersion { get; set; } = McpProtocol.Version;
        public string RequestId { get; set; }
        public string ClientId { get; set; }
        public string ResumeSecret { get; set; }
        public string InstanceId { get; set; }
        public string Method { get; set; }
        public string ToolName { get; set; }
        public string ArgumentsJson { get; set; } = "{}";
        public string CancelRequestId { get; set; }
    }

    internal sealed class McpResponse
    {
        public int ProtocolVersion { get; set; } = McpProtocol.Version;
        public string RequestId { get; set; }
        public string InstanceId { get; set; }
        public bool Success { get; set; }
        public string ResultJson { get; set; }
        public string ErrorCode { get; set; }
        public string ErrorMessage { get; set; }
    }

    internal sealed class McpToolDefinition
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string InputSchemaJson { get; set; }
        public bool ReadOnly { get; set; }
        public bool Destructive { get; set; }
    }

    internal sealed class McpCatalog
    {
        public string Instructions { get; set; }
        public McpToolDefinition[] Tools { get; set; }
    }
}
