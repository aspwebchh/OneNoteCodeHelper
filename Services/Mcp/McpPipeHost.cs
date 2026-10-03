using System;
using System.Collections.Concurrent;
using System.ComponentModel;
using System.IO;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Win32.SafeHandles;
using OneNoteCodeHelper.Mcp;

namespace OneNoteCodeHelper.Services.Mcp
{
    /// <summary>只接受同一 Windows 用户的本机连接。停止不等待 COM，已开始写入的任务由提交器核验。</summary>
    internal sealed class McpPipeHost : IDisposable
    {
        private readonly McpEditService _service;
        private readonly string _name;
        private readonly CancellationTokenSource _stop = new CancellationTokenSource();
        private readonly ConcurrentDictionary<NamedPipeServerStream, byte> _pipes = new ConcurrentDictionary<NamedPipeServerStream, byte>();
        private readonly ConcurrentDictionary<string, byte> _clients = new ConcurrentDictionary<string, byte>();
        private readonly TaskCompletionSource<bool> _ready = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _started;
        internal Task Ready => _ready.Task;
        internal McpPipeHost(McpEditService service, string name = null) { _service = service; _name = name ?? McpProtocol.DefaultPipeName; }
        internal void Start()
        {
            if (Interlocked.Exchange(ref _started, 1) != 0) throw new InvalidOperationException("管道服务已经启动。");
            _ = Task.Run(AcceptAsync);
        }

        private async Task AcceptAsync()
        {
            var first = true;
            try
            {
                while (!_stop.IsCancellationRequested)
                {
                    var pipe = CreatePipe(_name, first);
                    first = false;
                    _pipes.TryAdd(pipe, 0);
                    if (_stop.IsCancellationRequested) { _pipes.TryRemove(pipe, out _); pipe.Dispose(); break; }
                    _ready.TrySetResult(true);
                    try { await pipe.WaitForConnectionAsync(_stop.Token).ConfigureAwait(false); }
                    catch { _pipes.TryRemove(pipe, out _); pipe.Dispose(); throw; }
                    // 读循环继续处理 cancel；页面会话自身负责串行执行。
                    _ = ServeAsync(pipe);
                }
            }
            catch (Exception ex)
            {
                if (!_stop.IsCancellationRequested)
                {
                    _ready.TrySetException(ex);
                    AddInLog.Info("MCP pipe state=failed type=" + ex.GetType().Name);
                    Dispose();
                }
            }
            finally { _ready.TrySetCanceled(); }
        }

        private async Task ServeAsync(NamedPipeServerStream pipe)
        {
            string client = null;
            var owned = false;
            var pending = new ConcurrentDictionary<string, CancellationTokenSource>();
            var writes = new SemaphoreSlim(1, 1);
            using (var disconnected = CancellationTokenSource.CreateLinkedTokenSource(_stop.Token))
            {
                try
                {
                    // 防止无握手连接占用全部服务实例。
                    using (var helloTimeout = CancellationTokenSource.CreateLinkedTokenSource(disconnected.Token))
                    {
                        helloTimeout.CancelAfter(McpProtocol.ConnectTimeoutMs);
                        var first = await McpProtocol.ReadAsync(pipe, helloTimeout.Token).ConfigureAwait(false);
                        if (first == null) return;
                        var hello = McpJson.Deserialize<McpRequest>(first);
                        var reply = Response(hello);
                        try
                        {
                            CheckVersion(hello);
                            if (hello.Method != "hello") throw new McpFault("hello_required", "必须先连接插件。");
                            _service.BindClient(hello.ClientId, hello.ResumeSecret);
                            client = hello.ClientId;
                            if (!_clients.TryAdd(client, 0)) throw new McpFault("client_connected", "该客户端连接仍在处理断开，请稍后查询状态。");
                            owned = true; reply.Success = true;
                            reply.ResultJson = McpJson.Serialize(new { instance_id = _service.InstanceId, ipc_version = McpProtocol.Version });
                        }
                        catch (McpFault ex) { reply.ErrorCode = ex.Code; reply.ErrorMessage = ex.Message; }
                        await SendAsync(pipe, writes, reply, disconnected.Token).ConfigureAwait(false);
                        if (!reply.Success) return;
                    }
                    while (!disconnected.IsCancellationRequested)
                    {
                        var json = await McpProtocol.ReadAsync(pipe, disconnected.Token).ConfigureAwait(false);
                        if (json == null) break;
                        var request = McpJson.Deserialize<McpRequest>(json);
                        CheckVersion(request);
                        if (request.ClientId != client || request.InstanceId != _service.InstanceId) throw new McpFault("instance_mismatch", "插件实例或客户端不匹配。");
                        if (request.Method == "cancel")
                        {
                            if (request.CancelRequestId != null && pending.TryGetValue(request.CancelRequestId, out var target))
                                try { target.Cancel(); } catch (ObjectDisposedException) { }
                            continue; // cancel 是单向消息，不污染响应队列。
                        }
                        if (string.IsNullOrEmpty(request.RequestId) || request.RequestId.Length > 128)
                            throw new McpFault("invalid_request", "请求标识无效。");
                        if (pending.Count >= 32)
                        {
                            var limited = Response(request); limited.ErrorCode = "request_limit"; limited.ErrorMessage = "并发请求达到上限，请等待已有调用结束。";
                            await SendAsync(pipe, writes, limited, disconnected.Token).ConfigureAwait(false);
                            continue;
                        }
                        var source = CancellationTokenSource.CreateLinkedTokenSource(disconnected.Token);
                        if (!pending.TryAdd(request.RequestId, source)) { source.Dispose(); throw new McpFault("duplicate_request", "请求标识重复。"); }
                        _ = Task.Run(async () =>
                        {
                            var reply = Response(request);
                            try
                            {
                                if (request.Method == "catalog") reply.ResultJson = McpJson.Serialize(_service.Catalog);
                                else if (request.Method == "call") reply.ResultJson = _service.Call(client, request.ToolName, request.ArgumentsJson, source.Token);
                                else throw new McpFault("unknown_method", "未知 IPC 方法。");
                                reply.Success = true;
                            }
                            catch (McpFault ex) { reply.ErrorCode = ex.Code; reply.ErrorMessage = ex.Message; }
                            catch (AiException ex) { reply.ErrorCode = "invalid_arguments"; reply.ErrorMessage = ex.Message; }
                            catch (OperationCanceledException) { reply.ErrorCode = "cancelled"; reply.ErrorMessage = "请求已取消；请查询快照状态，禁止自动重发写入。"; }
                            catch (Exception) { reply.ErrorCode = "operation_failed"; reply.ErrorMessage = "插件操作失败，请查询状态或检查 OneNote 是否仍连接。"; }
                            try { await SendAsync(pipe, writes, reply, disconnected.Token).ConfigureAwait(false); }
                            catch (Exception) { /* 结果留在会话内，断线不重放写入。 */ }
                            finally { pending.TryRemove(request.RequestId, out _); source.Dispose(); }
                        });
                    }
                }
                catch (Exception ex)
                {
                    if (!_stop.IsCancellationRequested && !(ex is IOException) && !(ex is OperationCanceledException))
                        AddInLog.Info("MCP connection state=closed type=" + ex.GetType().Name);
                }
                finally
                {
                    disconnected.Cancel();
                    // 不等任务收尾；取消令牌和完成记录仍由服务保存。
                    if (owned) { _service.DisconnectClient(client); _clients.TryRemove(client, out _); }
                    _pipes.TryRemove(pipe, out _); pipe.Dispose();
                    // writes 不能在仍有回包的任务使用时释放；SemaphoreSlim 无内核句柄。
                }
            }
        }

        private McpResponse Response(McpRequest request) => new McpResponse { RequestId = request?.RequestId, InstanceId = _service.InstanceId };
        private static void CheckVersion(McpRequest request)
        {
            if (request == null || request.ProtocolVersion != McpProtocol.Version) throw new McpFault("protocol_mismatch", "IPC 协议版本不匹配，请一起更新插件和 MCP 程序。");
        }
        private static async Task SendAsync(NamedPipeServerStream pipe, SemaphoreSlim writes, McpResponse reply, CancellationToken cancellation)
        {
            var json = McpJson.Serialize(reply);
            if (System.Text.Encoding.UTF8.GetByteCount(json) > McpProtocol.MaxFrameBytes)
            {
                reply.Success = false; reply.ResultJson = null; reply.ErrorCode = "response_too_large"; reply.ErrorMessage = "结果超出 IPC 大小限制，请缩小读取批次。";
                json = McpJson.Serialize(reply);
            }
            await writes.WaitAsync(cancellation).ConfigureAwait(false);
            try { await McpProtocol.WriteAsync(pipe, json, cancellation).ConfigureAwait(false); }
            finally { writes.Release(); }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SecurityAttributes { internal int Length; internal IntPtr Descriptor; internal int InheritHandle; }
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafePipeHandle CreateNamedPipe(string name, uint openMode, uint pipeMode, uint maxInstances,
            uint outBuffer, uint inBuffer, uint timeout, ref SecurityAttributes security);

        private static NamedPipeServerStream CreatePipe(string name, bool first)
        {
            if (string.IsNullOrEmpty(name) || name.IndexOfAny(new[] { '\\', '/' }) >= 0) throw new ArgumentException("管道名称无效。");
            var security = new PipeSecurity();
            var user = WindowsIdentity.GetCurrent().User;
            security.SetAccessRuleProtection(true, false);
            security.SetOwner(user);
            security.AddAccessRule(new PipeAccessRule(new SecurityIdentifier(WellKnownSidType.NetworkSid, null), PipeAccessRights.FullControl, AccessControlType.Deny));
            security.AddAccessRule(new PipeAccessRule(user, PipeAccessRights.FullControl, AccessControlType.Allow));
            var bytes = security.GetSecurityDescriptorBinaryForm();
            var pointer = Marshal.AllocHGlobal(bytes.Length);
            try
            {
                Marshal.Copy(bytes, 0, pointer, bytes.Length);
                var attrs = new SecurityAttributes { Length = Marshal.SizeOf(typeof(SecurityAttributes)), Descriptor = pointer };
                // PIPE_REJECT_REMOTE_CLIENTS；第一次创建还拒绝已存在的同名管道，防止意外连到别的服务。
                var handle = CreateNamedPipe("\\\\.\\pipe\\" + name, 0x40000003u | (first ? 0x00080000u : 0), 0x00000008u, 16, 65536, 65536, 0, ref attrs);
                if (handle.IsInvalid) { var error = Marshal.GetLastWin32Error(); handle.Dispose(); throw new Win32Exception(error); }
                try { return new NamedPipeServerStream(PipeDirection.InOut, true, false, handle); }
                catch { handle.Dispose(); throw; }
            }
            finally { Marshal.FreeHGlobal(pointer); }
        }

        public void Dispose()
        {
            if (_stop.IsCancellationRequested) return;
            _stop.Cancel(); _service.Dispose();
            foreach (var pipe in _pipes.Keys) pipe.Dispose();
        }
    }
}
