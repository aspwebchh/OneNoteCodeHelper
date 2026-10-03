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
        /// <summary>同时保持的连接数（含正在等待连接的监听实例）。满额时新客户端等空位，已有连接不受影响。</summary>
        internal const int MaxConnections = 64;
        private const int MinRetryDelayMs = 500;
        private const int MaxRetryDelayMs = 30000;
        private readonly McpEditService _service;
        private readonly string _name;
        private readonly CancellationTokenSource _stop = new CancellationTokenSource();
        private readonly ConcurrentDictionary<NamedPipeServerStream, byte> _pipes = new ConcurrentDictionary<NamedPipeServerStream, byte>();
        private readonly ConcurrentDictionary<string, byte> _clients = new ConcurrentDictionary<string, byte>();
        private readonly SemaphoreSlim _slots = new SemaphoreSlim(MaxConnections, MaxConnections);
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
            var delay = MinRetryDelayMs;
            try
            {
                while (!_stop.IsCancellationRequested)
                {
                    // 每个实例从创建到连接结束都占一个名额；不会去建超出系统上限的实例而失败。
                    await _slots.WaitAsync(_stop.Token).ConfigureAwait(false);
                    NamedPipeServerStream pipe = null;
                    try
                    {
                        pipe = CreatePipe(_name, first);
                        first = false;
                        _pipes.TryAdd(pipe, 0);
                        if (_stop.IsCancellationRequested) { Release(pipe); break; }
                        _ready.TrySetResult(true);
                        await pipe.WaitForConnectionAsync(_stop.Token).ConfigureAwait(false);
                    }
                    catch (Exception ex) when (!first && !_stop.IsCancellationRequested)
                    {
                        // 单个监听实例出错（比如客户端连上随即断开）不拖垮已有连接和草稿，退避后重建。
                        // 只有第一次建管道失败（同名管道属于别的进程）才放弃。
                        Release(pipe);
                        AddInLog.Info("MCP pipe state=retry type=" + ex.GetType().Name);
                        await Task.Delay(delay, _stop.Token).ConfigureAwait(false);
                        delay = Math.Min(delay * 2, MaxRetryDelayMs);
                        continue;
                    }
                    catch { Release(pipe); throw; }
                    delay = MinRetryDelayMs;
                    // 读循环继续处理 cancel；页面会话自身负责串行执行。名额由 ServeAsync 结束时归还。
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
                    // 防止无握手连接一直占着名额。net48 的管道读取不响应取消令牌，超时只能关掉管道，挂起的读取随即以 0 字节结束。
                    string first;
                    using (var helloTimeout = new CancellationTokenSource(McpProtocol.ConnectTimeoutMs))
                    using (helloTimeout.Token.Register(pipe.Dispose))
                    {
                        try { first = await McpProtocol.ReadAsync(pipe, CancellationToken.None).ConfigureAwait(false); }
                        catch (Exception) when (helloTimeout.IsCancellationRequested) { first = null; }
                        if (helloTimeout.IsCancellationRequested)
                        {
                            AddInLog.Info("MCP connection state=hello_timeout");
                            return;
                        }
                    }
                    if (first == null) return;
                    // 握手的 hello、reply 限在这个块里，和下面请求处理里的同名变量隔开。
                    {
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
                    Release(pipe);
                    // writes 不能在仍有回包的任务使用时释放；SemaphoreSlim 无内核句柄。
                }
            }
        }

        /// <summary>关掉一个实例并归还名额。实例还没建出来时 pipe 为 null，只归还名额。</summary>
        private void Release(NamedPipeServerStream pipe)
        {
            if (pipe != null) { _pipes.TryRemove(pipe, out _); pipe.Dispose(); }
            _slots.Release();
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
                var handle = CreateNamedPipe("\\\\.\\pipe\\" + name, 0x40000003u | (first ? 0x00080000u : 0), 0x00000008u, MaxConnections, 65536, 65536, 0, ref attrs);
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
