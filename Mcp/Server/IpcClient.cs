using System.Collections.Concurrent;
using System.IO.Pipes;
using System.Security.Principal;
using System.Text.Json;

namespace OneNoteCodeHelper.Mcp;

internal sealed class IpcException(string code, string message) : Exception(message)
{
    internal string Code { get; } = code;
}

/// <summary>仅转发工具 RPC；连接恢复不会自动重放任何调用。</summary>
internal sealed class IpcClient(string pipeName) : IAsyncDisposable
{
    private sealed class Connection(NamedPipeClientStream pipe)
    {
        internal NamedPipeClientStream Pipe { get; } = pipe;
        internal SemaphoreSlim Writes { get; } = new(1, 1);
        internal ConcurrentDictionary<string, TaskCompletionSource<McpResponse>> Pending { get; } = new();
        internal string Instance = "";
    }

    private readonly string _client = Guid.NewGuid().ToString("N");
    private readonly string _secret = Guid.NewGuid().ToString("N");
    private readonly SemaphoreSlim _connectGate = new(1, 1);
    private readonly ConcurrentDictionary<string, string> _snapshots = new();
    private volatile Connection? _connection;
    private volatile bool _disposed;

    internal async Task ConnectAsync(CancellationToken cancellation)
    {
        await _connectGate.WaitAsync(cancellation);
        try
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (_connection?.Pipe.IsConnected == true) return;
            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
            timeout.CancelAfter(McpProtocol.ConnectTimeoutMs);
            var pipe = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut, PipeOptions.Asynchronous, TokenImpersonationLevel.Identification);
            Connection? connection = null;
            try
            {
                await pipe.ConnectAsync(timeout.Token);
                connection = new Connection(pipe);
                _connection = connection;
                _ = ReadAsync(connection);
                var hello = await SendAsync(connection, new McpRequest { Method = "hello", ResumeSecret = _secret }, timeout.Token, false);
                connection.Instance = hello.InstanceId;
            }
            catch (Exception ex) when (ex is IOException or OperationCanceledException or IpcException or UnauthorizedAccessException)
            {
                if (connection != null) Broken(connection); else pipe.Dispose();
                if (cancellation.IsCancellationRequested) throw new OperationCanceledException(cancellation);
                if (ex is IpcException) throw;
                throw new IpcException("plugin_unavailable", "3 秒内未连接插件。请先打开 OneNote，确认 OneNoteCodeHelper 已加载，并检查是否为同一 Windows 用户和登录会话。");
            }
        }
        finally { _connectGate.Release(); }
    }

    internal async Task<McpCatalog> GetCatalogAsync(CancellationToken cancellation)
    {
        await ConnectAsync(cancellation);
        var connection = _connection ?? throw Lost();
        var reply = await SendAsync(connection, new McpRequest { Method = "catalog" }, cancellation);
        return JsonSerializer.Deserialize<McpCatalog>(reply.ResultJson) ?? throw new IpcException("invalid_catalog", "插件工具目录无效。");
    }

    internal async Task<string> CallAsync(string name, string arguments, CancellationToken cancellation)
    {
        await ConnectAsync(cancellation);
        var connection = _connection ?? throw Lost();
        using var parsed = JsonDocument.Parse(arguments);
        if (parsed.RootElement.TryGetProperty("snapshot_id", out var snapshot) && snapshot.ValueKind == JsonValueKind.String &&
            _snapshots.TryGetValue(snapshot.GetString()!, out var instance) && instance != connection.Instance)
            throw new IpcException("snapshot_expired", "插件已重新启动，旧快照失效；不能重发旧草稿，请重新 begin_edit。");
        var reply = await SendAsync(connection, new McpRequest { Method = "call", ToolName = name, ArgumentsJson = arguments }, cancellation);
        if (name == "begin_edit")
        {
            using var result = JsonDocument.Parse(reply.ResultJson);
            _snapshots[result.RootElement.GetProperty("snapshot_id").GetString()!] = connection.Instance;
        }
        return reply.ResultJson;
    }

    private async Task<McpResponse> SendAsync(Connection connection, McpRequest request, CancellationToken cancellation, bool sendCancellation = true)
    {
        request.RequestId = Guid.NewGuid().ToString("N"); request.ClientId = _client; request.InstanceId = connection.Instance;
        var completion = new TaskCompletionSource<McpResponse>(TaskCreationOptions.RunContinuationsAsynchronously);
        connection.Pending[request.RequestId] = completion;
        try
        {
            await WriteAsync(connection, request, cancellation);
            using var registration = cancellation.Register(() =>
            {
                completion.TrySetCanceled(cancellation);
                if (sendCancellation) _ = CancelAsync(connection, request.RequestId);
            });
            var reply = await completion.Task;
            if (reply.ProtocolVersion != McpProtocol.Version) throw new IpcException("protocol_mismatch", "IPC 协议版本不匹配，请一起更新插件和 MCP 程序。");
            if (!reply.Success) throw new IpcException(reply.ErrorCode ?? "operation_failed", reply.ErrorMessage ?? "插件操作失败。");
            return reply;
        }
        finally { connection.Pending.TryRemove(request.RequestId, out _); }
    }

    private async Task CancelAsync(Connection connection, string requestId)
    {
        try
        {
            using var timeout = new CancellationTokenSource(McpProtocol.ConnectTimeoutMs);
            await WriteAsync(connection, new McpRequest { RequestId = Guid.NewGuid().ToString("N"), ClientId = _client, InstanceId = connection.Instance,
                Method = "cancel", CancelRequestId = requestId }, timeout.Token);
        }
        catch (Exception) { /* 断线也会取消未提交草稿；不重放写入。 */ }
    }

    private async Task WriteAsync(Connection connection, McpRequest request, CancellationToken cancellation)
    {
        await connection.Writes.WaitAsync(cancellation);
        try { await McpProtocol.WriteAsync(connection.Pipe, JsonSerializer.Serialize(request), cancellation); }
        catch (Exception)
        {
            // 半帧不可恢复，也不能猜测写入结果；下一次调用只能重新连接并查询状态。
            Broken(connection); throw new IpcException("connection_lost", "插件连接中断，当前请求结果未知；请查询编辑状态，禁止自动重发写入。");
        }
        finally { connection.Writes.Release(); }
    }

    private async Task ReadAsync(Connection connection)
    {
        try
        {
            while (true)
            {
                var json = await McpProtocol.ReadAsync(connection.Pipe, CancellationToken.None);
                if (json == null) break;
                var reply = JsonSerializer.Deserialize<McpResponse>(json) ?? throw new IOException("无效响应。");
                if (reply.RequestId != null && connection.Pending.TryGetValue(reply.RequestId, out var completion)) completion.TrySetResult(reply);
            }
        }
        catch (Exception) { /* 仅向调用者报告安全诊断，不记录正文或凭据。 */ }
        finally { Broken(connection); }
    }

    private void Broken(Connection connection)
    {
        Interlocked.CompareExchange(ref _connection, null, connection);
        connection.Pipe.Dispose();
        foreach (var completion in connection.Pending.Values) completion.TrySetException(Lost());
    }

    private static IpcException Lost() => new("connection_lost", "插件连接中断，当前请求结果未知；请查询编辑状态，禁止自动重发写入。");

    public ValueTask DisposeAsync()
    {
        _disposed = true;
        var connection = Interlocked.Exchange(ref _connection, null);
        if (connection != null) Broken(connection);
        return ValueTask.CompletedTask;
    }
}
