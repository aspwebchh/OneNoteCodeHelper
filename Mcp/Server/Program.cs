using System.Text;
using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol;
using ModelContextProtocol.Server;
using ModelContextProtocol.Protocol;

namespace OneNoteCodeHelper.Mcp;

internal static class Program
{
    private static async Task<int> Main(string[] args)
    {
        Console.InputEncoding = new UTF8Encoding(false);
        Console.OutputEncoding = new UTF8Encoding(false);
        var doctor = false;
        var pipe = McpProtocol.DefaultPipeName;
        for (var i = 0; i < args.Length; i++)
        {
            if (args[i] == "--doctor") doctor = true;
            else if (args[i] == "--pipe-name" && i + 1 < args.Length) pipe = args[++i];
            else { await Console.Error.WriteLineAsync("用法：OneNoteCodeHelper.Mcp.exe [--doctor] [--pipe-name <本机管道名称>]"); return 2; }
        }
        await using var ipc = new IpcClient(pipe);
        try
        {
            // 在接管 stdio 前确认插件就绪；无连接时明确退出，不留下不可用的 MCP 会话。
            await ipc.ConnectAsync(CancellationToken.None);
            var catalog = await ipc.GetCatalogAsync(CancellationToken.None);
            if (doctor)
            {
                var status = await ipc.CallAsync("get_status", "{}", CancellationToken.None);
                await Console.Out.WriteLineAsync(status);
                return 0;
            }
            var advertised = catalog.Tools.Select(t => new Tool
            {
                Name = t.Name, Description = t.Description, InputSchema = ParseElement(t.InputSchemaJson),
                OutputSchema = t.OutputSchemaJson == null ? null : ParseElement(t.OutputSchemaJson),
                Annotations = new ToolAnnotations { ReadOnlyHint = t.ReadOnly, DestructiveHint = t.Destructive,
                    IdempotentHint = t.ReadOnly || t.Name is "finish_edit" or "undo_edit" or "abort_edit" or "navigate_to" or "create_page" or "copy_page" or "move_page" or "create_section" or "append_content" or "insert_content" or "replace_text", OpenWorldHint = false }
            }).ToList();
            var builder = Host.CreateApplicationBuilder();
            // SDK 的异常日志可能带工具参数；诊断由本程序输出安全的错误码和状态。
            builder.Logging.ClearProviders();
            builder.Services.Configure<HostOptions>(options => options.ShutdownTimeout = TimeSpan.FromSeconds(3));
            builder.Services.AddMcpServer(options =>
            {
                options.ServerInfo = new Implementation { Name = "OneNoteCodeHelper", Version = "1.1.0" };
                options.ServerInstructions = catalog.Instructions;
            }).WithStdioServerTransport()
                .WithListToolsHandler((_, _) => ValueTask.FromResult(new ListToolsResult { Tools = advertised }))
                .WithCallToolHandler(async (context, cancellation) =>
                {
                    var request = context.Params!;
                    try
                    {
                        using var progressStop = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
                        var progress = request.ProgressToken is { } token ? ReportProgressAsync(context.Server, token, progressStop.Token) : Task.CompletedTask;
                        try
                        {
                            var json = await ipc.CallAsync(request.Name, request.Arguments == null ? "{}" : JsonSerializer.Serialize(request.Arguments), cancellation);
                            return Result(json);
                        }
                        finally { progressStop.Cancel(); await progress; }
                    }
                    catch (IpcException ex) { return Result(JsonSerializer.Serialize(new { ok = false, error_code = ex.Code, error = ex.Message }), true); }
                    catch (JsonException) { return Result("{\"ok\":false,\"error_code\":\"invalid_arguments\",\"error\":\"工具参数不是有效 JSON 对象。\"}", true); }
                });
            using var host = builder.Build();
            await host.RunAsync();
            return 0;
        }
        catch (IpcException ex) { await Console.Error.WriteLineAsync($"OneNote MCP [{ex.Code}]: {ex.Message}"); return 1; }
        catch (Exception ex) { await Console.Error.WriteLineAsync($"OneNote MCP 启动或传输失败：{ex.GetType().Name}"); return 3; }
    }

    private static JsonElement ParseElement(string json)
    {
        using var document = JsonDocument.Parse(json);
        return document.RootElement.Clone();
    }

    private static async Task ReportProgressAsync(McpServer server, ProgressToken token, CancellationToken cancellation)
    {
        var elapsed = 0;
        try
        {
            while (true)
            {
                await server.NotifyProgressAsync(token, new ProgressNotificationValue { Progress = elapsed, Message = $"OneNote 正在处理，已用 {elapsed} 秒。" }, null, cancellation);
                await Task.Delay(2000, cancellation); elapsed += 2;
            }
        }
        catch (Exception) { /* 进度失败不得影响提交结果，也不记录参数或正文。 */ }
    }

    private static CallToolResult Result(string json, bool error = false)
    {
        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;
        if (root.ValueKind == JsonValueKind.Object)
        {
            if (root.TryGetProperty("ok", out var ok) && ok.ValueKind == JsonValueKind.False) error = true;
            if (root.TryGetProperty("status", out var status) && status.GetString() == "CommitOutcomeUnknown") error = true;
        }
        return new CallToolResult { Content = [new TextContentBlock { Text = json }], StructuredContent = root.Clone(), IsError = error };
    }
}
