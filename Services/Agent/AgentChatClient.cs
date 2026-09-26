using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace OneNoteCodeHelper.Services.Agent
{
    internal sealed class AgentToolCall
    {
        internal string Id = "";
        internal string Name = "";
        internal string Arguments = "";
        internal object ToMessage() => new { id = Id, type = "function", function = new { name = Name, arguments = Arguments } };
    }

    internal sealed class AgentReply
    {
        internal string Content = "";
        internal string Reasoning = "";
        internal string FinishReason;
        internal bool Done;
        internal readonly SortedDictionary<int, AgentToolCall> Calls = new SortedDictionary<int, AgentToolCall>();
        internal object ToMessage(bool replayReasoning)
        {
            var message = new Dictionary<string, object> { ["role"] = "assistant", ["content"] = Content.Length == 0 ? null : Content };
            if (Calls.Count > 0) message["tool_calls"] = Calls.Values.Select(c => c.ToMessage()).ToArray();
            if (replayReasoning) message["reasoning_content"] = Reasoning;
            return message;
        }
        internal void Validate()
        {
            if (FinishReason != "stop" && FinishReason != "tool_calls") throw new AiException("Agent 输出未完整结束，未执行本轮工具。");
            if (Calls.Count > 0 && FinishReason != "tool_calls") throw new AiException("工具调用结束标记无效。");
            if (FinishReason == "tool_calls" && Calls.Count == 0) throw new AiException("接口没有返回工具调用。");
            if (Calls.Values.Any(c => string.IsNullOrEmpty(c.Id) || string.IsNullOrEmpty(c.Name) || string.IsNullOrWhiteSpace(c.Arguments)) ||
                Calls.Values.Select(c => c.Id).Distinct().Count() != Calls.Count) throw new AiException("工具调用格式不完整。");
        }
    }

    internal interface IAgentChatClient
    {
        Task<AgentReply> CompleteAsync(List<object> messages, object[] tools, IProgress<string> progress, CancellationToken cancellation);
    }

    internal sealed class AgentChatClient : IAgentChatClient
    {
        private readonly AiConfig _config;
        private readonly string _model;
        private readonly string _effort;
        private readonly HttpClient _http;
        internal AgentChatClient(AiConfig config, string model, string effort, HttpClient http = null)
        { _config = config; _model = model; _effort = effort; _http = http ?? AiClient.Transport; }
        internal static JavaScriptSerializer Serializer() => new JavaScriptSerializer { MaxJsonLength = 600000, RecursionLimit = 40 };
        internal static object Parse(string json)
        {
            try { return Serializer().DeserializeObject(json); }
            catch (Exception) { throw new AiException("Agent 返回了无效或超限的 JSON。"); }
        }

        public async Task<AgentReply> CompleteAsync(List<object> messages, object[] tools, IProgress<string> progress, CancellationToken cancellation)
        {
            var body = new Dictionary<string, object> { ["model"] = _model, ["messages"] = messages, ["tools"] = tools,
                ["tool_choice"] = "auto", ["stream"] = true };
            if (_config.Agent.StreamUsage) body["stream_options"] = new { include_usage = true };
            if (_config.Agent.SendThinking) AiEfforts.ApplyTo(body, _effort);
            if (_config.MaxTokens > 0) body["max_tokens"] = _config.MaxTokens;
            var json = Serializer().Serialize(body);
            if (json.Length > _config.Agent.MaxRequestChars) throw new AiException("Agent 会话达到上下文预算，未提交草稿。请缩小处理范围。");
            using (var total = new CancellationTokenSource(TimeSpan.FromSeconds(_config.TimeoutSeconds)))
            using (var idle = new CancellationTokenSource(TimeSpan.FromSeconds(AiClient.IdleTimeoutSeconds)))
            using (var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellation, total.Token, idle.Token))
            using (var request = new HttpRequestMessage(HttpMethod.Post, AiClient.BuildEndpoint(_config.ApiUrl)))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _config.ApiKey);
                request.Content = new StringContent(json, Encoding.UTF8, "application/json");
                try
                {
                    using (var response = await _http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, linked.Token).ConfigureAwait(false))
                    using (linked.Token.Register(response.Dispose))
                    {
                        if (!response.IsSuccessStatusCode) throw new AiException($"Agent 接口返回 HTTP {(int)response.StatusCode}；请检查模型是否支持工具调用及 AI 配置。");
                        var reply = new AgentReply();
                        var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
                        using (var reader = new StreamReader(stream, Encoding.UTF8))
                        {
                            if (string.Equals(response.Content.Headers.ContentType?.MediaType, "text/event-stream", StringComparison.OrdinalIgnoreCase))
                            {
                                var eventData = new StringBuilder();
                                var received = 0;
                                string line;
                                while (!reply.Done && (line = await reader.ReadLineAsync().ConfigureAwait(false)) != null)
                                {
                                    idle.CancelAfter(TimeSpan.FromSeconds(AiClient.IdleTimeoutSeconds));
                                    received += line.Length;
                                    if (received > 600000) throw new AiException("Agent 响应超过大小限制。");
                                    if (line.Length == 0)
                                    {
                                        if (eventData.Length > 0) AbsorbEvent(eventData.ToString(), reply);
                                        eventData.Clear();
                                        progress?.Report($"模型处理中：已接收 {reply.Reasoning.Length} 个思考字符，{reply.Content.Length + reply.Calls.Values.Sum(c => c.Arguments.Length)} 个输出字符");
                                    }
                                    else if (line.StartsWith("data:", StringComparison.Ordinal))
                                    {
                                        if (eventData.Length > 0) eventData.Append('\n');
                                        eventData.Append(line.Substring(5).TrimStart());
                                    }
                                }
                                if (eventData.Length > 0 && !reply.Done) AbsorbEvent(eventData.ToString(), reply);
                            }
                            else
                            {
                                var text = new StringBuilder();
                                var buffer = new char[4096];
                                int count;
                                while ((count = await reader.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false)) > 0)
                                {
                                    idle.CancelAfter(TimeSpan.FromSeconds(AiClient.IdleTimeoutSeconds));
                                    text.Append(buffer, 0, count);
                                    if (text.Length > 600000) throw new AiException("Agent 响应超过大小限制。");
                                }
                                AbsorbWhole(Parse(text.ToString()), reply);
                            }
                        }
                        reply.Validate();
                        return reply;
                    }
                }
                catch (Exception) when (linked.IsCancellationRequested)
                {
                    cancellation.ThrowIfCancellationRequested();
                    throw new AiException(total.IsCancellationRequested ? "Agent 接口请求超时，未提交草稿。" : "Agent 接口 60 秒无新数据，未提交草稿。");
                }
                catch (HttpRequestException) { throw new AiException("无法连接 Agent 接口，未提交草稿。"); }
                catch (IOException) { throw new AiException("Agent 连接中断，未提交草稿。"); }
            }
        }

        internal static void AbsorbEvent(string data, AgentReply reply)
        {
            if (data.Trim() == "[DONE]") { reply.Done = true; return; }
            var root = Parse(data);
            if (AiClient.Get(root, "error") != null) throw new AiException("Agent 接口在流中返回错误。");
            if (!(AiClient.Get(root, "choices") is IList choices)) throw new AiException("Agent 流数据缺少 choices。");
            foreach (var choice in choices)
            {
                if (Convert.ToInt32(AiClient.Get(choice, "index") ?? 0) != 0) throw new AiException("Agent 只接受一个模型候选结果。");
                var delta = AiClient.Get(choice, "delta");
                reply.Content += AiClient.Get(delta, "content") as string ?? "";
                reply.Reasoning += AiClient.Get(delta, "reasoning_content") as string ?? "";
                reply.FinishReason = AiClient.Get(choice, "finish_reason") as string ?? reply.FinishReason;
                if (!(AiClient.Get(delta, "tool_calls") is IList calls)) continue;
                foreach (var call in calls)
                {
                    var rawIndex = AiClient.Get(call, "index");
                    if (!(rawIndex is int index) || index < 0 || index > 47) throw new AiException("工具片段索引无效。");
                    if (!reply.Calls.TryGetValue(index, out var value)) reply.Calls[index] = value = new AgentToolCall();
                    var id = AiClient.Get(call, "id") as string;
                    var name = AiClient.Get(AiClient.Get(call, "function"), "name") as string;
                    if (id != null) { if (value.Id.Length > 0 && value.Id != id) throw new AiException("工具 ID 在流中改变。"); value.Id = id; }
                    if (name != null) { if (value.Name.Length > 0 && value.Name != name) throw new AiException("工具名称在流中改变。"); value.Name = name; }
                    var type = AiClient.Get(call, "type") as string;
                    if (type != null && type != "function") throw new AiException("未知工具类型。");
                    value.Arguments += AiClient.Get(AiClient.Get(call, "function"), "arguments") as string ?? "";
                    if (value.Arguments.Length > 64000) throw new AiException("工具参数过大。");
                }
            }
        }
        private static void AbsorbWhole(object root, AgentReply reply)
        {
            var choice = AiClient.FirstItem(AiClient.Get(root, "choices"));
            var message = AiClient.Get(choice, "message");
            reply.Content = AiClient.Get(message, "content") as string ?? "";
            reply.Reasoning = AiClient.Get(message, "reasoning_content") as string ?? "";
            reply.FinishReason = AiClient.Get(choice, "finish_reason") as string;
            if (AiClient.Get(message, "tool_calls") is IList calls)
                foreach (var c in calls) reply.Calls.Add(reply.Calls.Count, new AgentToolCall
                { Id = AiClient.Get(c, "id") as string, Name = AiClient.Get(AiClient.Get(c, "function"), "name") as string,
                    Arguments = AiClient.Get(AiClient.Get(c, "function"), "arguments") as string });
            reply.Done = true;
        }
    }
}
