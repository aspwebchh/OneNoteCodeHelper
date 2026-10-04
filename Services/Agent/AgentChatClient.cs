using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
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
    /// <summary>
    /// Agent 请求体和工具结果的 JSON。JavaScriptSerializer 把 &lt; &gt; &amp; ' 转成 < 这样的 6 个字符，工具结果作为字符串再序列化一次又变成 7 个；
    /// JSON 本来不要求转义它们，这里还原成原字符，模型看到的是原文，上下文预算也只算实际内容。
    /// </summary>
    internal static class AgentJson
    {
        internal static string Serialize(object value) => Unescape(AgentChatClient.Serializer().Serialize(value));

        /// <summary>工具结果：去掉值为 null 的字段，空数组保留。</summary>
        internal static string ToolResult(object outcome)
        {
            var serializer = AgentChatClient.Serializer();
            return Serialize(Prune(serializer.DeserializeObject(serializer.Serialize(outcome))));
        }

        private static object Prune(object value)
        {
            if (value is IDictionary<string, object> map) return map.Where(p => p.Value != null).ToDictionary(p => p.Key, p => Prune(p.Value));
            if (value is object[] list) return list.Select(Prune).ToArray();
            return value;
        }

        private static string Unescape(string json)
        {
            if (json.IndexOf("\\u00", StringComparison.Ordinal) < 0) return json;
            var result = new StringBuilder(json.Length);
            for (var i = 0; i < json.Length; i++)
            {
                var ch = json[i];
                if (ch != '\\' || i + 1 >= json.Length) { result.Append(ch); continue; }
                if (json[i + 1] == 'u' && i + 5 < json.Length)
                {
                    var plain = Plain(json.Substring(i + 2, 4));
                    if (plain != '\0') { result.Append(plain); i += 5; continue; }
                }
                // 其他转义连同下一个字符原样保留，免得把 \\u003c（原文里的反斜杠）后面的 u003c 当成转义。
                result.Append(ch).Append(json[i + 1]);
                i++;
            }
            return result.ToString();
        }

        private static char Plain(string hex)
        {
            switch (hex.ToLowerInvariant())
            {
                case "003c": return '<';
                case "003e": return '>';
                case "0026": return '&';
                case "0027": return '\'';
                default: return '\0';
            }
        }
    }

    internal sealed class AgentToolCall
    {
        internal string Id = "";
        internal string Name = "";
        internal readonly StringBuilder ArgumentsBuffer = new StringBuilder();
        internal string Arguments { get => ArgumentsBuffer.ToString(); set { ArgumentsBuffer.Clear(); ArgumentsBuffer.Append(value ?? ""); } }
        internal int ArgumentsLength => ArgumentsBuffer.Length;
        internal object ToMessage() => new { id = Id, type = "function", function = new { name = Name, arguments = Arguments } };
    }

    internal sealed class AgentReply
    {
        internal readonly StringBuilder ContentBuffer = new StringBuilder();
        internal readonly StringBuilder ReasoningBuffer = new StringBuilder();
        internal string Content { get => ContentBuffer.ToString(); set { ContentBuffer.Clear(); ContentBuffer.Append(value ?? ""); } }
        internal string Reasoning { get => ReasoningBuffer.ToString(); set { ReasoningBuffer.Clear(); ReasoningBuffer.Append(value ?? ""); } }
        internal int ContentLength => ContentBuffer.Length;
        internal int ReasoningLength => ReasoningBuffer.Length;
        internal string FinishReason;
        internal bool Done;
        /// <summary>接口返回的 usage，只用来写日志；接口不给时为 null。</summary>
        internal object Usage;
        /// <summary>输出达到 max_tokens 被截断。本轮工具不能执行，但草稿没动，由 <see cref="AgentRunner"/> 决定是否提醒模型分批重试。</summary>
        internal bool Truncated => FinishReason == "length";
        internal readonly SortedDictionary<int, AgentToolCall> Calls = new SortedDictionary<int, AgentToolCall>();
        internal object ToMessage()
        {
            var message = new Dictionary<string, object> { ["role"] = "assistant", ["content"] = ContentLength == 0 ? null : Content };
            if (Calls.Count > 0) message["tool_calls"] = Calls.Values.Select(c => c.ToMessage()).ToArray();
            message["reasoning_content"] = Reasoning;
            return message;
        }
        internal void Validate()
        {
            if (FinishReason != "stop" && FinishReason != "tool_calls") throw new AiException(Unfinished());
            if (Calls.Count > 0 && FinishReason != "tool_calls") throw new AiException("工具调用结束标记无效。");
            if (FinishReason == "tool_calls" && Calls.Count == 0) throw new AiException("接口没有返回工具调用。");
            if (Calls.Values.Any(c => string.IsNullOrEmpty(c.Id) || string.IsNullOrEmpty(c.Name) || string.IsNullOrWhiteSpace(c.Arguments)) ||
                Calls.Values.Select(c => c.Id).Distinct().Count() != Calls.Count) throw new AiException("工具调用格式不完整。");
        }

        /// <summary>没有正常结束时按结束原因说明，用户才知道该调哪一项、还是重试就行。</summary>
        private string Unfinished()
        {
            switch (FinishReason)
            {
                case "length":
                    return "模型本轮输出达到「最大输出 token」上限被截断，未执行本轮工具。可以在「AI 配置」调大「最大输出 token」（MaxTokens，0 表示用接口默认值）、" +
                        "降低思考强度，或缩小处理范围。";
                case "content_filter": return "模型输出被接口的内容审核拦截，未执行本轮工具。";
                case "insufficient_system_resource": return "接口资源不足，模型输出被中断，未执行本轮工具。请稍后重试。";
                case null: return Done ? "接口没有返回结束原因，未执行本轮工具。" : "Agent 连接中途结束，没有收到完整回复，未执行本轮工具。请重试。";
                default: return $"Agent 输出没有正常结束（结束原因 {Reason(FinishReason)}），未执行本轮工具。";
            }
        }

        /// <summary>结束原因来自接口，显示和写日志前只留短的英文标识。</summary>
        internal static string Reason(string finish) =>
            finish == null ? "-" : finish.Length <= 40 && finish.All(c => c < 128 && (char.IsLetterOrDigit(c) || c == '_' || c == '-')) ? finish : "未知";
    }

    internal interface IAgentChatClient
    {
        Task<AgentReply> CompleteAsync(List<object> messages, object[] tools, IProgress<AgentProgress> progress, CancellationToken cancellation);
    }

    internal sealed class AgentChatClient : IAgentChatClient
    {
        // SSE 每个小片段都会重复 JSON 字段名；传输字数不能当成模型输出字数。
        internal const int MaxStreamWireChars = 8 * 1024 * 1024;
        internal const int MaxReplyChars = 500000;
        internal const int MaxToolArgumentsChars = 64000;
        /// <summary>错误响应仅用于识别固定错误码，限制大小和等待时间，不回显或记录上游正文。</summary>
        private const int MaxHttpErrorChars = 16384;
        /// <summary>序列化上限，须大于 MaxRequestChars 的上限（300 万字）再加一轮的增量；真正的预算由 MaxRequestChars 控制。</summary>
        internal const int MaxJsonChars = 16 * 1024 * 1024;
        /// <summary>流式进度最多每隔这么久（毫秒）报告一次，免得把界面线程的消息队列塞满。</summary>
        private const int ReportIntervalMs = 200;
        private readonly AiConfig _config;
        private readonly string _model;
        private readonly string _effort;
        private readonly HttpClient _http;
        internal AgentChatClient(AiConfig config, string model, string effort, HttpClient http = null)
        { _config = config; _model = model; _effort = effort; _http = http ?? AiClient.Transport; }
        internal static JavaScriptSerializer Serializer() => new JavaScriptSerializer { MaxJsonLength = MaxJsonChars, RecursionLimit = 40 };
        internal static object Parse(string json)
        {
            try { return Serializer().DeserializeObject(json); }
            catch (Exception) { throw new AiException("Agent 返回了无效或超限的 JSON。"); }
        }

        public async Task<AgentReply> CompleteAsync(List<object> messages, object[] tools, IProgress<AgentProgress> progress, CancellationToken cancellation)
        {
            var body = new Dictionary<string, object> { ["model"] = _model, ["messages"] = messages, ["tools"] = tools,
                ["tool_choice"] = "auto", ["stream"] = true, ["stream_options"] = new { include_usage = true } };
            AiEfforts.ApplyTo(body, _effort);
            if (_config.MaxTokens > 0) body["max_tokens"] = _config.MaxTokens;
            var json = AgentJson.Serialize(body);
            if (json.Length > _config.Agent.MaxRequestChars)
                throw new AiException($"Agent 会话达到上下文预算（约 {json.Length} 字，上限 {_config.Agent.MaxRequestChars}），未提交草稿。可以缩小处理范围，或在「AI 配置」的 Agent 页调大「请求字符上限」（MaxRequestChars）。");
            using (var total = new CancellationTokenSource(TimeSpan.FromSeconds(_config.TimeoutSeconds)))
            using (var idle = new CancellationTokenSource(TimeSpan.FromSeconds(AiClient.IdleTimeoutSeconds)))
            using (var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellation, total.Token, idle.Token))
            using (var request = new HttpRequestMessage(HttpMethod.Post, AiClient.BuildEndpoint(_config.ApiUrl)))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _config.ApiKey);
                request.Content = new StringContent(json, Encoding.UTF8, "application/json");
                var watch = Stopwatch.StartNew();
                try
                {
                    using (var response = await _http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, linked.Token).ConfigureAwait(false))
                    using (linked.Token.Register(response.Dispose))
                    {
                        if (!response.IsSuccessStatusCode)
                        {
                            var status = (int)response.StatusCode;
                            var code = await ReadHttpErrorCode(response, linked.Token).ConfigureAwait(false);
                            AddInLog.Info($"Agent HTTP 失败：status={status} code={code ?? "unknown"}，耗时 {watch.Elapsed.TotalSeconds:0.0}s。");
                            throw new AiException(DescribeHttpError(status, code));
                        }
                        var reply = new AgentReply();
                        var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
                        using (var reader = new StreamReader(stream, Encoding.UTF8))
                        {
                            if (string.Equals(response.Content.Headers.ContentType?.MediaType, "text/event-stream", StringComparison.OrdinalIgnoreCase))
                            {
                                var eventData = new StringBuilder();
                                var received = 0;
                                Stopwatch sinceReport = null;
                                string line;
                                while (!reply.Done && (line = await reader.ReadLineAsync().ConfigureAwait(false)) != null)
                                {
                                    idle.CancelAfter(TimeSpan.FromSeconds(AiClient.IdleTimeoutSeconds));
                                    received += line.Length;
                                    if (received > MaxStreamWireChars)
                                        throw new AiException("Agent 流式传输超过 8 MB，未执行本轮工具。请缩小处理范围或降低思考强度。");
                                    if (line.Length == 0)
                                    {
                                        if (eventData.Length > 0) AbsorbEvent(eventData.ToString(), reply);
                                        eventData.Clear();
                                        if (progress != null && (sinceReport == null || sinceReport.ElapsedMilliseconds >= ReportIntervalMs))
                                        {
                                            progress.Report(DescribeStream(reply));
                                            sinceReport = Stopwatch.StartNew();
                                        }
                                    }
                                    else if (line.StartsWith("data:", StringComparison.Ordinal))
                                    {
                                        if (eventData.Length > 0) eventData.Append('\n');
                                        eventData.Append(line.Substring(5).TrimStart());
                                        if (eventData.Length > 600000) throw new AiException("Agent 单条流事件超过大小限制，未执行本轮工具。");
                                    }
                                }
                                if (eventData.Length > 0 && !reply.Done) AbsorbEvent(eventData.ToString(), reply);
                                // 限频可能吞掉最后几段，收完再报一次最终状态。
                                progress?.Report(DescribeStream(reply));
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
                                    if (text.Length > 600000) throw new AiException("Agent 非流式响应超过大小限制，未执行本轮工具。");
                                }
                                AbsorbWhole(Parse(text.ToString()), reply);
                            }
                        }
                        AddInLog.Info(Summarize(reply, _model, _effort, json.Length, watch.Elapsed.TotalSeconds));
                        // 截断的一轮交给 AgentRunner：草稿没动，可以提醒模型分批后重试；其余情况在这里就拒绝。
                        if (!reply.Truncated) reply.Validate();
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

        private static async Task<string> ReadHttpErrorCode(HttpResponseMessage response, CancellationToken cancellation)
        {
            if (response.Content == null) return null;
            using (var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(3)))
            using (var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellation, timeout.Token))
            using (linked.Token.Register(response.Dispose))
            {
                try
                {
                    using (var reader = new StreamReader(await response.Content.ReadAsStreamAsync().ConfigureAwait(false), Encoding.UTF8))
                    {
                        var text = new StringBuilder();
                        var buffer = new char[2048];
                        int count;
                        while ((count = await reader.ReadAsync(buffer, 0, Math.Min(buffer.Length, MaxHttpErrorChars + 1 - text.Length)).ConfigureAwait(false)) > 0)
                        {
                            linked.Token.ThrowIfCancellationRequested();
                            text.Append(buffer, 0, count);
                            if (text.Length > MaxHttpErrorChars) return null;
                        }
                        var code = AiClient.Get(AiClient.Get(Parse(text.ToString()), "error"), "code") as string;
                        return string.Equals(code, "model_not_found", StringComparison.OrdinalIgnoreCase) ? "model_not_found" : null;
                    }
                }
                catch (Exception)
                {
                    cancellation.ThrowIfCancellationRequested();
                    // 非 JSON、过大、读取失败或错误正文迟迟不结束，仍按已经收到的 HTTP 状态提示。
                    return null;
                }
            }
        }

        private static string DescribeHttpError(int status, string code)
        {
            var prefix = $"Agent 接口返回 HTTP {status}；";
            string reason;
            if (code == "model_not_found") reason = "当前模型不存在或未开通，请检查「AI 配置」里的模型名及网关模型通道。";
            else switch (status)
            {
                case 401:
                case 403: reason = "接口拒绝访问，请检查「AI 配置」里的 ApiKey 和访问权限。"; break;
                case 404: reason = "请求地址不存在，请检查「AI 配置」里的 ApiUrl。"; break;
                case 429: reason = "接口限流或额度不足，请稍后再试，或检查接口额度。"; break;
                case 500:
                case 502:
                case 503: reason = "网关或模型服务暂时不可用，请稍后再试；持续失败时请检查网关服务及模型通道。"; break;
                case 504: reason = "网关等待上游服务超时，请稍后再试。"; break;
                case 400:
                case 422: reason = "接口拒绝请求参数，请检查模型的工具调用支持及 Agent 兼容选项。"; break;
                default: reason = "请求失败，请检查接口服务和 AI 配置。"; break;
            }
            return prefix + reason + "未提交草稿。";
        }

        /// <summary>每轮请求一行日志：结束原因、用量和各部分字数，不含模型原文、工具参数或笔记内容。</summary>
        internal static string Summarize(AgentReply reply, string model, string effort, int requestChars, double seconds) =>
            $"Agent 请求：model={model} effort={effort} 请求 {requestChars} 字，耗时 {seconds:0.0}s，finish={AgentReply.Reason(reply.FinishReason)}，" +
            $"token 输入/输出/思考 = {AiClient.Get(reply.Usage, "prompt_tokens") ?? "-"}/{AiClient.Get(reply.Usage, "completion_tokens") ?? "-"}/" +
            $"{AiClient.Get(AiClient.Get(reply.Usage, "completion_tokens_details"), "reasoning_tokens") ?? "-"}，" +
            $"思考 {reply.ReasoningLength} 字，回复 {reply.ContentLength} 字，工具 {reply.Calls.Count} 个（参数 {reply.Calls.Values.Sum(c => c.ArgumentsLength)} 字）";

        /// <summary>
        /// 流式返回期间的进度：模型在准备哪个工具、在回复还是在思考，以及回复（没有时是思考）里最新的几句核心内容，
        /// 工具名、段落 ID 换成中文说法。
        /// 还没说完新的一句时 Thinking 为 null，摘录框留着上一次的。
        /// </summary>
        internal static AgentProgress DescribeStream(AgentReply reply)
        {
            var call = reply.Calls.Values.LastOrDefault(c => !string.IsNullOrEmpty(c.Name));
            var status = call != null ? "模型正在准备：" + AgentTools.DisplayName(call.Name)
                : reply.ContentLength > 0 ? "模型正在回复…"
                : reply.ReasoningLength > 0 ? "模型正在思考…"
                : "等待模型响应…";
            // 开始回复或调用工具后思考就结束了，最后一句没有句末标点也算说完。
            var answered = reply.Done || reply.FinishReason != null || reply.Calls.Count > 0;
            var thinking = (reply.ContentLength > 0 ? LiveText.Gist(reply.ContentBuffer, answered, AgentTools.LocalizeThought) : null)
                           ?? LiveText.Gist(reply.ReasoningBuffer, reply.ContentLength > 0 || answered, AgentTools.LocalizeThought);
            return new AgentProgress { Status = status, Thinking = thinking };
        }

        internal static void AbsorbEvent(string data, AgentReply reply)
        {
            if (data.Trim() == "[DONE]") { reply.Done = true; return; }
            var root = Parse(data);
            if (AiClient.Get(root, "error") != null) throw new AiException("Agent 接口在流中返回错误。");
            var usage = AiClient.Get(root, "usage");
            if (usage != null) reply.Usage = usage;
            if (!(AiClient.Get(root, "choices") is IList choices))
            {
                // 有的网关把用量单独发一条，不带 choices。
                if (usage != null) return;
                throw new AiException("Agent 流数据缺少 choices。");
            }
            foreach (var choice in choices)
            {
                if (Convert.ToInt32(AiClient.Get(choice, "index") ?? 0) != 0) throw new AiException("Agent 只接受一个模型候选结果。");
                var delta = AiClient.Get(choice, "delta");
                AppendReply(reply.ContentBuffer, AiClient.Get(delta, "content") as string, reply);
                AppendReply(reply.ReasoningBuffer, AiClient.Get(delta, "reasoning_content") as string, reply);
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
                    AppendReply(value.ArgumentsBuffer, AiClient.Get(AiClient.Get(call, "function"), "arguments") as string, reply);
                    if (value.ArgumentsLength > MaxToolArgumentsChars) throw new AiException("工具参数过大。");
                }
            }
        }
        private static void AppendReply(StringBuilder target, string fragment, AgentReply reply)
        {
            if (string.IsNullOrEmpty(fragment)) return;
            var current = reply.ContentLength + reply.ReasoningLength + reply.Calls.Values.Sum(c => c.ArgumentsLength);
            if (current + fragment.Length > MaxReplyChars)
                throw new AiException("Agent 本轮有效输出超过 50 万字符，未执行本轮工具。请缩小处理范围或降低思考强度。");
            target.Append(fragment);
        }
        private static void AbsorbWhole(object root, AgentReply reply)
        {
            var choice = AiClient.FirstItem(AiClient.Get(root, "choices"));
            var message = AiClient.Get(choice, "message");
            reply.Content = AiClient.Get(message, "content") as string ?? "";
            reply.Reasoning = AiClient.Get(message, "reasoning_content") as string ?? "";
            reply.FinishReason = AiClient.Get(choice, "finish_reason") as string;
            reply.Usage = AiClient.Get(root, "usage");
            if (AiClient.Get(message, "tool_calls") is IList calls)
                foreach (var c in calls) reply.Calls.Add(reply.Calls.Count, new AgentToolCall
                { Id = AiClient.Get(c, "id") as string, Name = AiClient.Get(AiClient.Get(c, "function"), "name") as string,
                    Arguments = AiClient.Get(AiClient.Get(c, "function"), "arguments") as string });
            if (reply.ContentLength + reply.ReasoningLength + reply.Calls.Values.Sum(c => c.ArgumentsLength) > MaxReplyChars)
                throw new AiException("Agent 本轮有效输出超过 50 万字符，未执行本轮工具。请缩小处理范围或降低思考强度。");
            if (reply.Calls.Values.Any(c => c.ArgumentsLength > MaxToolArgumentsChars)) throw new AiException("工具参数过大。");
            reply.Done = true;
        }
    }
}
