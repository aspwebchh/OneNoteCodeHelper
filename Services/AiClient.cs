using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace OneNoteCodeHelper.Services
{
    /// <summary>调用 AI 接口时的失败。Message 是可以直接给用户看的中文说明。</summary>
    internal sealed class AiException : Exception
    {
        internal AiException(string message, Exception inner = null)
            : base(message, inner)
        {
        }
    }

    /// <summary>一次请求收到的结果。流式返回时一段段拼起来。</summary>
    internal sealed class AiReply
    {
        internal StringBuilder Content { get; } = new StringBuilder();

        internal int ReasoningChars { get; set; }

        internal string FinishReason { get; set; }

        internal object Usage { get; set; }

        /// <summary>收到了流的结束标记 [DONE]。</summary>
        internal bool Done { get; set; }

        /// <summary>从发请求到收到第一段数据用了多少秒，还没收到时为 null。</summary>
        internal double? FirstDataSeconds { get; set; }

        /// <summary>收到了结束标记或 finish_reason，说明结果是完整的。</summary>
        internal bool IsComplete => Done || FinishReason != null;
    }

    /// <summary>
    /// OpenAI 兼容的 Chat Completions 接口的最小封装：一问一答，要求模型输出 JSON 对象。
    /// </summary>
    internal static class AiClient
    {
        /// <summary>
        /// 连续这么多秒一个字节都没收到就认定卡住了。流式返回时思考、输出都是持续一段段到的，
        /// 模型排队时接口也会发保活的注释行，所以这个时间不用跟着思考强度或内容多少调。
        /// </summary>
        internal const int IdleTimeoutSeconds = 60;

        private static readonly TimeSpan IdleTimeout = TimeSpan.FromSeconds(IdleTimeoutSeconds);

        private static readonly HttpClient Http;

        internal static HttpClient Transport => Http;

        static AiClient()
        {
            // 插件跑在 dllhost 里，AppDomain 拿不到目标框架信息，.NET 会按老规矩只开 SSL3/TLS1.0，
            // 现在的 HTTPS 服务一律握手失败。这里显式把 TLS1.2 加上。
            try
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            }
            catch (Exception ex)
            {
                AddInLog.Warn("开启 TLS1.2 失败，HTTPS 请求可能连不上。", ex);
            }

            // 超时按每次请求单独算（见 CompleteAsync），这里不设，免得和配置里的超时打架。
            Http = new HttpClient { Timeout = Timeout.InfiniteTimeSpan };
        }

        internal static JavaScriptSerializer CreateSerializer()
        {
            return new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
        }

        /// <summary>
        /// 发一次请求，返回模型输出的正文（message.content）。
        ///
        /// 走流式：思考和输出是一段段边生成边到的，每到一段回调一次 onChunk（这一段的思考原文、正文原文），
        /// Agent 窗口靠它显示 AI 在想什么、已经返回了几段修改，用户能看出 AI 在干活。
        /// 不走流式的话，整个结果生成完才会有第一个字节，思考得久一点窗口就一动不动，
        /// 接口卡住时也只能干等到总超时；而网关（One API）的日志要等请求结束才记，那段时间在后台也查不到这个请求。
        ///
        /// 连续 <see cref="IdleTimeoutSeconds"/> 秒一个字节都没收到就当卡住了直接报错，不等到配置里的总超时。
        /// 失败一律抛 <see cref="AiException"/>；用户取消抛 <see cref="OperationCanceledException"/>。
        /// </summary>
        internal static async Task<string> CompleteAsync(AiConfig config, string modelId, string effort,
            string systemPrompt, string userContent, Action<string, string> onChunk, CancellationToken cancellation)
        {
            var body = BuildRequestBody(config, modelId, effort, systemPrompt, userContent);
            var serializer = CreateSerializer();
            var watch = Stopwatch.StartNew();
            var reply = new AiReply();
            var responded = false;

            using (var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(config.TimeoutSeconds)))
            using (var idle = new CancellationTokenSource(IdleTimeout))
            using (var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellation, timeout.Token, idle.Token))
            using (var request = new HttpRequestMessage(HttpMethod.Post, BuildEndpoint(config.ApiUrl)))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", config.ApiKey);
                request.Content = new StringContent(serializer.Serialize(body), Encoding.UTF8, "application/json");

                try
                {
                    // 拿到响应头就返回，正文边到边读。
                    using (var response = await Http
                               .SendAsync(request, HttpCompletionOption.ResponseHeadersRead, linked.Token)
                               .ConfigureAwait(false))
                    // 读正文用的 ReadLineAsync 不认取消令牌：取消或超时就直接关掉响应，让卡在读取上的调用马上失败。
                    using (linked.Token.Register(response.Dispose))
                    {
                        responded = true;
                        idle.CancelAfter(IdleTimeout);

                        if (!response.IsSuccessStatusCode)
                        {
                            var text = await ReadTextAsync(response).ConfigureAwait(false);
                            throw new AiException(DescribeHttpError(response.StatusCode,
                                TryDeserialize(serializer, text), text));
                        }

                        if (IsEventStream(response))
                        {
                            await ReadEventStreamAsync(response, serializer, reply, idle, watch, onChunk)
                                .ConfigureAwait(false);
                        }
                        else
                        {
                            // 接口不支持流式、直接回了完整的 JSON，也照样认。
                            AbsorbWholeReply(TryDeserialize(serializer,
                                await ReadTextAsync(response).ConfigureAwait(false)), reply);
                            reply.FirstDataSeconds = watch.Elapsed.TotalSeconds;
                        }
                    }
                }
                catch (Exception ex) when (linked.IsCancellationRequested && !(ex is AiException))
                {
                    if (cancellation.IsCancellationRequested)
                    {
                        throw new OperationCanceledException(cancellation);
                    }

                    if (timeout.IsCancellationRequested)
                    {
                        throw new AiException(
                            $"AI 接口 {config.TimeoutSeconds} 秒内没有返回完，已超时。可以少选一些内容、降低思考强度，" +
                            "或者在「AI 配置」里调大 TimeoutSeconds。");
                    }

                    throw new AiException(responded
                        ? $"AI 接口已经 {IdleTimeoutSeconds} 秒没有返回新内容，像是卡住了，请稍后重试。"
                        : $"{IdleTimeoutSeconds} 秒内没有收到 AI 接口的任何响应，可能是网络或接口暂时卡住了，请稍后重试。");
                }
                catch (HttpRequestException ex)
                {
                    throw new AiException("连不上 AI 接口：" + Innermost(ex).Message, ex);
                }
                catch (Exception ex) when (ex is IOException || ex is WebException)
                {
                    throw new AiException("和 AI 接口的连接中途断开了：" + Innermost(ex).Message + "\n请重试。", ex);
                }
            }

            var content = reply.Content.ToString();
            var usage = reply.Usage;

            AddInLog.Info(string.Format(
                "AI 请求完成：model={0} effort={1} 输入 {2} 字，首包 {3}s，耗时 {4:0.0}s，" +
                "token 输入/输出/思考 = {5}/{6}/{7}，finish={8}",
                modelId, effort, userContent.Length, reply.FirstDataSeconds?.ToString("0.0") ?? "-",
                watch.Elapsed.TotalSeconds,
                Get(usage, "prompt_tokens"), Get(usage, "completion_tokens"),
                Get(Get(usage, "completion_tokens_details"), "reasoning_tokens") ?? 0, reply.FinishReason));

            if (string.Equals(reply.FinishReason, "length", StringComparison.OrdinalIgnoreCase))
            {
                throw new AiException("AI 的输出太长被截断了。请少选一些内容，或者降低思考强度后再试。");
            }

            if (string.IsNullOrWhiteSpace(content))
            {
                throw new AiException("AI 接口没有返回内容。");
            }

            return content;
        }

        /// <summary>
        /// 逐行读流式返回（SSE），直到收到 [DONE] 或连接关闭。收到任何一行（包括保活的注释行）都把「卡住」的计时清零。
        /// </summary>
        private static async Task ReadEventStreamAsync(HttpResponseMessage response, JavaScriptSerializer serializer,
            AiReply reply, CancellationTokenSource idle, Stopwatch watch, Action<string, string> onChunk)
        {
            using (var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
            using (var reader = new StreamReader(stream, Encoding.UTF8))
            {
                string line;
                while (!reply.Done && (line = await reader.ReadLineAsync().ConfigureAwait(false)) != null)
                {
                    idle.CancelAfter(IdleTimeout);

                    var (reasoning, content) = AbsorbStreamLine(line, serializer, reply);
                    if (reply.FirstDataSeconds == null && line.StartsWith("data:", StringComparison.Ordinal))
                    {
                        reply.FirstDataSeconds = watch.Elapsed.TotalSeconds;
                    }

                    if (reasoning.Length + content.Length > 0)
                    {
                        onChunk?.Invoke(reasoning, content);
                    }
                }
            }

            // 既没收到 [DONE] 也没有 finish_reason：连接被中途掐断了，拿到的 JSON 多半不完整。
            if (!reply.IsComplete)
            {
                throw new AiException("和 AI 接口的连接中途断开了，没有收到完整的结果。请重试。");
            }
        }

        /// <summary>
        /// 处理流式返回的一行，返回这一行带来的思考和正文（没有时为空串）。
        /// 注释行（: keep-alive）、空行、解析不了的行都忽略；接口在流里报错时抛 <see cref="AiException"/>。
        /// </summary>
        internal static (string Reasoning, string Content) AbsorbStreamLine(string line, JavaScriptSerializer serializer,
            AiReply reply)
        {
            if (line == null || !line.StartsWith("data:", StringComparison.Ordinal))
            {
                return (string.Empty, string.Empty);
            }

            var data = line.Substring(5).Trim();
            if (data == "[DONE]")
            {
                reply.Done = true;
                return (string.Empty, string.Empty);
            }

            var chunk = TryDeserialize(serializer, data);
            var error = Get(chunk, "error");
            if (error != null)
            {
                throw new AiException("AI 接口返回错误：" + (error as string ?? Get(error, "message") as string ?? data));
            }

            var choice = FirstItem(Get(chunk, "choices"));
            var delta = Get(choice, "delta");
            var reasoning = Get(delta, "reasoning_content") as string ?? string.Empty;
            var content = Get(delta, "content") as string ?? string.Empty;

            reply.Content.Append(content);
            reply.ReasoningChars += reasoning.Length;
            reply.FinishReason = Get(choice, "finish_reason") as string ?? reply.FinishReason;
            reply.Usage = Get(chunk, "usage") ?? reply.Usage;

            return (reasoning, content);
        }

        private static void AbsorbWholeReply(object root, AiReply reply)
        {
            var choice = FirstItem(Get(root, "choices"));
            reply.Content.Append(Get(Get(choice, "message"), "content") as string);
            reply.FinishReason = Get(choice, "finish_reason") as string;
            reply.Usage = Get(root, "usage");
            reply.Done = true;
        }

        private static bool IsEventStream(HttpResponseMessage response)
        {
            return string.Equals(response.Content.Headers.ContentType?.MediaType, "text/event-stream",
                StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>自己按 UTF-8 解：有的网关不带 charset，交给 ReadAsStringAsync 猜可能猜错。</summary>
        private static async Task<string> ReadTextAsync(HttpResponseMessage response)
        {
            var bytes = await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
            return Encoding.UTF8.GetString(bytes);
        }

        /// <summary>请求体：要求输出 JSON 对象、流式返回，思考强度按 <see cref="AiEfforts.ApplyTo"/> 的映射传。</summary>
        internal static Dictionary<string, object> BuildRequestBody(AiConfig config, string modelId, string effort,
            string systemPrompt, string userContent)
        {
            var body = new Dictionary<string, object>
            {
                ["model"] = modelId,
                ["messages"] = new[]
                {
                    new Dictionary<string, object> { ["role"] = "system", ["content"] = systemPrompt },
                    new Dictionary<string, object> { ["role"] = "user", ["content"] = userContent }
                },
                ["response_format"] = new Dictionary<string, object> { ["type"] = "json_object" },
                ["stream"] = true,

                // 流式默认不给用量，要了才会在最后一段带上，日志里要记。
                ["stream_options"] = new Dictionary<string, object> { ["include_usage"] = true }
            };

            AiEfforts.ApplyTo(body, effort);

            if (config.MaxTokens > 0)
            {
                body["max_tokens"] = config.MaxTokens;
            }

            return body;
        }

        /// <summary>在配置的基础地址后补上 /chat/completions；已写全的地址直接使用。</summary>
        internal static string BuildEndpoint(string apiUrl)
        {
            var url = (apiUrl ?? string.Empty).Trim().TrimEnd('/');
            return url.EndsWith("/chat/completions", StringComparison.OrdinalIgnoreCase)
                ? url
                : url + "/chat/completions";
        }

        private static string DescribeHttpError(HttpStatusCode status, object root, string responseText)
        {
            var detail = Get(Get(root, "error"), "message") as string;

            // 模型名是用户在配置里手填的，填了网关上没有的（或者网关还没开通），网关回 503 + model_not_found。
            if (string.Equals(Get(Get(root, "error"), "code") as string, "model_not_found", StringComparison.OrdinalIgnoreCase))
            {
                return $"AI 接口上没有这个模型（HTTP {(int)status}）。请在「模型」里换一个，" +
                       $"或者检查「AI 配置」里的模型名是否写对、网关是否已经开通。\n{detail}";
            }
            if (string.IsNullOrWhiteSpace(detail))
            {
                detail = responseText?.Trim() ?? string.Empty;
                if (detail.Length > 200)
                {
                    detail = detail.Substring(0, 200) + "…";
                }
            }

            switch ((int)status)
            {
                case 401:
                case 403:
                    return $"AI 接口拒绝访问（HTTP {(int)status}），请检查「AI 配置」里的 ApiKey。\n{detail}";
                case 404:
                    return $"AI 接口地址不对（HTTP 404），请检查「AI 配置」里的 ApiUrl。\n{detail}";
                case 429:
                    return $"AI 接口限流了（HTTP 429），请稍后再试。\n{detail}";
                default:
                    return $"AI 接口返回错误（HTTP {(int)status}）：{detail}";
            }
        }

        private static Exception Innermost(Exception ex)
        {
            while (ex.InnerException != null)
            {
                ex = ex.InnerException;
            }

            return ex;
        }

        private static object TryDeserialize(JavaScriptSerializer serializer, string json)
        {
            try
            {
                return string.IsNullOrWhiteSpace(json) ? null : serializer.DeserializeObject(json);
            }
            catch (Exception)
            {
                // 网关出错时可能回一段 HTML，交给调用方按原文报错。
                return null;
            }
        }

        /// <summary>从 JavaScriptSerializer 解出来的对象里按键取值，取不到返回 null。</summary>
        internal static object Get(object node, string key)
        {
            return node is IDictionary<string, object> map && map.TryGetValue(key, out var value) ? value : null;
        }

        internal static object FirstItem(object node)
        {
            return node is IList list && list.Count > 0 ? list[0] : null;
        }
    }
}
