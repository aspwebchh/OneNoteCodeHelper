using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
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

    /// <summary>
    /// OpenAI 兼容的 Chat Completions 接口的最小封装：一问一答，要求模型输出 JSON 对象。
    /// </summary>
    internal static class AiClient
    {
        private static readonly HttpClient Http;

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
        /// 失败一律抛 <see cref="AiException"/>；用户取消抛 <see cref="OperationCanceledException"/>。
        /// </summary>
        internal static async Task<string> CompleteAsync(AiConfig config, string modelId, string effort,
            string systemPrompt, string userContent, CancellationToken cancellation)
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
                ["stream"] = false
            };

            if (!string.IsNullOrEmpty(effort))
            {
                body["reasoning_effort"] = effort;
            }

            if (config.MaxTokens > 0)
            {
                body["max_tokens"] = config.MaxTokens;
            }

            var serializer = CreateSerializer();
            var watch = Stopwatch.StartNew();
            string responseText;
            HttpStatusCode status;

            using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellation))
            using (var request = new HttpRequestMessage(HttpMethod.Post, BuildEndpoint(config.ApiUrl)))
            {
                timeout.CancelAfter(TimeSpan.FromSeconds(config.TimeoutSeconds));
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", config.ApiKey);
                request.Content = new StringContent(serializer.Serialize(body), Encoding.UTF8, "application/json");

                try
                {
                    // 默认的 ResponseContentRead 会在 SendAsync 里把正文读完，超时也覆盖到读正文。
                    using (var response = await Http.SendAsync(request, timeout.Token).ConfigureAwait(false))
                    {
                        status = response.StatusCode;

                        // 自己按 UTF-8 解：有的网关不带 charset，交给 ReadAsStringAsync 猜可能猜错。
                        var bytes = await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                        responseText = Encoding.UTF8.GetString(bytes);
                    }
                }
                catch (OperationCanceledException) when (!cancellation.IsCancellationRequested)
                {
                    throw new AiException(
                        $"AI 接口 {config.TimeoutSeconds} 秒内没有返回，已超时。可以少选一些内容、降低思考强度，" +
                        "或者在「AI 配置」里调大 TimeoutSeconds。");
                }
                catch (HttpRequestException ex)
                {
                    throw new AiException("连不上 AI 接口：" + Innermost(ex).Message, ex);
                }
            }

            var root = TryDeserialize(serializer, responseText);

            if ((int)status < 200 || (int)status >= 300)
            {
                throw new AiException(DescribeHttpError(status, root, responseText));
            }

            var choice = FirstItem(Get(root, "choices"));
            var content = Get(Get(choice, "message"), "content") as string;
            var finishReason = Get(choice, "finish_reason") as string;
            var usage = Get(root, "usage");

            AddInLog.Info(string.Format(
                "AI 请求完成：model={0} effort={1} 输入 {2} 字，耗时 {3:0.0}s，token 输入/输出/思考 = {4}/{5}/{6}，finish={7}",
                modelId, effort, userContent.Length, watch.Elapsed.TotalSeconds,
                Get(usage, "prompt_tokens"), Get(usage, "completion_tokens"),
                Get(Get(usage, "completion_tokens_details"), "reasoning_tokens") ?? 0, finishReason));

            if (string.Equals(finishReason, "length", StringComparison.OrdinalIgnoreCase))
            {
                throw new AiException("AI 的输出太长被截断了。请少选一些内容，或者降低思考强度后再试。");
            }

            if (string.IsNullOrWhiteSpace(content))
            {
                throw new AiException("AI 接口没有返回内容。");
            }

            return content;
        }

        /// <summary>配置里写到 /v1 为止；已经写全了 /chat/completions 的也认。</summary>
        private static string BuildEndpoint(string apiUrl)
        {
            var url = (apiUrl ?? string.Empty).Trim().TrimEnd('/');
            return url.EndsWith("/chat/completions", StringComparison.OrdinalIgnoreCase)
                ? url
                : url + "/chat/completions";
        }

        private static string DescribeHttpError(HttpStatusCode status, object root, string responseText)
        {
            var detail = Get(Get(root, "error"), "message") as string;
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
