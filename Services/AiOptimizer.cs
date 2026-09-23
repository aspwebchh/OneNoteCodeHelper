using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OneNoteCodeHelper.Services
{
    /// <summary>
    /// 进度：一句话说明 + 已完成 / 总批数，Total 为 0 表示眼下没法给出比例。
    /// Detail 是等 AI 时的实时情况（已思考、已输出多少字），不在等 AI 时为 null。
    /// </summary>
    internal sealed class AiProgress
    {
        internal AiProgress(string message, int done = 0, int total = 0, string detail = null)
        {
            Message = message;
            Done = done;
            Total = total;
            Detail = detail;
        }

        internal string Message { get; }

        internal int Done { get; }

        internal int Total { get; }

        internal string Detail { get; }
    }

    /// <summary>
    /// 一次「AI 优化」的结果，进度窗按它画结果区。Error 不为 null 表示失败，这时其余字段没有意义。
    /// </summary>
    internal sealed class AiReport
    {
        private AiReport(string error)
        {
            Error = error;
            Changes = Array.Empty<string>();
        }

        internal AiReport(string scope, IReadOnlyList<string> changes, int applied, int removedBlankLines,
            int conflicted)
        {
            Scope = scope;
            Changes = changes;
            Applied = applied;
            RemovedBlankLines = removedBlankLines;
            Conflicted = conflicted;
        }

        internal string Error { get; }

        internal bool Success => Error == null;

        /// <summary>「整页」或「选中的」。</summary>
        internal string Scope { get; }

        /// <summary>AI 自己给的改动说明，只算真正写回了的段落，已去重。</summary>
        internal IReadOnlyList<string> Changes { get; }

        /// <summary>写回了改动的段落数。不显示，只用来判断页面有没有动、AI 有没有漏写说明。</summary>
        internal int Applied { get; }

        /// <summary>删掉的空行数。同上，不显示。</summary>
        internal int RemovedBlankLines { get; }

        /// <summary>处理期间被用户改过、为了不覆盖而跳过的段落数。不单独提示，只用来选结果的状态。</summary>
        internal int Conflicted { get; }

        /// <summary>页面有没有被改动。</summary>
        internal bool Changed => Applied + RemovedBlankLines > 0;

        internal static AiReport Failed(string message) => new AiReport(message ?? "AI 优化失败。");
    }

    /// <summary>AI 对一个段落的回复：改后的全文，和它自己写的改动说明（可能为空）。</summary>
    internal sealed class AiParagraphReply
    {
        internal AiParagraphReply(string text, IReadOnlyList<string> changes)
        {
            Text = text;
            Changes = changes;
        }

        internal string Text { get; }

        internal IReadOnlyList<string> Changes { get; }
    }

    /// <summary>
    /// 一次「AI 优化」：读段落 → 分批问 AI → 写回。整个过程在线程池（MTA）上跑，
    /// 和「高亮选中」一样直接调 OneNote 的 COM 对象，不用封送。
    /// </summary>
    internal sealed class AiOptimizer
    {
        /// <summary>每批大约多少字。太大单次输出容易被截断，太小请求数多、上下文也少。</summary>
        private const int BatchChars = 3000;

        /// <summary>同时在飞的请求数。</summary>
        private const int MaxConcurrency = 3;

        /// <summary>流式返回时，进度最多每隔这么久（毫秒）刷新一次。</summary>
        private const int ReportIntervalMs = 200;

        /// <summary>接在用户提示词后面的固定约定。放在代码里，用户改提示词也不会破坏输入输出格式。</summary>
        private const string Protocol =
            "【输入格式】用户消息是一个 JSON 对象：{\"paragraphs\":[{\"id\":段落编号,\"text\":\"段落原文\"}]}，" +
            "每一项是笔记里的一个段落。\n" +
            "【输出要求】\n" +
            "1. 只输出一个 JSON 对象：" +
            "{\"paragraphs\":[{\"id\":段落编号,\"text\":\"修改后的完整段落\",\"changes\":[\"改动说明\"]}]}，" +
            "不要在 JSON 之外输出任何解释，也不要用 Markdown 代码块包裹。\n" +
            "2. 只列出确实做了修改的段落；没有任何修改时输出 {\"paragraphs\":[]}。\n" +
            "3. id 必须与输入对应。每个段落单独处理：不能合并、拆分或调换段落，text 里不要加入换行。\n" +
            "4. 段落中的代码、命令、网址、文件路径、邮箱地址保持原样。\n" +
            "5. changes 用简短的中文逐条说明这一段改了什么，每条只说一件事，" +
            "例如 \"帐号 → 账号\"、\"中英文之间加空格\"；同一类改动合成一条。";

        private readonly PageEditor _editor;

        private readonly AiConfig _config;

        private readonly AiFunction _function;

        private readonly AiModel _model;

        private readonly string _effort;

        internal AiOptimizer(PageEditor editor, AiConfig config, AiFunction function, AiModel model, string effort)
        {
            _editor = editor;
            _config = config;
            _function = function;
            _model = model;
            _effort = effort;
        }

        internal async Task<AiReport> RunAsync(IProgress<AiProgress> progress, CancellationToken cancellation)
        {
            progress.Report(new AiProgress("正在读取页面…"));

            var read = _editor.ReadAiTargets(out var targets);
            if (!read.Success)
            {
                return AiReport.Failed(read.Message);
            }

            var paragraphs = targets.Paragraphs;
            var batches = SplitIntoBatches(paragraphs.Count, i => paragraphs[i].Text.Length);
            var scope = targets.WholePage ? "整页" : "选中的";
            AddInLog.Info($"AI 优化开始：{_function.Name}，{scope} {paragraphs.Count} 段，分 {batches.Count} 批。");

            var replies = await AskInBatchesAsync(paragraphs, batches, scope, progress, cancellation)
                .ConfigureAwait(false);

            var edits = new List<AiParagraphEdit>();

            for (var i = 0; i < paragraphs.Count; i++)
            {
                var source = paragraphs[i];
                var text = source.Text;
                IReadOnlyList<string> changes = Array.Empty<string>();

                // AI 回了什么就用什么，不再按改动大小把关。
                if (replies.TryGetValue(i, out var reply))
                {
                    text = CleanReplyText(source.Text, reply.Text);
                    changes = reply.Changes;
                }

                // 段内多余的换行由插件自己删，AI 没改的段落也照样删。
                if (_function.RemoveExtraBlankLines)
                {
                    text = BlankLines.CollapseInText(text);
                }

                if (text != source.Text)
                {
                    edits.Add(new AiParagraphEdit(source, text, changes));
                }
            }

            // 开始写回之后就不再响应取消：写到一半停下来，页面会处在谁也说不清的状态。
            cancellation.ThrowIfCancellationRequested();

            var applied = new List<AiParagraphEdit>();
            var conflicted = 0;
            var removedBlankLines = 0;

            // 要删空行时即使没有段落要改也得写回：空行段落不经过 AI，只有写回时才知道删不删。
            if (edits.Count > 0 || _function.RemoveExtraBlankLines)
            {
                progress.Report(new AiProgress("正在写回 OneNote…"));

                var write = _editor.ApplyParagraphEdits(targets, edits, _function.RemoveExtraBlankLines,
                    out applied, out conflicted, out removedBlankLines);
                if (!write.Success)
                {
                    return AiReport.Failed(write.Message);
                }
            }

            AddInLog.Info($"AI 优化结束：改了 {applied.Count} 段，删了 {removedBlankLines} 个空行，跳过 {conflicted} 段。");

            // 只汇总真正写回了的段落：处理期间被用户改过而跳过的，它们的说明不算数。
            return new AiReport(scope, MergeChanges(applied.Select(e => e.Changes)), applied.Count,
                removedBlankLines, conflicted);
        }

        /// <summary>把各段的改动说明按先后顺序连起来，去掉重复的和空的。</summary>
        internal static List<string> MergeChanges(IEnumerable<IReadOnlyList<string>> changes)
        {
            var seen = new HashSet<string>(StringComparer.Ordinal);
            var result = new List<string>();

            foreach (var change in changes.SelectMany(c => c))
            {
                var text = change?.Trim();
                if (!string.IsNullOrEmpty(text) && seen.Add(text))
                {
                    result.Add(text);
                }
            }

            return result;
        }

        /// <summary>
        /// 分批并发地问 AI，返回「段落下标 → AI 的回复」。任何一批失败就取消其余的，抛出那一批的异常。
        /// </summary>
        private async Task<Dictionary<int, AiParagraphReply>> AskInBatchesAsync(IReadOnlyList<AiParagraph> paragraphs,
            List<List<int>> batches, string scope, IProgress<AiProgress> progress, CancellationToken cancellation)
        {
            var systemPrompt = BuildSystemPrompt(_function.Prompt);
            var results = new Dictionary<int, AiParagraphReply>();
            var sync = new object();
            var done = 0;
            var reasoningChars = 0;
            var contentChars = 0;
            var sinceReport = Stopwatch.StartNew();

            // 流式返回的每一段都会调到这里（几批同时在跑时来自不同线程）。段很碎，
            // 除了一批做完，其余的限一下频率，免得把界面线程的消息队列塞满。
            void Update(int reasoning, int content, bool batchDone)
            {
                AiProgress snapshot;
                lock (sync)
                {
                    reasoningChars += reasoning;
                    contentChars += content;
                    if (batchDone)
                    {
                        done++;
                    }
                    else if (sinceReport.ElapsedMilliseconds < ReportIntervalMs)
                    {
                        return;
                    }

                    sinceReport.Restart();
                    snapshot = new AiProgress(
                        done == 0
                            ? $"正在请 AI {_function.Name}：{scope} {paragraphs.Count} 段…"
                            : $"正在请 AI {_function.Name}：已完成 {done} / {batches.Count} 批…",
                        done, batches.Count, DescribeLive(reasoningChars, contentChars));
                }

                progress.Report(snapshot);
            }

            progress.Report(new AiProgress($"正在请 AI {_function.Name}：{scope} {paragraphs.Count} 段…",
                0, batches.Count, DescribeLive(0, 0)));

            using (var gate = new SemaphoreSlim(MaxConcurrency))
            using (var failFast = CancellationTokenSource.CreateLinkedTokenSource(cancellation))
            {
                var tasks = batches.Select(async batch =>
                {
                    await gate.WaitAsync(failFast.Token).ConfigureAwait(false);
                    try
                    {
                        // 编号从 1 开始、每批各自编：模型看到的 id 越短越不容易抄错。
                        var content = await AiClient.CompleteAsync(
                            _config, _model.Id, _effort, systemPrompt,
                            BuildUserMessage(batch.Select((index, n) => (n + 1, paragraphs[index].Text))),
                            (reasoning, text) => Update(reasoning, text, false),
                            failFast.Token).ConfigureAwait(false);

                        var reply = ParseReply(content);
                        lock (results)
                        {
                            foreach (var item in reply.Where(r => r.Key >= 1 && r.Key <= batch.Count))
                            {
                                results[batch[item.Key - 1]] = item.Value;
                            }
                        }

                        Update(0, 0, true);
                    }
                    catch
                    {
                        failFast.Cancel();
                        throw;
                    }
                    finally
                    {
                        gate.Release();
                    }
                }).ToList();

                try
                {
                    await Task.WhenAll(tasks).ConfigureAwait(false);
                }
                catch
                {
                    // WhenAll 抛的是任务列表里第一个失败的，可能只是被 failFast 连带取消的那批。
                    // 找出真正出错的那个抛出去，用户才能看到原因。
                    var cause = tasks.Where(t => t.IsFaulted)
                        .SelectMany(t => t.Exception.InnerExceptions)
                        .FirstOrDefault(e => !(e is OperationCanceledException));

                    if (cause != null)
                    {
                        ExceptionDispatchInfo.Capture(cause).Throw();
                    }

                    throw;
                }
            }

            return results;
        }

        /// <summary>按字数把段落切成批，返回每批的段落下标。超长的单段自己一批。</summary>
        internal static List<List<int>> SplitIntoBatches(int count, Func<int, int> lengthOf)
        {
            var batches = new List<List<int>>();
            var current = new List<int>();
            var chars = 0;

            for (var i = 0; i < count; i++)
            {
                var length = lengthOf(i);
                if (current.Count > 0 && chars + length > BatchChars)
                {
                    batches.Add(current);
                    current = new List<int>();
                    chars = 0;
                }

                current.Add(i);
                chars += length;
            }

            if (current.Count > 0)
            {
                batches.Add(current);
            }

            return batches;
        }

        /// <summary>等 AI 时显示的实时情况：思考、输出各收到了多少字。</summary>
        internal static string DescribeLive(int reasoningChars, int contentChars)
        {
            if (reasoningChars == 0 && contentChars == 0)
            {
                return "等待 AI 响应";
            }

            var parts = new List<string>();
            if (reasoningChars > 0)
            {
                parts.Add($"已思考 {reasoningChars} 字");
            }

            if (contentChars > 0)
            {
                parts.Add($"已输出 {contentChars} 字");
            }

            return "AI 正在处理：" + string.Join("，", parts);
        }

        internal static string BuildSystemPrompt(string taskPrompt)
        {
            return taskPrompt.Trim() + "\n\n" + Protocol;
        }

        internal static string BuildUserMessage(IEnumerable<(int Id, string Text)> paragraphs)
        {
            var body = new Dictionary<string, object>
            {
                ["paragraphs"] = paragraphs
                    .Select(p => new Dictionary<string, object> { ["id"] = p.Id, ["text"] = p.Text })
                    .ToArray()
            };

            return AiClient.CreateSerializer().Serialize(body);
        }

        /// <summary>
        /// 解析模型的输出，返回「id → 回复」。容忍外面包了 ``` 代码块、直接给了数组、id 写成字符串，
        /// changes 缺了、或者写成了单个字符串。格式完全不对时抛 <see cref="AiException"/>。
        /// </summary>
        internal static Dictionary<int, AiParagraphReply> ParseReply(string content)
        {
            var json = StripCodeFence(content);

            object root;
            try
            {
                root = AiClient.CreateSerializer().DeserializeObject(json);
            }
            catch (Exception ex)
            {
                throw new AiException("AI 返回的不是约定的 JSON 格式，没有写回。可以换个模型或者再试一次。", ex);
            }

            var items = AiClient.Get(root, "paragraphs") ?? root;
            if (!(items is IList list))
            {
                throw new AiException("AI 返回的 JSON 里没有 paragraphs 数组，没有写回。可以换个模型或者再试一次。");
            }

            var result = new Dictionary<int, AiParagraphReply>();
            foreach (var item in list)
            {
                var id = AiClient.Get(item, "id");
                var text = AiClient.Get(item, "text") as string;

                if (text != null && int.TryParse(Convert.ToString(id, CultureInfo.InvariantCulture),
                        NumberStyles.Integer, CultureInfo.InvariantCulture, out var number))
                {
                    result[number] = new AiParagraphReply(text, ReadChanges(AiClient.Get(item, "changes")));
                }
            }

            return result;
        }

        /// <summary>说明只是给人看的，写得不规范也不影响写回：不是字符串的项跳过，空的丢掉。</summary>
        private static IReadOnlyList<string> ReadChanges(object value)
        {
            var items = value is string single ? new object[] { single } : value as IList;
            if (items == null)
            {
                return Array.Empty<string>();
            }

            return items.OfType<string>()
                .Select(s => s.Trim())
                .Where(s => s.Length > 0)
                .ToList();
        }

        /// <summary>
        /// 约定里说了不许加换行，模型偶尔还是会加。原文没有换行（&lt;br&gt;）时把它去掉：
        /// 两边都是英文字母或数字才补一个空格，中文之间直接接上。
        /// </summary>
        internal static string CleanReplyText(string original, string text)
        {
            text = text ?? string.Empty;
            if (original.IndexOf('\n') >= 0 || (text.IndexOf('\n') < 0 && text.IndexOf('\r') < 0))
            {
                return text;
            }

            var builder = new StringBuilder(text.Length);
            var i = 0;
            while (i < text.Length)
            {
                if (text[i] != '\r' && text[i] != '\n')
                {
                    builder.Append(text[i++]);
                    continue;
                }

                while (i < text.Length && (text[i] == '\r' || text[i] == '\n'))
                {
                    i++;
                }

                if (builder.Length > 0 && i < text.Length
                    && IsAsciiWord(builder[builder.Length - 1]) && IsAsciiWord(text[i]))
                {
                    builder.Append(' ');
                }
            }

            return builder.ToString();
        }

        private static bool IsAsciiWord(char ch)
        {
            return ch < 128 && char.IsLetterOrDigit(ch);
        }

        private static string StripCodeFence(string content)
        {
            var text = (content ?? string.Empty).Trim();
            if (!text.StartsWith("```", StringComparison.Ordinal))
            {
                return text;
            }

            var firstLineEnd = text.IndexOf('\n');
            var lastFence = text.LastIndexOf("```", StringComparison.Ordinal);
            return firstLineEnd > 0 && lastFence > firstLineEnd
                ? text.Substring(firstLineEnd + 1, lastFence - firstLineEnd - 1).Trim()
                : text;
        }
    }
}
