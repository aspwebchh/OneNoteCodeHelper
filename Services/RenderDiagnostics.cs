using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using OneNoteCodeHelper.Highlighting;
using OneNoteCodeHelper.Highlighting.Themes;

namespace OneNoteCodeHelper.Services
{
    /// <summary>
    /// 不接触 OneNote 就能跑通整条渲染链路的入口，供排查问题和验证新语言用。
    /// Tools/selftest.ps1 通过反射调用这里。
    /// </summary>
    internal static class RenderDiagnostics
    {
        /// <summary>渲染成逐行 HTML，行间用 \n 连接。</summary>
        internal static string RenderLines(string code, string languageId, string themeId)
        {
            var language = ResolveLanguage(code, languageId);
            var theme = CodeThemes.Find(themeId);
            var settings = new AddInSettings { ThemeId = theme.Id };

            var lines = OneNoteHtmlEncoder.BuildLines(
                code, language.Tokenize(code), theme, settings.TabWidth);

            return string.Join("\n", lines);
        }

        /// <summary>渲染成完整的 one:Table XML。</summary>
        internal static string RenderTableXml(string code, string languageId, string themeId)
        {
            var language = ResolveLanguage(code, languageId);
            var theme = CodeThemes.Find(themeId);
            var settings = new AddInSettings { ThemeId = theme.Id };

            return CodeBlockBuilder.BuildTable(code, language, theme, settings)
                .ToString(SaveOptions.None);
        }

        /// <summary>
        /// 生成 PageEditor 实际会发给 UpdatePageContent 的那份 XML，用于核对命名空间前缀、
        /// OE 数量、字体样式这些只有在完整上下文里才看得出的东西。
        /// </summary>
        internal static string RenderPageChangesXml(string code, string languageId, string themeId)
        {
            var language = ResolveLanguage(code, languageId);
            var theme = CodeThemes.Find(themeId);
            var settings = new AddInSettings { ThemeId = theme.Id };

            var table = CodeBlockBuilder.BuildTable(code, language, theme, settings);
            var outline = CodeBlockBuilder.BuildOutline(table, 36, 86, settings.CodeBlockWidth);

            return PageEditor.BuildPageChanges("{TEST-PAGE-ID}", outline);
        }

        /// <summary>把 token 序列转成 "kind:文本" 的可读形式，便于核对词法分析结果。</summary>
        internal static string DumpTokens(string code, string languageId)
        {
            var language = ResolveLanguage(code, languageId);

            return string.Join("\n", language.Tokenize(code)
                .Where(t => t.Kind != TokenKind.Plain)
                .Select(t => t.Kind + ": " + code.Substring(t.Start, t.Length).Replace("\n", "\\n")));
        }

        /// <summary>检查 token 是否完整、无重叠地覆盖了整个源码——这是 ILanguage 的核心契约。</summary>
        internal static string VerifyCoverage(string code, string languageId)
        {
            var language = ResolveLanguage(code, languageId);
            var position = 0;

            foreach (var token in language.Tokenize(code))
            {
                if (token.Start != position)
                {
                    return $"FAIL: 位置 {position} 处出现断裂或重叠，下一个 token 从 {token.Start} 开始";
                }

                if (token.Length <= 0)
                {
                    return $"FAIL: 位置 {token.Start} 处出现空 token";
                }

                position = token.End;
            }

            return position == code.Length
                ? "OK: 完整覆盖 " + position + " 个字符"
                : $"FAIL: 只覆盖到 {position}，源码共 {code.Length} 个字符";
        }

        /// <summary>返回自动识别的结果名，识别不出来返回 "(无法确定)"。</summary>
        internal static string DetectLanguageName(string code)
        {
            return LanguageRegistry.Detect(code)?.DisplayName ?? "(无法确定)";
        }

        /// <summary>自动识别时各语言的得分，形如 "java=12 csharp=5"，不列 0 分的。Tools/detect-test.ps1 调这里。</summary>
        internal static string DescribeDetectionScores(string code)
        {
            if (string.IsNullOrWhiteSpace(code))
            {
                return string.Empty;
            }

            return string.Join(" ", LanguageRegistry.Rank(code)
                .Where(x => x.Score != 0)
                .Select(x => x.Language.Id + "=" + x.Score));
        }

        /// <summary>
        /// 解析一个 one:OE 的 XML，返回给 AI 看的纯文本。拼不回原样的段落返回 "UNSUPPORTED"。
        /// Tools/ai-merge-test.ps1 调这里。
        /// </summary>
        internal static string ParagraphText(string oeXml)
        {
            var rich = RichParagraph.Parse(XElement.Parse(oeXml));
            return rich.IsLossless ? rich.Text : "UNSUPPORTED";
        }

        /// <summary>把新文本合并进一个 one:OE，返回合并后各个 one:T 的 HTML，用 | 隔开。</summary>
        internal static string MergeParagraph(string oeXml, string newText)
        {
            var oe = XElement.Parse(oeXml);
            RichParagraph.Parse(oe).Apply(newText);
            return string.Join("|", oe.Elements(OneNoteApi.One + "T").Select(t => t.Value));
        }

        /// <summary>解析模型输出，返回 "id=文本" 逐行，按 id 排序；有改动说明时后面跟 " [说明1; 说明2]"。</summary>
        internal static string ParseAiReply(string content)
        {
            return string.Join("\n", AiOptimizer.ParseReply(content)
                .OrderBy(x => x.Key)
                .Select(x => x.Key + "=" + x.Value.Text +
                             (x.Value.Changes.Count > 0 ? " [" + string.Join("; ", x.Value.Changes) + "]" : "")));
        }

        /// <summary>
        /// 汇总各段的改动说明。paragraphs 里段与段用 | 隔开，一段里的说明用 ; 隔开；返回汇总后的清单，用 | 隔开。
        /// </summary>
        internal static string MergeAiChanges(string paragraphs)
        {
            return string.Join("|", AiOptimizer.MergeChanges(paragraphs.Split('|')
                .Select(p => (IReadOnlyList<string>)p.Split(';'))));
        }

        internal static string CleanAiReplyText(string original, string text)
        {
            return AiOptimizer.CleanReplyText(original, text);
        }

        /// <summary>按默认配置生成的请求体 JSON（不含提示词正文），用来核对思考强度映射成了哪些参数。</summary>
        internal static string DescribeRequestBody(string effort)
        {
            var config = AiConfigStore.Default;
            var body = AiClient.BuildRequestBody(config, config.Models[0].Id, effort, "system", "user");
            body.Remove("messages");
            return AiClient.CreateSerializer().Serialize(body);
        }

        /// <summary>
        /// 把一段流式返回（SSE）逐行交给解析，返回 "正文|finish_reason|思考字数|complete 或 incomplete"；
        /// 流里报错时返回 "ERROR: 消息"。
        /// </summary>
        internal static string ParseAiStream(string sse)
        {
            var reply = new AiReply();
            var serializer = AiClient.CreateSerializer();

            try
            {
                foreach (var line in sse.Split('\n'))
                {
                    AiClient.AbsorbStreamLine(line.TrimEnd('\r'), serializer, reply);
                }
            }
            catch (AiException ex)
            {
                return "ERROR: " + ex.Message;
            }

            return $"{reply.Content}|{reply.FinishReason}|{reply.ReasoningChars}|" +
                   (reply.IsComplete ? "complete" : "incomplete");
        }

        internal static string DescribeAiLive(bool started, bool outputting, int returned, string lastChange)
        {
            return AiOptimizer.DescribeLive(started, outputting, returned, lastChange);
        }

        /// <summary>
        /// 把 AI 返回的正文按片段（用 \u001F 隔开）逐段交给 <see cref="ReplyPeek"/>，
        /// 返回 "已返回段数|最新改动说明"。
        /// </summary>
        internal static string PeekAiReply(string fragments)
        {
            var peek = new ReplyPeek();
            foreach (var fragment in fragments.Split('\u001F'))
            {
                peek.Feed(fragment);
            }

            return $"{peek.Paragraphs}|{peek.LastChange}";
        }

        /// <summary>思考摘录，null 返回空串。</summary>
        internal static string LiveTextExcerpt(string text, int maxChars)
        {
            return LiveText.Excerpt(new System.Text.StringBuilder(text), maxChars) ?? string.Empty;
        }

        internal static string CollapseBlankLinesInText(string text)
        {
            return BlankLines.CollapseInText(text);
        }

        /// <summary>
        /// 在一整页 XML 上删多余的空行。removableIds 是允许删的空行 objectID，逗号分隔；null 表示都能删。
        /// 返回「删了几行: 各摞段落」，摞之间用 " / " 隔开、行之间用 | 隔开：空行写成 _，
        /// 看得见但没有字的（待办、项目符号、表格、代码的空行）写成 #，缩进一级前面加一个 &gt;，删空了的 OEChildren 写成 !。
        /// </summary>
        internal static string RemoveBlankLines(string pageXml, string removableIds)
        {
            var page = XElement.Parse(pageXml);
            var ids = removableIds?.Split(',');
            var removed = BlankLines.RemoveFromPage(page,
                oe => ids == null || ids.Contains((string)oe.Attribute("objectID")), new HashSet<XElement>());

            return removed + ": " + string.Join(" / ", BlankLines.FlowsOf(page)
                .Select(flow => string.Join("|", DescribeLines(flow, 0))));
        }

        /// <summary>
        /// 拿一份带选区标记的页面 XML 跑一遍「高亮选中」的读选区和替换，代码框用一个空的 one:Table 代替。
        /// 返回「源码 =&gt; 替换后的文本块」：源码的换行写成 |、制表符写成 \t；文本块的写法同 <see cref="RemoveBlankLines"/>，
        /// 代码框是 #。读选区失败时返回 "FAIL: 说明"。Tools/highlight-selection-test.ps1 调这里。
        /// </summary>
        internal static string HighlightSelection(string pageXml)
        {
            var page = XElement.Parse(pageXml);
            var result = PageEditor.ReadCodeSelection(page, out var selection);
            if (!result.Success)
            {
                return "FAIL: " + result.Message;
            }

            selection.ReplaceWith(new XElement(OneNoteApi.One + "Table"));

            return selection.Code.Replace("\t", "\\t").Replace("\n", "|") + " => " +
                   string.Join("|", DescribeLines(selection.Block.Element(OneNoteApi.One + "OEChildren"), 0));
        }

        private static IEnumerable<string> DescribeLines(XElement oeChildren, int depth)
        {
            var prefix = new string('>', depth);
            var paragraphs = oeChildren.Elements(OneNoteApi.One + "OE").ToList();
            if (paragraphs.Count == 0)
            {
                yield return prefix + "!";
                yield break;
            }

            foreach (var oe in paragraphs)
            {
                var text = RichParagraph.Parse(oe).Text;
                yield return prefix + (BlankLines.IsBlankLine(oe) ? "_" : string.IsNullOrWhiteSpace(text) ? "#" : text);

                foreach (var line in oe.Elements(OneNoteApi.One + "OEChildren").SelectMany(c => DescribeLines(c, depth + 1)))
                {
                    yield return line;
                }
            }
        }

        /// <summary>按一份 ai-settings.xml 读出各功能，返回「名字=是否删空行」，用 | 隔开。</summary>
        internal static string DescribeAiFunctions(string configXml)
        {
            return string.Join("|", AiConfigStore.Parse(XElement.Parse(configXml)).Functions
                .Select(f => f.Name + "=" + f.RemoveExtraBlankLines));
        }

        /// <summary>第一次点「AI 配置」时生成的那份默认配置。</summary>
        internal static string DefaultAiConfigXml()
        {
            return AiConfigStore.BuildDefaultDocument().ToString();
        }

        /// <summary>
        /// 用本机 ai-settings.xml 真调一次接口：按「AI 优化」完全一样的提示词和格式问，返回解析后的结果。
        /// paragraphs 用 \n 分隔。
        /// </summary>
        internal static string RunAiSample(string functionName, string modelId, string effort, string paragraphs)
        {
            var config = AiConfigStore.Load();
            var function = config.FindFunction(functionName);
            var texts = paragraphs.Split('\n');

            // 模型名原样用，不在配置里找：测试时想试哪个模型就试哪个，找不到也不会悄悄换成别的。
            var content = AiClient.CompleteAsync(
                    config, modelId, AiEfforts.Normalize(effort),
                    AiOptimizer.BuildSystemPrompt(function.Prompt),
                    AiOptimizer.BuildUserMessage(texts.Select((text, i) => (i + 1, text))),
                    null, System.Threading.CancellationToken.None)
                .GetAwaiter().GetResult();

            return $"[{function.Name} · {modelId} · {AiEfforts.Normalize(effort)}]\n" + ParseAiReply(content);
        }

        private static ILanguage ResolveLanguage(string code, string languageId)
        {
            var language = LanguageRegistry.Resolve(languageId, code);
            if (language == null)
            {
                throw new InvalidOperationException("无法确定语言：" + languageId);
            }

            return language;
        }
    }
}
