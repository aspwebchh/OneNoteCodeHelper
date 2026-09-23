using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using OneNoteCodeHelper.Highlighting;
using OneNoteCodeHelper.Highlighting.Themes;

namespace OneNoteCodeHelper.Services
{
    /// <summary>一次页面操作的结果。失败时带一句可以直接给用户看的中文说明。</summary>
    internal sealed class EditResult
    {
        private EditResult(bool success, string message)
        {
            Success = success;
            Message = message;
        }

        internal bool Success { get; }

        internal string Message { get; }

        internal static EditResult Ok(string message = null) => new EditResult(true, message);

        internal static EditResult Fail(string message) => new EditResult(false, message);
    }

    /// <summary>要交给 AI 处理的一个段落。</summary>
    internal sealed class AiParagraph
    {
        internal AiParagraph(string objectId, string text)
        {
            ObjectId = objectId;
            Text = text;
        }

        /// <summary>段落（one:OE）的 objectID，写回时靠它找回这一段。</summary>
        internal string ObjectId { get; }

        /// <summary>发给 AI 的纯文本，也是写回前核对「这段期间有没有被改过」的依据。</summary>
        internal string Text { get; }
    }

    /// <summary>AI 改过的一个段落。</summary>
    internal sealed class AiParagraphEdit
    {
        internal AiParagraphEdit(AiParagraph source, string newText, IReadOnlyList<string> changes)
        {
            Source = source;
            NewText = newText;
            Changes = changes;
        }

        internal AiParagraph Source { get; }

        internal string NewText { get; }

        /// <summary>AI 对这一段写的改动说明。只删了段内空行、或 AI 的改动没采用时为空。</summary>
        internal IReadOnlyList<string> Changes { get; }
    }

    /// <summary>一次 AI 优化的处理对象：哪一页、哪些段落、是不是因为没选中文字而处理了整页。</summary>
    internal sealed class AiTargets
    {
        private readonly HashSet<string> _selectedBlankLines;

        internal AiTargets(string pageId, IReadOnlyList<AiParagraph> paragraphs, bool wholePage,
            HashSet<string> selectedBlankLines)
        {
            PageId = pageId;
            Paragraphs = paragraphs;
            WholePage = wholePage;
            _selectedBlankLines = selectedBlankLines;
        }

        internal string PageId { get; }

        internal IReadOnlyList<AiParagraph> Paragraphs { get; }

        internal bool WholePage { get; }

        /// <summary>这个空行段落在不在处理范围里：处理整页时都在，否则只有选区里的在。</summary>
        internal bool Covers(XElement blankLine)
        {
            return WholePage || _selectedBlankLines.Contains((string)blankLine.Attribute("objectID") ?? string.Empty);
        }
    }

    /// <summary>读写 OneNote 页面：识别选区、就地替换成代码框、插入新代码框、AI 改写段落。</summary>
    internal sealed class PageEditor
    {
        /// <summary>
        /// 段落字体是这些时当成代码，不交给 AI。本插件生成的代码框每个段落都带 font-family，
        /// 用的就是插入窗口里那几个等宽字体。
        /// </summary>
        private static readonly string[] CodeFonts =
        {
            "Consolas", "NSimSun", "新宋体", "Cascadia Mono", "Cascadia Code", "Courier New", "Courier"
        };

        private readonly OneNoteApi _api;

        private static XNamespace One => OneNoteApi.One;

        internal PageEditor(OneNoteApi api)
        {
            _api = api;
        }

        /// <summary>
        /// 把当前选中的段落替换成高亮代码框。
        /// </summary>
        internal EditResult HighlightSelection(AddInSettings settings, string languageId)
        {
            var pageId = _api.GetCurrentPageId();
            if (string.IsNullOrEmpty(pageId))
            {
                return EditResult.Fail("找不到当前页面。请先在 OneNote 里打开一个页面再试。");
            }

            var page = XDocument.Parse(_api.GetPageContent(pageId, PageInfo.piSelection)).Root;
            if (page == null)
            {
                return EditResult.Fail("读取当前页面内容失败。");
            }

            var selected = FindSelectedParagraphs(page);

            if (selected.Count == 0)
            {
                return EditResult.Fail("没有检测到选中的文本。请先在页面上选中要高亮的代码，再点「高亮选中」。");
            }

            var code = string.Join("\n", selected.Select(ExtractPlainText));
            if (string.IsNullOrWhiteSpace(code))
            {
                return EditResult.Fail("选中的内容里没有文字。");
            }

            var language = LanguageRegistry.Resolve(languageId, code);
            if (language == null)
            {
                return EditResult.Fail("无法自动判断这段代码的语言。请在功能区的「语言」下拉里明确选择后重试。");
            }

            // 选中的段落必须在同一个父节点下，否则替换后的结构会很怪。
            var parent = selected[0].Parent;
            if (selected.Any(oe => oe.Parent != parent))
            {
                return EditResult.Fail("选中的内容跨越了不同的区块（比如同时选了表格内外）。请只选中同一块里的代码。");
            }

            var outline = selected[0].Ancestors(One + "Outline").FirstOrDefault();
            var outlineId = (string)outline?.Attribute("objectID");
            if (outline == null || string.IsNullOrEmpty(outlineId))
            {
                return EditResult.Fail("选中的内容不在一个可编辑的区块里，无法替换。");
            }

            var table = CodeBlockBuilder.BuildTable(code, language, settings.Theme, settings);

            // 在第一个被选中的段落位置放入代码框，再把原来那些段落删掉。
            selected[0].AddBeforeSelf(new XElement(One + "OE", table));
            foreach (var oe in selected)
            {
                oe.Remove();
            }

            var changes = BuildPageChanges(pageId, outline);
            return Submit(pageId, page, changes,
                $"已按 {language.DisplayName} 高亮 {selected.Count} 行所在的选区。");
        }

        /// <summary>
        /// 在当前页面末尾插入一个新的高亮代码框。
        /// </summary>
        internal EditResult InsertCode(string code, ILanguage language, AddInSettings settings)
        {
            if (string.IsNullOrWhiteSpace(code))
            {
                return EditResult.Fail("没有要插入的代码。");
            }

            var pageId = _api.GetCurrentPageId();
            if (string.IsNullOrEmpty(pageId))
            {
                return EditResult.Fail("找不到当前页面。请先在 OneNote 里打开一个页面再试。");
            }

            var page = XDocument.Parse(_api.GetPageContent(pageId, PageInfo.piBasic)).Root;
            if (page == null)
            {
                return EditResult.Fail("读取当前页面内容失败。");
            }

            var table = CodeBlockBuilder.BuildTable(code, language, settings.Theme, settings);
            var (x, y) = NextFreePosition(page);
            var outline = CodeBlockBuilder.BuildOutline(table, x, y, settings.CodeBlockWidth);

            var changes = BuildPageChanges(pageId, outline);
            return Submit(pageId, page, changes, $"已按 {language.DisplayName} 插入代码框。");
        }

        /// <summary>
        /// 读出「AI 优化」要处理的段落：有选中文字就只取选中的段落，否则取整页（标题加所有文本框）。
        /// 代码段落、拼不回原格式的段落、空段落都不取。
        /// </summary>
        internal EditResult ReadAiTargets(out AiTargets targets)
        {
            targets = null;

            var pageId = _api.GetCurrentPageId();
            if (string.IsNullOrEmpty(pageId))
            {
                return EditResult.Fail("找不到当前页面。请先在 OneNote 里打开一个页面再试。");
            }

            var page = XDocument.Parse(_api.GetPageContent(pageId, PageInfo.piSelection)).Root;
            if (page == null)
            {
                return EditResult.Fail("读取当前页面内容失败。");
            }

            var textParagraphs = page.Elements()
                .Where(e => e.Name == One + "Title" || e.Name == One + "Outline")
                .SelectMany(e => e.Descendants(One + "OE"))
                .Where(oe => oe.Elements(One + "T").Any())
                .ToList();

            var selected = textParagraphs.Where(IsSelectedWithText).ToList();
            var wholePage = selected.Count == 0;

            var paragraphs = new List<AiParagraph>();
            foreach (var oe in wholePage ? textParagraphs : selected)
            {
                var objectId = (string)oe.Attribute("objectID");
                if (string.IsNullOrEmpty(objectId) || IsCodeParagraph(oe))
                {
                    continue;
                }

                var rich = RichParagraph.Parse(oe);
                if (!rich.IsLossless)
                {
                    AddInLog.Info("段落里有拼不回原样的 HTML，不交给 AI：" + objectId);
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(rich.Text))
                {
                    paragraphs.Add(new AiParagraph(objectId, rich.Text));
                }
            }

            if (paragraphs.Count == 0)
            {
                return EditResult.Fail(wholePage
                    ? "当前页面上没有可以处理的文字（代码框不会交给 AI）。"
                    : "选中的内容里没有可以处理的文字（代码框不会交给 AI）。");
            }

            targets = new AiTargets(pageId, paragraphs, wholePage,
                wholePage ? new HashSet<string>() : FindSelectedBlankLines(page));
            return EditResult.Ok();
        }

        /// <summary>
        /// 选区里的空行段落的 objectID。拖选经过空行时，空段落会不会被标成选中没有把握，
        /// 所以夹在第一个和最后一个选中段落之间的空行也算：选区总是连续的一段。
        /// </summary>
        private static HashSet<string> FindSelectedBlankLines(XElement page)
        {
            var lines = page.Elements(One + "Outline").Descendants(One + "OE").ToList();
            var first = lines.FindIndex(IsSelectedWithText);
            var last = lines.FindLastIndex(IsSelectedWithText);

            return new HashSet<string>(lines
                .Where((oe, i) => (i > first && i < last) || IsSelected(oe))
                .Where(BlankLines.IsBlankLine)
                .Select(oe => (string)oe.Attribute("objectID"))
                .Where(id => !string.IsNullOrEmpty(id)));
        }

        /// <summary>
        /// 把 AI 改过的段落写回页面；removeBlankLines 时顺带删掉处理范围里多余的空行。
        ///
        /// AI 要跑好一阵，这期间用户可能还在改这一页，所以不能拿开始时读到的页面写回：
        /// 重新读一遍，按 objectID 找回每一段，文字和当初发给 AI 的一样才改，否则跳过（计入 conflicted）。
        /// 哪些是空行也按重新读到的页面算，处理期间在空行里打了字的就不会被删。
        /// 只回传有改动的那几个文本框 / 标题。applied 是真正写上去的那些。
        /// </summary>
        internal EditResult ApplyParagraphEdits(AiTargets targets, IReadOnlyList<AiParagraphEdit> edits,
            bool removeBlankLines, out List<AiParagraphEdit> applied, out int conflicted, out int removedBlankLines)
        {
            applied = new List<AiParagraphEdit>();
            conflicted = 0;
            removedBlankLines = 0;

            var pageId = targets.PageId;
            var page = XDocument.Parse(_api.GetPageContent(pageId, PageInfo.piBasic)).Root;
            if (page == null)
            {
                return EditResult.Fail("重新读取页面失败，没有写回。");
            }

            var paragraphsById = page.Descendants(One + "OE")
                .Where(oe => oe.Attribute("objectID") != null)
                .GroupBy(oe => (string)oe.Attribute("objectID"))
                .ToDictionary(g => g.Key, g => g.First());

            var changedContainers = new HashSet<XElement>();

            foreach (var edit in edits)
            {
                if (!paragraphsById.TryGetValue(edit.Source.ObjectId, out var oe))
                {
                    conflicted++;
                    continue;
                }

                var rich = RichParagraph.Parse(oe);
                if (!rich.IsLossless || rich.Text != edit.Source.Text)
                {
                    conflicted++;
                    continue;
                }

                rich.Apply(edit.NewText);
                applied.Add(edit);
                changedContainers.Add(oe.AncestorsAndSelf().First(e => e.Parent == page));
            }

            if (removeBlankLines)
            {
                removedBlankLines = BlankLines.RemoveFromPage(page, targets.Covers, changedContainers);
            }

            if (applied.Count == 0 && removedBlankLines == 0)
            {
                return EditResult.Ok();
            }

            // 按页面上的先后顺序回传，标题在文本框前面。
            var changed = page.Elements().Where(changedContainers.Contains).ToArray();
            return Submit(pageId, page, BuildPageChanges(pageId, changed),
                $"AI 已修改 {applied.Count} 段，删掉 {removedBlankLines} 个空行。");
        }

        /// <summary>
        /// 回传被改动的子树，而不是整页覆盖：整页回传会把图片、墨迹等二进制内容置于风险中。
        /// </summary>
        internal static string BuildPageChanges(string pageId, params XElement[] changedElements)
        {
            var root = new XElement(
                One + "Page",
                new XAttribute(XNamespace.Xmlns + "one", OneNoteApi.OneNs),
                new XAttribute("ID", pageId),
                changedElements);

            return root.ToString(SaveOptions.DisableFormatting);
        }

        /// <summary>
        /// 提交改动。先带 lastModifiedTime 做冲突检测；若因为期间 OneNote 自己又保存了一次而失败，
        /// 退回不校验重试一次——这种情况下页面内容是我们刚读到的，覆盖是安全的。
        /// </summary>
        private EditResult Submit(string pageId, XElement page, string changesXml, string successMessage)
        {
            var lastModified = ParseLastModified(page);

            try
            {
                _api.UpdatePageContent(changesXml, lastModified);
                return EditResult.Ok(successMessage);
            }
            catch (Exception ex) when (lastModified != DateTime.MinValue)
            {
                AddInLog.Warn("带时间戳回写失败，退回不校验重试。", ex);

                try
                {
                    _api.UpdatePageContent(changesXml, DateTime.MinValue);
                    return EditResult.Ok(successMessage);
                }
                catch (Exception retryEx)
                {
                    AddInLog.Error("回写页面失败。pageId=" + pageId, retryEx);
                    return EditResult.Fail("写回 OneNote 失败：" + retryEx.Message);
                }
            }
            catch (Exception ex)
            {
                AddInLog.Error("回写页面失败。pageId=" + pageId, ex);
                return EditResult.Fail("写回 OneNote 失败：" + ex.Message);
            }
        }

        private static DateTime ParseLastModified(XElement page)
        {
            var raw = (string)page.Attribute("lastModifiedTime");
            if (string.IsNullOrEmpty(raw))
            {
                return DateTime.MinValue;
            }

            // 只能用 RoundtripKind，不能再或上 AdjustToUniversal——这两个是互斥的，
            // 组合起来 TryParse 会直接抛 ArgumentException。OneNote 给的值带 Z，
            // RoundtripKind 已经足够解析成 Kind=Utc 的时间。
            return DateTime.TryParse(raw, CultureInfo.InvariantCulture,
                DateTimeStyles.RoundtripKind, out var parsed)
                ? parsed
                : DateTime.MinValue;
        }

        /// <summary>找一个不压住已有内容的位置：所有 Outline 的底边取最大值再往下留点空。</summary>
        private static (double X, double Y) NextFreePosition(XElement page)
        {
            const double DefaultX = 36;
            const double DefaultY = 86;
            const double Gap = 20;

            var bottom = page.Descendants(One + "Outline")
                .Select(outline =>
                {
                    var y = ReadDouble(outline.Element(One + "Position"), "y");
                    var height = ReadDouble(outline.Element(One + "Size"), "height");
                    return y + height;
                })
                .DefaultIfEmpty(0)
                .Max();

            return (DefaultX, bottom > 0 ? bottom + Gap : DefaultY);
        }

        private static double ReadDouble(XElement element, string attributeName)
        {
            var raw = (string)element?.Attribute(attributeName);
            return double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
                ? value
                : 0;
        }

        /// <summary>
        /// 找出选区涉及的段落，碰到一点就算整行。
        ///
        /// 不能只认 OE 上的 selected="all"：本机实测，在文本框里拖选几行时，OneNote 只把段落里的
        /// one:T 标成 "all"，段落本身一律标 "partial"——哪怕整行文字都选上了。只有点文本框边框、
        /// Ctrl+A 这类整体选中才会把 OE 标成 "all"，所以以前只能高亮整个框。
        /// 选区从行中间开始或结束时，OneNote 会把那行的 one:T 在边界处拆开，只有选中那截标 "all"，
        /// 这里同样按整行算，代码只高亮半行没有意义。
        /// </summary>
        internal static List<XElement> FindSelectedParagraphs(XElement page)
        {
            return page.Descendants(One + "OE")
                .Where(oe => oe.Elements(One + "T").Any())
                .Where(IsSelected)
                .ToList();
        }

        private static bool IsSelected(XElement oe)
        {
            return (string)oe.Attribute("selected") == "all"
                   || oe.Elements(One + "T").Any(t => (string)t.Attribute("selected") == "all");
        }

        /// <summary>
        /// AI 优化用的选区判定。和 <see cref="FindSelectedParagraphs"/> 一样碰到一点就算整段，
        /// 但只认真有文字被选中的：光标只是停在某一行时，OneNote 也可能把一个空的 one:T 标成 "all"，
        /// 那种情况应当按「没选中」处理整页，而不是只处理光标所在的那一段。
        /// </summary>
        private static bool IsSelectedWithText(XElement oe)
        {
            return (string)oe.Attribute("selected") == "all"
                   || oe.Elements(One + "T").Any(t => (string)t.Attribute("selected") == "all"
                                                     && !string.IsNullOrWhiteSpace(OneNoteHtmlEncoder.DecodeToPlainText(t.Value)));
        }

        /// <summary>
        /// 段落样式里的 font-family 是等宽字体就当成代码。本插件写的是 OE 上的 style；
        /// 万一 OneNote 回存时把字体挪到了各个 one:T 上，全部 one:T 都是等宽字体也算。
        /// </summary>
        internal static bool IsCodeParagraph(XElement oe)
        {
            if (IsMonospaceStyle((string)oe.Attribute("style")))
            {
                return true;
            }

            var runs = oe.Elements(One + "T").ToList();
            return runs.Count > 0 && runs.All(t => IsMonospaceStyle((string)t.Attribute("style")));
        }

        private static bool IsMonospaceStyle(string style)
        {
            if (string.IsNullOrEmpty(style))
            {
                return false;
            }

            foreach (var declaration in style.Split(';'))
            {
                var colon = declaration.IndexOf(':');
                if (colon < 0 || !string.Equals(declaration.Substring(0, colon).Trim(), "font-family",
                        StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var family = declaration.Substring(colon + 1).Split(',')[0].Trim().Trim('\'', '"');
                return CodeFonts.Any(f => string.Equals(f, family, StringComparison.OrdinalIgnoreCase));
            }

            return false;
        }

        /// <summary>把一个 one:OE 里的文字还原成纯文本。</summary>
        private static string ExtractPlainText(XElement oe)
        {
            var parts = oe.Elements(One + "T").Select(t => OneNoteHtmlEncoder.DecodeToPlainText(t.Value));
            return string.Concat(parts);
        }

        /// <summary>供诊断用：导出当前选区解析出来的纯文本。</summary>
        internal IReadOnlyList<string> DebugReadSelectionLines()
        {
            var pageId = _api.GetCurrentPageId();
            if (string.IsNullOrEmpty(pageId))
            {
                return Array.Empty<string>();
            }

            var page = XDocument.Parse(_api.GetPageContent(pageId, PageInfo.piSelection)).Root;
            return page == null
                ? Array.Empty<string>()
                : FindSelectedParagraphs(page)
                    .Select(ExtractPlainText)
                    .ToList();
        }
    }
}
