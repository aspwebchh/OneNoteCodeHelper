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

    /// <summary>读写 OneNote 页面：识别选区、就地替换成代码框、插入新代码框。</summary>
    internal sealed class PageEditor
    {
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
        /// 回传被改动的子树，而不是整页覆盖：整页回传会把图片、墨迹等二进制内容置于风险中。
        /// </summary>
        internal static string BuildPageChanges(string pageId, XElement changedElement)
        {
            var root = new XElement(
                One + "Page",
                new XAttribute(XNamespace.Xmlns + "one", OneNoteApi.OneNs),
                new XAttribute("ID", pageId),
                changedElement);

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
                .Where(oe => (string)oe.Attribute("selected") == "all"
                    || oe.Elements(One + "T").Any(t => (string)t.Attribute("selected") == "all"))
                .ToList();
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
