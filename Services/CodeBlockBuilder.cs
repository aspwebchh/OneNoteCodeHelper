using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using OneNoteCodeHelper.Highlighting;
using OneNoteCodeHelper.Highlighting.Themes;

namespace OneNoteCodeHelper.Services
{
    /// <summary>
    /// 生成 OneNote 的「代码框」XML。
    ///
    /// OneNote 没有代码块这种元素，所以用一个单行单列的表格来仿：
    /// one:Cell 的 shadingColor 给底色，bordersVisible 给边框，单元格里每行源码一个 one:OE。
    /// </summary>
    internal static class CodeBlockBuilder
    {
        private static XNamespace One => OneNoteApi.One;

        /// <summary>标在代码框上的记号，用来认出「这块是本插件生成的」。</summary>
        internal const string MarkerAttribute = "author";

        /// <summary>
        /// 把源码渲染成一个 one:Table。
        /// </summary>
        internal static XElement BuildTable(
            string code, ILanguage language, CodeTheme theme, AddInSettings settings)
        {
            // 去掉末尾空白：粘贴来的代码几乎都带一个结尾换行，不去掉会在代码框里多出一个空行。
            // 中间的空行不受影响，那是有意义的。
            code = (code ?? string.Empty).TrimEnd();

            var tokens = language.Tokenize(code);
            var lines = OneNoteHtmlEncoder.BuildLines(code, tokens, theme, settings.TabWidth);

            // 字体放在 OE 上而不是每个 span 里，省体积；默认文字色也一并给上，
            // 这样深色主题里没被 span 包住的字符（空白占位等）也不会掉回黑色。
            var oeStyle = string.Format(
                CultureInfo.InvariantCulture,
                "font-family:{0};font-size:{1}pt;color:{2}",
                settings.FontFamily,
                settings.FontSize,
                theme.DefaultStyle.Color);

            var children = new XElement(
                One + "OEChildren",
                lines.Select(line => new XElement(
                    One + "OE",
                    new XAttribute("style", oeStyle),
                    new XElement(One + "T", new XCData(line)))));

            return new XElement(
                One + "Table",
                new XAttribute("bordersVisible", settings.ShowBorders ? "true" : "false"),
                new XAttribute("hasHeaderRow", "false"),
                new XElement(
                    One + "Columns",
                    new XElement(
                        One + "Column",
                        new XAttribute("index", 0),
                        new XAttribute("width", settings.CodeBlockWidth.ToString("0.#", CultureInfo.InvariantCulture)))),
                new XElement(
                    One + "Row",
                    new XElement(
                        One + "Cell",
                        new XAttribute("shadingColor", theme.Background),
                        children)));
        }

        /// <summary>
        /// 把代码框包进一个新的 one:Outline，用于「插入代码」。
        /// 不带 objectID 表示新建；位置给在页面已有内容的下方。
        /// </summary>
        internal static XElement BuildOutline(XElement table, double x, double y, double width)
        {
            return new XElement(
                One + "Outline",
                new XElement(
                    One + "Position",
                    new XAttribute("x", x.ToString("0.#", CultureInfo.InvariantCulture)),
                    new XAttribute("y", y.ToString("0.#", CultureInfo.InvariantCulture)),
                    new XAttribute("z", 0)),
                new XElement(
                    One + "Size",
                    new XAttribute("width", width.ToString("0.#", CultureInfo.InvariantCulture)),
                    new XAttribute("height", "100")),
                new XElement(One + "OEChildren", new XElement(One + "OE", table)));
        }
    }
}
