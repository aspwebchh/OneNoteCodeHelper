using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using OneNoteCodeHelper.Services;
using OneNoteCodeHelper.Services.Agent;
using OneNoteCodeHelper.Services.Markdown;

namespace OneNoteCodeHelper.Views
{
    /// <summary>
    /// 把解析好的 Markdown 画成 FlowDocument，用于「插入代码」窗口的预览。和写到 OneNote 的是同一份块：
    /// 标题、正文、引用的字号颜色取同一套段落预设，代码框直接用 <see cref="CodePreviewRenderer"/> 画。
    /// 列表符号、待办框按 OneNote 的样子用文字前缀示意。
    /// </summary>
    internal static class MarkdownPreviewRenderer
    {
        /// <summary>每一层缩进的宽度（设备无关像素），接近 OneNote 里按一次 Tab 的缩进。</summary>
        private const double IndentStep = 24;

        private const string LinkColor = "#0563C1";
        private const string TableBorder = "#BFBFBF";

        internal static FlowDocument Build(MarkdownDocument doc, AddInSettings settings, AgentOptions options, int maxBlocks)
        {
            var body = ParagraphStyles.Appearance("body");
            var document = new FlowDocument
            {
                FontFamily = new FontFamily(options.FontFamily),
                FontSize = CodePreviewRenderer.ToDeviceUnits(body.Size),
                Foreground = ToBrush(body.Color),
                PagePadding = new Thickness(0),
                PageWidth = double.NaN
            };

            // 各层编号列表当前的编号；每组第一项带着起始编号，之后的项接着加一。
            var numbers = new Dictionary<int, int>();
            foreach (var block in doc.Blocks.Take(maxBlocks))
            {
                var left = block.Level * IndentStep;
                switch (block.Kind)
                {
                    case MarkdownBlockKind.Blank:
                        document.Blocks.Add(new Paragraph(new Run(" ")) { Margin = new Thickness(0) });
                        break;
                    case MarkdownBlockKind.Code:
                        document.Blocks.Add(Code(block, settings, left));
                        break;
                    case MarkdownBlockKind.Table:
                        document.Blocks.Add(Table(block, settings, left));
                        break;
                    default:
                        var look = ParagraphStyles.Appearance(block.Preset);
                        var paragraph = new Paragraph
                        {
                            FontSize = CodePreviewRenderer.ToDeviceUnits(look.Size),
                            Foreground = ToBrush(look.Color),
                            FontWeight = look.Bold || block.BoldText ? FontWeights.Bold : FontWeights.Normal,
                            Margin = new Thickness(left, Points(look.SpaceBefore), 0, Points(look.SpaceAfter))
                        };
                        // 待办框用 Segoe UI Symbol 的单色字形：默认字体里 ☑ 会落到彩色 emoji，和 ☐ 不一样大。
                        if (block.Todo != null)
                            paragraph.Inlines.Add(new Run(block.Todo.Value ? "☑ " : "☐ ") { FontFamily = new FontFamily("Segoe UI Symbol"), FontWeight = FontWeights.Normal });
                        var prefix = Prefix(block, numbers);
                        if (prefix.Length > 0) paragraph.Inlines.Add(new Run(prefix) { FontWeight = FontWeights.Normal });
                        AddRuns(paragraph.Inlines, block.Inline, settings);
                        if (paragraph.Inlines.Count == 0) paragraph.Inlines.Add(new Run(" "));
                        document.Blocks.Add(paragraph);
                        break;
                }
            }

            return document;
        }

        /// <summary>列表符号：圆点或编号。</summary>
        private static string Prefix(MarkdownBlock block, Dictionary<int, int> numbers)
        {
            if (block.ListKind == "bullet") return "• ";
            if (block.ListKind != "number") return string.Empty;
            numbers.TryGetValue(block.Level, out var current);
            current = block.Number ?? current + 1;
            numbers[block.Level] = current;
            return current.ToString(CultureInfo.InvariantCulture) + ". ";
        }

        private static void AddRuns(InlineCollection inlines, IEnumerable<MarkdownRun> runs, AddInSettings settings)
        {
            foreach (var run in runs)
            {
                if (run.Break)
                {
                    inlines.Add(new LineBreak());
                    continue;
                }

                var element = new Run(run.Text);
                if (run.Bold) element.FontWeight = FontWeights.Bold;
                if (run.Italic) element.FontStyle = FontStyles.Italic;
                if (run.Code) element.FontFamily = new FontFamily(settings.FontFamily);
                var decorations = new TextDecorationCollection();
                if (run.Strike) decorations.Add(TextDecorations.Strikethrough);
                if (run.Href != null)
                {
                    element.Foreground = ToBrush(LinkColor);
                    decorations.Add(TextDecorations.Underline);
                }

                if (decorations.Count > 0) element.TextDecorations = decorations;
                inlines.Add(element);
            }
        }

        /// <summary>代码框和「插入代码」的预览完全一样，只是挪进这份文档、按层级缩进。</summary>
        private static Block Code(MarkdownBlock block, AddInSettings settings, double left)
        {
            var theme = settings.Theme;
            var code = CodePreviewRenderer.Build(block.Code.TrimEnd(), block.Language, theme, settings);
            CodePreviewRenderer.ApplyTheme(code, theme);
            var section = code.Blocks.FirstBlock as Section ?? new Section(new Paragraph(new Run(" ")))
            {
                Background = ToBrush(theme.Background)
            };
            code.Blocks.Remove(section);
            section.FontFamily = code.FontFamily;
            section.FontSize = code.FontSize;
            section.Foreground = code.Foreground;
            section.Margin = new Thickness(left, 4, 0, 4);
            return section;
        }

        private static Block Table(MarkdownBlock block, AddInSettings settings, double left)
        {
            var table = new Table { CellSpacing = 0, Margin = new Thickness(left, 4, 0, 4) };
            var columns = block.Rows.Max(r => r.Count);
            for (var c = 0; c < columns; c++) table.Columns.Add(new TableColumn());
            var group = new TableRowGroup();
            for (var r = 0; r < block.Rows.Count; r++)
            {
                var row = new TableRow();
                for (var c = 0; c < columns; c++)
                {
                    var paragraph = new Paragraph { Margin = new Thickness(0) };
                    AddRuns(paragraph.Inlines, block.Rows[r][c], settings);
                    if (r == 0) paragraph.FontWeight = FontWeights.Bold;
                    var alignment = block.Alignments[c];
                    paragraph.TextAlignment = alignment == "center" ? TextAlignment.Center : alignment == "right" ? TextAlignment.Right : TextAlignment.Left;
                    var cell = new TableCell(paragraph)
                    {
                        BorderBrush = ToBrush(TableBorder),
                        BorderThickness = new Thickness(0.5),
                        Padding = new Thickness(6, 3, 6, 3)
                    };
                    if (r == 0) cell.Background = ToBrush(TableLook.Shadings[1]);
                    row.Cells.Add(cell);
                }

                group.Rows.Add(row);
            }

            table.RowGroups.Add(group);
            return table;
        }

        private static double Points(double points) => CodePreviewRenderer.ToDeviceUnits(points);

        private static Brush ToBrush(string hex)
        {
            try
            {
                var value = uint.Parse(hex.TrimStart('#'), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                var brush = new SolidColorBrush(Color.FromRgb((byte)((value >> 16) & 0xFF), (byte)((value >> 8) & 0xFF), (byte)(value & 0xFF)));
                brush.Freeze();
                return brush;
            }
            catch (Exception)
            {
                return Brushes.Black;
            }
        }
    }
}
