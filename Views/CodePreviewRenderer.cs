using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using OneNoteCodeHelper.Highlighting;
using OneNoteCodeHelper.Highlighting.Themes;
using OneNoteCodeHelper.Services;

namespace OneNoteCodeHelper.Views
{
    /// <summary>
    /// 把同一份 token 流渲染成 WPF 的 FlowDocument，用于插入前的预览。
    /// 走的是和 OneNote 输出完全相同的词法分析与配色，所见即所得。
    ///
    /// FlowDocument 自己没有边框，所以所有行都装进一个 Section 里，
    /// 由它来仿 OneNote 那个单元格：底色、内边距和（可选的）边框都挂在它身上。
    /// </summary>
    internal static class CodePreviewRenderer
    {
        internal static FlowDocument Build(
            string code, ILanguage language, CodeTheme theme, AddInSettings settings)
        {
            var document = new FlowDocument
            {
                FontFamily = new FontFamily(settings.FontFamily),
                FontSize = ToDeviceUnits(settings.FontSize),
                Foreground = ToBrush(theme.DefaultStyle.Color),
                PagePadding = new Thickness(0),
                PageWidth = double.NaN
            };

            if (string.IsNullOrEmpty(code))
            {
                return document;
            }

            var frame = new Section
            {
                Background = ToBrush(theme.Background),
                BorderBrush = ToBrush(theme.Border),
                BorderThickness = new Thickness(settings.ShowBorders ? 1 : 0),
                Padding = new Thickness(10)
            };
            document.Blocks.Add(frame);

            foreach (var line in SplitLines(code, language.Tokenize(code), settings.TabWidth))
            {
                var paragraph = new Paragraph { Margin = new Thickness(0) };

                foreach (var run in line)
                {
                    paragraph.Inlines.Add(run);
                }

                // 空行也要占一行高度，否则预览和实际插入的行数对不上。
                if (paragraph.Inlines.Count == 0)
                {
                    paragraph.Inlines.Add(new Run(" "));
                }

                frame.Blocks.Add(paragraph);
            }

            return document;
        }

        /// <summary>OneNote 的字号单位是磅，WPF 是设备无关像素，按 96/72 换算才能看着一致。</summary>
        internal static double ToDeviceUnits(double points)
        {
            return points * 96.0 / 72.0;
        }

        private static Section FindFrame(FlowDocument document)
        {
            return document.Blocks.FirstBlock as Section;
        }

        private static IEnumerable<List<Run>> SplitLines(string code, IEnumerable<Token> tokens, int tabWidth)
        {
            var line = new List<Run>();
            var column = 0;

            foreach (var token in tokens)
            {
                var segmentStart = token.Start;

                for (var i = token.Start; i < token.End; i++)
                {
                    var ch = code[i];
                    if (ch != '\n' && ch != '\r')
                    {
                        continue;
                    }

                    AddRun(line, code, segmentStart, i, token.Kind, ref column, tabWidth);
                    yield return line;

                    line = new List<Run>();
                    column = 0;

                    if (ch == '\r' && i + 1 < code.Length && code[i + 1] == '\n')
                    {
                        i++;
                    }

                    segmentStart = i + 1;
                }

                AddRun(line, code, segmentStart, token.End, token.Kind, ref column, tabWidth);
            }

            yield return line;
        }

        private static void AddRun(
            List<Run> line, string code, int start, int end, TokenKind kind, ref int column, int tabWidth)
        {
            if (end <= start)
            {
                return;
            }

            var text = ExpandTabs(code.Substring(start, end - start), ref column, tabWidth);
            line.Add(new Run(text) { Tag = kind });
        }

        private static string ExpandTabs(string text, ref int column, int tabWidth)
        {
            if (text.IndexOf('\t') < 0)
            {
                column += text.Length;
                return text;
            }

            var builder = new System.Text.StringBuilder(text.Length + tabWidth);
            foreach (var ch in text)
            {
                if (ch == '\t')
                {
                    var spaces = tabWidth - (column % tabWidth);
                    builder.Append(' ', spaces);
                    column += spaces;
                }
                else
                {
                    builder.Append(ch);
                    column++;
                }
            }

            return builder.ToString();
        }

        /// <summary>给已经排好的 Run 上色。分成两步是为了换主题时不必重新做词法分析。</summary>
        internal static void ApplyTheme(FlowDocument document, CodeTheme theme)
        {
            document.Foreground = ToBrush(theme.DefaultStyle.Color);

            var frame = FindFrame(document);
            if (frame == null)
            {
                return;
            }

            frame.Background = ToBrush(theme.Background);
            frame.BorderBrush = ToBrush(theme.Border);

            foreach (var block in frame.Blocks)
            {
                if (!(block is Paragraph paragraph))
                {
                    continue;
                }

                foreach (var inline in paragraph.Inlines)
                {
                    if (!(inline is Run run) || !(run.Tag is TokenKind kind))
                    {
                        continue;
                    }

                    var style = theme.StyleFor(kind);
                    run.Foreground = ToBrush(style.Color);
                    run.FontWeight = style.Bold ? FontWeights.Bold : FontWeights.Normal;
                    run.FontStyle = style.Italic ? FontStyles.Italic : FontStyles.Normal;
                }
            }
        }

        private static Brush ToBrush(string hex)
        {
            try
            {
                var value = uint.Parse(hex.TrimStart('#'), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                var brush = new SolidColorBrush(Color.FromRgb(
                    (byte)((value >> 16) & 0xFF),
                    (byte)((value >> 8) & 0xFF),
                    (byte)(value & 0xFF)));

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
