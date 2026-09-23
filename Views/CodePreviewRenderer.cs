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
    /// </summary>
    internal static class CodePreviewRenderer
    {
        internal static FlowDocument Build(
            string code, ILanguage language, CodeTheme theme, AddInSettings settings)
        {
            var document = new FlowDocument
            {
                FontFamily = new FontFamily(settings.FontFamily),
                // OneNote 的字号单位是磅，WPF 是设备无关像素，按 96/72 换算才能看着一致。
                FontSize = settings.FontSize * 96.0 / 72.0,
                Background = ToBrush(theme.Background),
                Foreground = ToBrush(theme.DefaultStyle.Color),
                PagePadding = new Thickness(10),
                PageWidth = double.NaN
            };

            if (string.IsNullOrEmpty(code))
            {
                return document;
            }

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

                document.Blocks.Add(paragraph);
            }

            return document;
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
            document.Background = ToBrush(theme.Background);
            document.Foreground = ToBrush(theme.DefaultStyle.Color);

            foreach (var block in document.Blocks)
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
