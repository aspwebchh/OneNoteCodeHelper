using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using OneNoteCodeHelper.Services.Agent;

namespace OneNoteCodeHelper.Services.Markdown
{
    /// <summary>
    /// 行内 Markdown：粗体、斜体、删除线、行内代码、链接（含徽章）、自动链接、裸网址、段内换行、反斜杠转义和 HTML 实体。
    /// 强调的配对规则与 strip_markdown 共用 <see cref="AgentMarkdown.InlineEmphasis"/>：snake_case、2 * 3 不当强调。
    /// 图片不下载，HTML 不执行，都按文字保留并记一条警告。
    /// </summary>
    internal static class MarkdownInline
    {
        /// <summary>和 MCP 导入一样只接受这几种链接，javascript: 之类按文字保留。</summary>
        private static readonly string[] Schemes = { "http", "https", "mailto", "onenote" };

        /// <summary>链接文字最多看这么长：每个 [ 都要往后找配对的 ]，满行的 [ 不能拖成平方级。</summary>
        private const int MaxLabel = 1000;

        /// <summary>紧跟在链接文字的 ] 后面的「(地址 "标题")」。</summary>
        private static readonly Regex Destination = new Regex(
            @"\G\((?<url><[^<>\n]*>|[^\s()<>]*(?:\([^\s()<>]*\)[^\s()<>]*)*)(?:[ \t]+(?:""[^""\n]*""|'[^'\n]*'))?[ \t]*\)");
        private static readonly Regex Entity = new Regex(@"\G&(?:#[0-9]{1,7}|#[xX][0-9A-Fa-f]{1,6}|[A-Za-z][A-Za-z0-9]{1,31});");
        private static readonly Regex AutoLink = new Regex(@"\G<(?<url>(?:https?|mailto|onenote):[^\s<>]+)>", RegexOptions.IgnoreCase);
        private static readonly Regex AutoEmail = new Regex(@"\G<(?<url>[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})>");
        private static readonly Regex LineBreak = new Regex(@"\G<br[ \t]*/?>", RegexOptions.IgnoreCase);
        private static readonly Regex HtmlTag = new Regex(@"\G(?:<!--.*?-->|</?[A-Za-z][A-Za-z0-9-]*(?:[ \t][^<>]*)?/?>)");
        /// <summary>裸网址到空白、尖括号、引号、汉字或中文标点为止（中文里网址后面常常不空格）；结尾的英文标点另外去掉。</summary>
        private static readonly Regex BareUrl = new Regex(
            @"\Ghttps?://[^\s<>""'`\p{IsCJKSymbolsandPunctuation}\p{IsCJKUnifiedIdeographsExtensionA}\p{IsCJKUnifiedIdeographs}\p{IsHalfwidthandFullwidthForms}]+", RegexOptions.IgnoreCase);

        /// <summary>解析一行行内文字。不支持的内容记到 doc 的警告里。</summary>
        internal static List<MarkdownRun> Parse(string text, MarkdownDocument doc)
        {
            var n = text.Length;
            var masked = text.ToCharArray();
            var removed = new bool[n];
            var code = new bool[n];
            var breaks = new bool[n];
            var href = new string[n];
            var replaced = new Dictionary<int, string>();

            void Remove(int start, int length)
            {
                for (var k = start; k < start + length; k++) { removed[k] = true; masked[k] = '\0'; }
            }
            void Mask(int start, int length)
            {
                for (var k = start; k < start + length; k++) masked[k] = '\0';
            }
            void SetLink(int start, int length, string url)
            {
                for (var k = start; k < start + length; k++) href[k] = url;
            }

            // 链接、图片的文字里照常处理转义、实体、行内代码和强调；扫到文字末尾（End）就跳过后面的「](地址)」（到 Jump）。
            // 图片可以在链接文字里（徽章 [![说明](图片)](链接)），所以是一个栈；链接里不再套链接，图片说明里什么都不认。
            var labels = new Stack<(int End, int Jump, bool Image)>();
            for (var i = 0; i < n;)
            {
                if (labels.Count > 0 && i >= labels.Peek().End)
                {
                    i = Math.Max(i, labels.Pop().Jump);
                    continue;
                }

                var limit = labels.Count > 0 ? labels.Peek().End : n;
                var ch = text[i];
                if (ch == '\\' && i + 1 < n && IsAsciiPunctuation(text[i + 1]))
                {
                    Remove(i, 1);
                    Mask(i + 1, 1);
                    i += 2;
                    continue;
                }

                if (ch == '&')
                {
                    // 认得的实体换成对应的字（&amp; &lt; &nbsp; &#169;），解出来的字不再参与强调配对；认不出的按原文。
                    var entity = Entity.Match(text, i);
                    var decoded = entity.Success ? WebUtility.HtmlDecode(entity.Value) : null;
                    if (decoded != null && decoded != entity.Value && IsPlainText(decoded))
                    {
                        Mask(i, 1);
                        Remove(i + 1, entity.Length - 1);
                        replaced[i] = decoded;
                        i += entity.Length;
                        continue;
                    }
                }

                if (ch == '`')
                {
                    var ticks = Run(text, i, '`');
                    var close = FindTicks(text, i + ticks, ticks, limit);
                    if (close < 0)
                    {
                        Mask(i, ticks);
                        i += ticks;
                        continue;
                    }

                    var start = i + ticks;
                    var length = close - start;
                    // CommonMark：两头各有一个空格、中间不全是空格时去掉这两个空格，方便写 `` `a` ``。
                    if (length >= 2 && text[start] == ' ' && text[close - 1] == ' ' && text.Substring(start, length).Trim().Length > 0)
                    {
                        Remove(start, 1);
                        Remove(close - 1, 1);
                    }
                    Remove(i, ticks);
                    Remove(close, ticks);
                    for (var k = start; k < close; k++) { code[k] = true; masked[k] = '\0'; }
                    i = close + ticks;
                    continue;
                }

                var isImage = ch == '!' && i + 1 < n && text[i + 1] == '[';
                if (isImage ? labels.All(l => !l.Image) : ch == '[' && labels.Count == 0)
                {
                    var open = isImage ? i + 1 : i;
                    var close = LabelEnd(text, open, limit);
                    var link = close < 0 ? Match.Empty : Destination.Match(text, close + 1);
                    if (link.Success && link.Index + link.Length <= limit)
                    {
                        // 图片不下载：地址能用时变成链接，文字取替代文字。链接地址不能用（相对路径、javascript: 等）时只留文字。
                        // 文字为空时用地址当文字。徽章里的图片沿用外面链接的地址（走到这里还在文字里，外面只能是链接）。
                        var labelStart = open + 1;
                        var labelLength = close - labelStart;
                        var url = Url(link.Groups["url"].Value);
                        var allowed = Allowed(url);
                        var linked = allowed && labels.Count == 0;
                        if (isImage) doc.Warn("image");
                        else if (!allowed) doc.Warn("link");
                        Remove(i, labelStart - i);
                        var end = link.Index + link.Length;
                        Remove(close, end - close);
                        if (labelLength == 0)
                        {
                            replaced[end - 1] = url;
                            if (linked) href[end - 1] = url;
                            i = end;
                            continue;
                        }

                        if (linked) SetLink(labelStart, labelLength, url);
                        labels.Push((close, end, isImage));
                        i = labelStart;
                        continue;
                    }
                }

                if (ch == '<')
                {
                    var auto = AutoLink.Match(text, i);
                    var email = auto.Success ? Match.Empty : AutoEmail.Match(text, i);
                    if (auto.Success || email.Success)
                    {
                        var m = auto.Success ? auto : email;
                        var url = (email.Success ? "mailto:" : string.Empty) + m.Groups["url"].Value;
                        Remove(i, 1);
                        Remove(i + m.Length - 1, 1);
                        Mask(i + 1, m.Length - 2);
                        SetLink(i + 1, m.Length - 2, url);
                        i += m.Length;
                        continue;
                    }

                    var br = LineBreak.Match(text, i);
                    if (br.Success)
                    {
                        Remove(i, br.Length);
                        breaks[i] = true;
                        i += br.Length;
                        continue;
                    }

                    var tag = HtmlTag.Match(text, i);
                    if (tag.Success)
                    {
                        // 原样当文字留着；遮住里面的星号、下划线，免得和外面的配成强调。
                        doc.Warn("html");
                        Mask(i, tag.Length);
                        i += tag.Length;
                        continue;
                    }
                }

                // 前面紧挨着英文字母、数字的不算（xhttp://），汉字后面的照认：中文里网址前后常常不空格。
                // 链接文字里的网址不再单独成链接。
                if ((ch == 'h' || ch == 'H') && labels.Count == 0 && (i == 0 || !(text[i - 1] < 128 && char.IsLetterOrDigit(text[i - 1]))))
                {
                    var bare = BareUrl.Match(text, i);
                    if (bare.Success)
                    {
                        var length = TrimUrl(bare.Value);
                        if (length > bare.Value.IndexOf("://", StringComparison.Ordinal) + 3)
                        {
                            Mask(i, length);
                            SetLink(i, length, text.Substring(i, length));
                            i += length;
                            continue;
                        }
                    }
                }

                i++;
            }

            var bold = new bool[n];
            var italic = new bool[n];
            var strike = new bool[n];
            var emphasis = AgentMarkdown.InlineEmphasis(masked);
            foreach (var (start, length) in emphasis.Marks) Remove(start, length);
            foreach (var (start, length, css) in emphasis.Formats)
            {
                for (var k = start; k < start + length; k++)
                {
                    if (css.TryGetValue("font-weight", out var weight) && weight == "bold") bold[k] = true;
                    if (css.TryGetValue("font-style", out var style) && style == "italic") italic[k] = true;
                    if (css.TryGetValue("text-decoration", out var decoration) && decoration == "line-through") strike[k] = true;
                }
            }

            var runs = new List<MarkdownRun>();
            for (var k = 0; k < n; k++)
            {
                if (breaks[k])
                {
                    runs.Add(new MarkdownRun { Break = true });
                    continue;
                }

                string piece;
                if (replaced.TryGetValue(k, out var replacement)) piece = replacement;
                else if (removed[k]) continue;
                else piece = text[k].ToString();

                var run = new MarkdownRun
                {
                    Text = piece, Bold = bold[k], Italic = italic[k], Strike = strike[k], Code = code[k], Href = href[k]
                };
                var last = runs.Count > 0 ? runs[runs.Count - 1] : null;
                if (last != null && last.SameFormat(run)) last.Text += piece;
                else runs.Add(run);
            }

            return runs;
        }

        /// <summary>
        /// 生成 one:T 里的 HTML：文字转义、行首空白用硬空格顶住，格式写在 span 的 style 上，链接用 a，换行用 br。
        /// 只用 span/a/br 和 style/href，Agent 之后还能照常读写这些段落。
        /// </summary>
        internal static string ToHtml(IEnumerable<MarkdownRun> runs, string codeFont, bool bold = false)
        {
            var html = new StringBuilder();
            var atLineStart = true;
            foreach (var run in runs)
            {
                if (run.Break)
                {
                    html.Append("<br>");
                    atLineStart = true;
                    continue;
                }

                if (run.Text.Length == 0) continue;
                var text = OneNoteHtmlEncoder.EncodeText(run.Text, ref atLineStart);
                var css = new List<string>();
                if (run.Bold || bold) css.Add("font-weight:bold");
                if (run.Italic) css.Add("font-style:italic");
                if (run.Strike) css.Add("text-decoration:line-through");
                if (run.Code) css.Add("font-family:" + codeFont);
                if (css.Count > 0) text = "<span style='" + string.Join(";", css) + "'>" + text + "</span>";
                if (run.Href != null) text = "<a href=\"" + WebUtility.HtmlEncode(run.Href) + "\">" + text + "</a>";
                html.Append(text);
            }

            return html.ToString();
        }

        private static bool IsAsciiPunctuation(char c) => c < 128 && char.IsPunctuation(c) || "$+<=>^`|~".IndexOf(c) >= 0;

        private static int Run(string text, int start, char c)
        {
            var end = start;
            while (end < text.Length && text[end] == c) end++;
            return end - start;
        }

        /// <summary>在 limit 之前找下一串正好 ticks 个反引号的位置，前后不能紧挨着别的反引号。</summary>
        private static int FindTicks(string text, int from, int ticks, int limit)
        {
            for (var i = from; i < limit;)
            {
                if (text[i] != '`') { i++; continue; }
                var length = Math.Min(Run(text, i, '`'), limit - i);
                if (length == ticks) return i;
                i += length;
            }

            return -1;
        }

        /// <summary>
        /// 从 text[open] 的 [ 找配对的 ]：里面的方括号要成对（徽章里的图片），反斜杠转义的和行内代码里的不算。
        /// 返回 ] 的下标；limit 之前（最多 <see cref="MaxLabel"/> 个字）找不到返回 -1。
        /// </summary>
        private static int LabelEnd(string text, int open, int limit)
        {
            limit = Math.Min(limit, open + MaxLabel);
            var depth = 0;
            for (var i = open; i < limit; i++)
            {
                var c = text[i];
                if (c == '\\')
                {
                    i++;
                    continue;
                }

                if (c == '`')
                {
                    var ticks = Run(text, i, '`');
                    var close = FindTicks(text, i + ticks, ticks, limit);
                    i = (close >= 0 ? close : i) + ticks - 1;
                    continue;
                }

                if (c == '[') depth++;
                else if (c == ']' && --depth == 0) return i;
            }

            return -1;
        }

        /// <summary>实体解出来的字能不能放进正文：控制字符、落单的代理项和 XML 不允许的字都不行（写 one:T 时会出错）。</summary>
        private static bool IsPlainText(string text)
        {
            if (text.Length == 0) return false;
            for (var k = 0; k < text.Length; k++)
            {
                if (k + 1 < text.Length && XmlConvert.IsXmlSurrogatePair(text[k + 1], text[k]))
                {
                    k++;
                    continue;
                }

                if (char.IsControl(text[k]) || !XmlConvert.IsXmlChar(text[k])) return false;
            }

            return true;
        }

        private static string Url(string raw)
        {
            var url = raw.StartsWith("<", StringComparison.Ordinal) && raw.EndsWith(">", StringComparison.Ordinal)
                ? raw.Substring(1, raw.Length - 2)
                : raw;
            // 地址里同样认反斜杠转义和实体（从网页复制来的地址常带 &amp;）。
            url = Regex.Replace(url, @"\\([!-/:-@\[-`{-~])", "$1");
            var decoded = WebUtility.HtmlDecode(url);
            return IsPlainText(decoded) ? decoded : url;
        }

        private static bool Allowed(string url) =>
            Uri.TryCreate(url, UriKind.Absolute, out var uri) && Schemes.Contains(uri.Scheme, StringComparer.OrdinalIgnoreCase);

        /// <summary>裸网址结尾的标点多半是句子的：去掉，右括号只在网址里有对应的左括号时保留。</summary>
        private static int TrimUrl(string url)
        {
            var length = url.Length;
            while (length > 0)
            {
                var c = url[length - 1];
                if (".,:;!?*_~'\"".IndexOf(c) >= 0)
                {
                    length--;
                    continue;
                }

                if (c == ')' && url.Take(length).Count(x => x == '(') < url.Take(length).Count(x => x == ')'))
                {
                    length--;
                    continue;
                }

                break;
            }

            return length;
        }
    }
}
