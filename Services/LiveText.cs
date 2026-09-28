using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OneNoteCodeHelper.Services
{
    /// <summary>
    /// 等模型时在 Agent 窗口里显示的实时文字：从思考（或回复）原文里挑出最新的几句核心内容。
    /// 直接截原文末尾总是半句，夹着 Markdown 符号和代码片段，每次刷新还都在跳；
    /// 这里只要说完的整句，去掉标记、代码和语气词，说完新的一句才换。只在窗口里显示，不写日志。
    /// </summary>
    internal static class LiveText
    {
        /// <summary>只看原文末尾这么多字，思考再长也不会整段 ToString。<see cref="KeepTail"/> 留下的比这多。</summary>
        private const int GistWindow = 1500;

        /// <summary>摘录框三行，一行约 30 个汉字。宽度按 <see cref="Width"/> 估：全角算 2，其余算 1。</summary>
        private const int GistLines = 3;

        private const int LineUnits = 58;

        /// <summary>有效字（见 <see cref="Weight"/>）不到这么多的句子多半是「嗯」「好的」「Got it」，不显示。</summary>
        private const int MinWeight = 8;

        /// <summary>句末标点。英文句号另算，见 <see cref="EndsSentence"/>。</summary>
        private const string Stops = "。！？!?…";

        /// <summary>句末标点后面连带的重复标点、引号和括号，归到这一句里。</summary>
        private const string Closers = "。！？!?….”’\"'」』）)】";

        /// <summary>这些符号占比太高的句子是代码、JSON 或表格行。</summary>
        private const string CodeSymbols = "{}[]<>\"=;|\\$";

        /// <summary>行首的 Markdown 标记：标题、引用、列表符号、待办框和编号。</summary>
        private static readonly Regex LineMarker = new Regex(
            @"^\s*(?:#{1,6}\s+|>\s*|[-*+•·]\s+|\[[ xX]\]\s+|\d{1,2}(?:[.)]\s+|、\s*)|[（(]\d{1,2}[)）]\s*)+");

        /// <summary>行内的加粗、删除线和代码标记。</summary>
        private static readonly Regex InlineMarker = new Regex(@"\*\*|__|~~|`+");

        private static readonly Regex Spaces = new Regex(@"\s+");

        /// <summary>句首的语气词。后面必须跟标点或空白，「好几段」的「好」不会被切掉。</summary>
        private static readonly Regex Filler = new Regex(
            @"^(?:(?:嗯+|哦|啊|好的?|OK|Okay|Hmm+|Wait|Alright|Well|So)(?:[，,、。.!！…]+\s*|\s+))+",
            RegexOptions.IgnoreCase);

        /// <summary>整句只是「让我想想」「我再检查一下」这类空话。</summary>
        private static readonly Regex Hollow = new Regex(
            @"^(?:(?:让我|我)?再?(?:想想|想一想|想一下|看看|看一下|检查一下|确认一下|思考一下)|let me (?:think|see|check|verify)(?: (?:again|this|that|it))?)[。.!！…]*$",
            RegexOptions.IgnoreCase);

        /// <summary>
        /// 最新的几句核心内容：句与句换行，最新的在最下面，合起来不超过摘录框的三行。
        /// 没有可显示的句子时返回 null，调用方留着上一次的，摘录框不闪。
        /// finished 表示这段文字已经收完（模型开始回复或调用工具了），最后一句没有句末标点也算说完。
        /// </summary>
        internal static string Gist(StringBuilder text, bool finished)
        {
            if (text == null || text.Length == 0)
            {
                return null;
            }

            var take = Math.Min(text.Length, GistWindow);
            var sentences = Sentences(text.ToString(text.Length - take, take), take < text.Length, finished)
                .Select(Clean)
                .Where(s => s != null)
                .ToList();
            if (sentences.Count == 0)
            {
                return null;
            }

            // 从最新一句往前加，放得下就继续。
            var picked = new List<string>();
            var lines = 0;
            for (var i = sentences.Count - 1; i >= 0; i--)
            {
                var need = (Width(sentences[i]) + LineUnits - 1) / LineUnits;
                if (lines + need > GistLines)
                {
                    break;
                }

                picked.Insert(0, sentences[i]);
                lines += need;
            }

            // 最新一句自己就超过三行：从头截，至少看得出这句在说什么。
            return picked.Count > 0
                ? string.Join("\n", picked)
                : Shorten(sentences[sentences.Count - 1], GistLines * LineUnits - 2);
        }

        /// <summary>
        /// 只保留末尾一段，长度超过 2 × keep 时把前面删掉。摘录只看末尾，没必要把整段思考都攒在内存里。
        /// </summary>
        internal static void KeepTail(StringBuilder text, int keep = 2000)
        {
            if (text.Length > keep * 2)
            {
                text.Remove(0, text.Length - keep);
            }
        }

        /// <summary>
        /// 按行、再按句末标点切句，行首的 Markdown 标记去掉，``` 围起来的代码整段跳过。
        /// truncated 时第一句是从中间截断的，丢掉；没收完时最后一个句末之后的半句也丢掉。
        /// </summary>
        private static IEnumerable<string> Sentences(string text, bool truncated, bool finished)
        {
            var lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            var code = false;
            for (var i = 0; i < lines.Length; i++)
            {
                if (lines[i].TrimStart().StartsWith("```", StringComparison.Ordinal))
                {
                    code = !code;
                    continue;
                }

                if (code)
                {
                    continue;
                }

                var open = i == lines.Length - 1 && !finished;
                var dropFirst = truncated && i == 0;
                var line = LineMarker.Replace(lines[i], "");
                var start = 0;
                for (var j = 0; j < line.Length; j++)
                {
                    if (!EndsSentence(line, j, open))
                    {
                        continue;
                    }

                    var end = j + 1;
                    while (end < line.Length && Closers.IndexOf(line[end]) >= 0)
                    {
                        end++;
                    }

                    if (dropFirst)
                    {
                        dropFirst = false;
                    }
                    else
                    {
                        yield return line.Substring(start, end - start);
                    }

                    start = end;
                    j = end - 1;
                }

                // 行尾也是句末，除非这一行还没收完。
                if (!open && !dropFirst && start < line.Length)
                {
                    yield return line.Substring(start);
                }
            }
        }

        /// <summary>
        /// line[j] 是不是句末。英文句号后面跟空白才算，11.5、p1.p2 不断开；
        /// 没收完的行里句号在最末尾时还不知道后面是什么，先不算。
        /// </summary>
        private static bool EndsSentence(string line, int j, bool open)
        {
            if (line[j] == '.')
            {
                return j + 1 < line.Length ? char.IsWhiteSpace(line[j + 1]) : !open;
            }

            return Stops.IndexOf(line[j]) >= 0;
        }

        /// <summary>去掉行内标记和句首语气词；太短、空话、像代码的句子返回 null。</summary>
        private static string Clean(string sentence)
        {
            var text = Spaces.Replace(InlineMarker.Replace(sentence, ""), " ").Trim();
            text = Filler.Replace(text, "").Trim();
            return Weight(text) < MinWeight || Hollow.IsMatch(text) || LooksLikeCode(text) ? null : text;
        }

        private static bool LooksLikeCode(string text)
        {
            if (text.Contains("{\"") || text.Contains("\":") || text.Contains("</") || text.Contains("/>") || text.Contains("=>"))
            {
                return true;
            }

            return text.Count(c => CodeSymbols.IndexOf(c) >= 0) * 5 > text.Length;
        }

        /// <summary>有效字数：字母、数字、汉字，全角的算 2。标点和空白不算。</summary>
        private static int Weight(string text)
        {
            return text.Sum(c => char.IsLetterOrDigit(c) ? (IsWide(c) ? 2 : 1) : 0);
        }

        /// <summary>估算显示宽度：汉字、全角标点和表情算 2，其余算 1。低位代理不算，和高位代理合起来是一个字。</summary>
        private static int Width(string text)
        {
            return text.Sum(c => char.IsLowSurrogate(c) ? 0 : IsWide(c) ? 2 : 1);
        }

        private static bool IsWide(char c)
        {
            return (c >= '⺀' && c <= '꓏') || (c >= '가' && c <= '힣') || (c >= '豈' && c <= '﫿') ||
                   (c >= '︰' && c <= '﹏') || (c >= '＀' && c <= '｠') || (c >= '￠' && c <= '￦') ||
                   char.IsHighSurrogate(c);
        }

        /// <summary>从头保留不超过 maxUnits 的宽度，末尾加省略号。截断点落在高位代理上，整个字一起去掉。</summary>
        private static string Shorten(string text, int maxUnits)
        {
            var units = 0;
            for (var i = 0; i < text.Length; i++)
            {
                units += char.IsLowSurrogate(text[i]) ? 0 : IsWide(text[i]) ? 2 : 1;
                if (units > maxUnits)
                {
                    return text.Substring(0, i).TrimEnd() + "…";
                }
            }

            return text;
        }
    }

    /// <summary>
    /// 边收边看文字功能的返回（约定的 JSON：{"paragraphs":[{"id":..,"text":..,"changes":[..]}]}），
    /// 不等整段收完就能说出「已经返回了几段修改、最新一条改动说明是什么」。
    ///
    /// 只做进度显示用的粗略扫描，片段可以在任何地方断开；真正写回用的还是收完之后
    /// <see cref="AiOptimizer.ParseReply"/> 的结果。JSON 外面的杂字符（比如 ``` 代码块标记）忽略。
    /// </summary>
    internal sealed class ReplyPeek
    {
        /// <summary>还没关上的 { 或 [，以及它是哪个键的值（数组里的元素、最外层为 null）。</summary>
        private readonly Stack<(bool IsArray, string Key)> _containers = new Stack<(bool, string)>();

        private readonly StringBuilder _string = new StringBuilder();

        private bool _inString;

        private bool _escape;

        /// <summary>\uXXXX 还差几位十六进制数，0 表示不在 \u 里。</summary>
        private int _hexLeft;

        private int _hexValue;

        /// <summary>对象里最近一个字符串，碰到冒号就成了键。</summary>
        private string _lastString;

        /// <summary>对象里当前值对应的键：冒号之后到逗号之前有效。</summary>
        private string _valueKey;

        /// <summary>已经完整收到的段落对象个数（直接位于数组里的对象闭合一次算一个）。</summary>
        internal int Paragraphs { get; private set; }

        /// <summary>最新一条完整收到的改动说明（changes 里的字符串），没有时为 null。</summary>
        internal string LastChange { get; private set; }

        /// <summary>收到的改动说明条数，调用方用它判断 <see cref="LastChange"/> 有没有更新。</summary>
        internal int Changes { get; private set; }

        internal void Feed(string fragment)
        {
            if (string.IsNullOrEmpty(fragment))
            {
                return;
            }

            foreach (var ch in fragment)
            {
                if (_inString)
                {
                    ReadStringChar(ch);
                    continue;
                }

                switch (ch)
                {
                    case '"':
                        _inString = true;
                        _string.Clear();
                        break;
                    case '{':
                        _containers.Push((false, CurrentValueKey()));
                        _valueKey = null;
                        break;
                    case '[':
                        _containers.Push((true, CurrentValueKey()));
                        break;
                    case '}':
                        if (_containers.Count > 0 && !_containers.Peek().IsArray)
                        {
                            _containers.Pop();
                            if (_containers.Count > 0 && _containers.Peek().IsArray)
                            {
                                Paragraphs++;
                            }
                        }

                        break;
                    case ']':
                        if (_containers.Count > 0 && _containers.Peek().IsArray)
                        {
                            _containers.Pop();
                        }

                        break;
                    case ':':
                        _valueKey = _lastString;
                        break;
                    case ',':
                        if (_containers.Count > 0 && !_containers.Peek().IsArray)
                        {
                            _valueKey = null;
                        }

                        break;
                }
            }
        }

        /// <summary>新开的容器是哪个键的值：在对象里就是冒号前的键，在数组里或最外层为 null。</summary>
        private string CurrentValueKey()
        {
            return _containers.Count > 0 && !_containers.Peek().IsArray ? _valueKey : null;
        }

        private void ReadStringChar(char ch)
        {
            if (_hexLeft > 0)
            {
                _hexValue = _hexValue * 16 + (Uri.IsHexDigit(ch) ? Uri.FromHex(ch) : 0);
                if (--_hexLeft == 0)
                {
                    _string.Append((char)_hexValue);
                }

                return;
            }

            if (_escape)
            {
                _escape = false;
                switch (ch)
                {
                    case 'n': _string.Append('\n'); break;
                    case 't': _string.Append('\t'); break;
                    case 'r': _string.Append('\r'); break;
                    case 'b': _string.Append('\b'); break;
                    case 'f': _string.Append('\f'); break;
                    case 'u':
                        _hexLeft = 4;
                        _hexValue = 0;
                        break;
                    default: _string.Append(ch); break;
                }

                return;
            }

            if (ch == '\\')
            {
                _escape = true;
            }
            else if (ch == '"')
            {
                _inString = false;
                EndString(_string.ToString());
            }
            else
            {
                _string.Append(ch);
            }
        }

        private void EndString(string value)
        {
            if (_containers.Count == 0)
            {
                return;
            }

            var top = _containers.Peek();
            if (!top.IsArray && _valueKey == null)
            {
                // 对象里冒号之前的字符串是键。
                _lastString = value;
                return;
            }

            // changes 一般是字符串数组，模型偶尔直接写成一个字符串。
            var isChange = top.IsArray ? top.Key == "changes" : _valueKey == "changes";
            if (isChange && !string.IsNullOrWhiteSpace(value))
            {
                LastChange = value.Trim();
                Changes++;
            }
        }
    }
}
