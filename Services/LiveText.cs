using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OneNoteCodeHelper.Services
{
    /// <summary>
    /// 等模型时在 Agent 窗口里显示的实时文字：思考（或回复）原文的最后几行。
    /// 只在窗口里显示，不写日志。
    /// </summary>
    internal static class LiveText
    {
        /// <summary>摘录最多多少字。窗口里的摘录框大约三行，多出来的从顶上裁掉。</summary>
        internal const int ExcerptChars = 160;

        /// <summary>截断后在开头这么多字以内找一个断点，免得从半个词开始。</summary>
        private const int BoundarySearch = 20;

        private const string Boundaries = "，。；：、！？,.;:!?)）】」』";

        /// <summary>取末尾一段做摘录。只拷贝末尾需要的部分，思考再长也不会整段 ToString。</summary>
        internal static string Excerpt(StringBuilder text, int maxChars = ExcerptChars)
        {
            if (text == null || text.Length == 0)
            {
                return null;
            }

            // 多拿一些，压掉空行后还够 maxChars。
            var take = Math.Min(text.Length, maxChars * 2);
            return Excerpt(text.ToString(text.Length - take, take), maxChars, take < text.Length);
        }

        internal static string Excerpt(string text, int maxChars = ExcerptChars)
        {
            return Excerpt(text, maxChars, false);
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

        private static string Excerpt(string text, int maxChars, bool truncated)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            // 空行（含只有空白的行）全部去掉，行尾空白也去掉：摘录框很小，一行都不能浪费。
            var lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n')
                .Select(l => l.TrimEnd())
                .Where(l => l.Length > 0);
            var result = string.Join("\n", lines).Trim();

            if (result.Length > maxChars)
            {
                result = result.Substring(result.Length - maxChars);
                truncated = true;
            }

            if (truncated)
            {
                var cut = FindBoundary(result);
                if (cut > 0)
                {
                    result = result.Substring(cut);
                }
                else if (result.Length > 0 && char.IsLowSurrogate(result[0]))
                {
                    // 截在了一个表情字符中间。
                    result = result.Substring(1);
                }

                result = result.TrimStart();
                if (result.Length == 0)
                {
                    return null;
                }

                result = "…" + result;
            }

            return result.Length == 0 ? null : result;
        }

        /// <summary>开头一小段里第一个空白或标点之后的位置，找不到返回 0。</summary>
        private static int FindBoundary(string text)
        {
            var limit = Math.Min(BoundarySearch, text.Length - 1);
            for (var i = 0; i < limit; i++)
            {
                if (char.IsWhiteSpace(text[i]) || Boundaries.IndexOf(text[i]) >= 0)
                {
                    return i + 1;
                }
            }

            return 0;
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
