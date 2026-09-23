using System;
using System.Collections.Generic;

namespace OneNoteCodeHelper.Highlighting
{
    /// <summary>
    /// 自动识别用的样本。每次识别只构造一次，所有语言共用。
    ///
    /// <see cref="Raw"/> 是原文（只取开头一段）；<see cref="Code"/> 是把注释文字和字符串内容换成空格后的文本，
    /// 和 Raw 等长、换行原样保留，所以 ^、$ 和行结构都跟原文对得上。关键字、结构类的特征应该在 Code 上匹配，
    /// 不然注释里的一句 "public class"、字符串里的一条 SQL 都会给别的语言加分。
    /// 要看注释标记本身或字符串内容的特征（C# 的 ///、Bash 的 "$1"、JSON 的键）才用 Raw。
    ///
    /// 去噪时还不知道是什么语言，所以只认各语言通用、不容易误伤的写法：整行注释、块注释、同一行内配对的引号，
    /// 以及 Python 的三引号。注释标记和引号本身保留，只抹掉里面的内容。
    /// </summary>
    internal sealed class DetectionSample
    {
        /// <summary>最多看这么多字符，约四五百行代码。</summary>
        private const int SampleLength = 16 * 1024;

        /// <summary>行首 # 后面跟空白时是注释，但 C 预处理指令也可以写成 "# include"，这些不能当注释抹掉。</summary>
        private static readonly HashSet<string> PreprocessorDirectives = new HashSet<string>(StringComparer.Ordinal)
        {
            "include", "define", "undef", "if", "ifdef", "ifndef", "elif", "else", "endif", "pragma", "error", "line"
        };

        internal DetectionSample(string source)
        {
            var truncated = source.Length > SampleLength;
            Raw = truncated ? TakeHead(source) : source;
            Code = MaskCommentsAndStrings(Raw);
            IsMarkupDocument = LooksLikeMarkupDocument(Raw, truncated);
        }

        /// <summary>原文开头的一段。</summary>
        internal string Raw { get; }

        /// <summary>注释文字、字符串内容换成空格后的 Raw，长度和行结构不变。</summary>
        internal string Code { get; }

        /// <summary>
        /// 整段是一份标记文档：以标签开头、以 &gt; 结尾（截断过的只看开头）。
        /// 这种情况下里面的 &lt;script&gt;、&lt;style&gt; 再像别的语言，整段也还是 HTML/XML。
        /// </summary>
        internal bool IsMarkupDocument { get; }

        /// <summary>
        /// 只拿开头一段来打分，在行尾截断。几百行足够看出是什么语言；整份几千行的文件全量跑上百条正则
        /// 要大半秒，而「插入代码」窗口每改一次都要重新识别。
        /// </summary>
        private static string TakeHead(string source)
        {
            var cut = source.LastIndexOf('\n', SampleLength - 1);
            return source.Substring(0, cut > 0 ? cut + 1 : SampleLength);
        }

        private static bool LooksLikeMarkupDocument(string text, bool truncated)
        {
            var first = 0;
            while (first < text.Length && char.IsWhiteSpace(text[first]))
            {
                first++;
            }

            if (first + 1 >= text.Length || text[first] != '<')
            {
                return false;
            }

            var next = text[first + 1];
            if (!char.IsLetter(next) && next != '!' && next != '?')
            {
                return false;
            }

            if (truncated)
            {
                return true;
            }

            var last = text.Length - 1;
            while (last > first && char.IsWhiteSpace(text[last]))
            {
                last--;
            }

            return text[last] == '>';
        }

        private static string MaskCommentsAndStrings(string text)
        {
            var chars = text.ToCharArray();
            var atLineStart = true;
            var i = 0;

            while (i < text.Length)
            {
                var ch = text[i];

                if (ch == '\n' || ch == '\r')
                {
                    atLineStart = true;
                    i++;
                    continue;
                }

                if (atLineStart)
                {
                    if (ch == ' ' || ch == '\t')
                    {
                        i++;
                        continue;
                    }

                    atLineStart = false;
                    var marker = LineCommentMarkerLength(text, i);
                    if (marker > 0)
                    {
                        i = Blank(chars, i + marker, LineEnd(text, i));
                        continue;
                    }
                }

                var skipTo = TryMaskBlock(text, chars, i);
                if (skipTo > i)
                {
                    i = skipTo;
                    continue;
                }

                i++;
            }

            return new string(chars);
        }

        /// <summary>
        /// 行首的整行注释标记长度，不是注释返回 0。认 //（含 ///）、# 加空白、-- 加空白：
        /// #include、#region、#!、CSS 的 #id 和 --var: 后面都不是空白，不会被当成注释。
        /// 行尾注释不处理：字符串里的 URL、CSS 颜色都可能被误认成注释开头。
        /// </summary>
        private static int LineCommentMarkerLength(string text, int i)
        {
            if (Matches(text, i, "//"))
            {
                return RunLength(text, i, '/');
            }

            if (text[i] == '#')
            {
                var run = RunLength(text, i, '#');
                return IsBlankOrLineEnd(text, i + run) && !IsPreprocessorDirective(text, i + run) ? run : 0;
            }

            if (Matches(text, i, "--") && IsBlankOrLineEnd(text, i + 2))
            {
                return 2;
            }

            return 0;
        }

        private static bool IsPreprocessorDirective(string text, int i)
        {
            while (i < text.Length && (text[i] == ' ' || text[i] == '\t'))
            {
                i++;
            }

            var start = i;
            while (i < text.Length && char.IsLetter(text[i]))
            {
                i++;
            }

            return i > start && PreprocessorDirectives.Contains(text.Substring(start, i - start));
        }

        /// <summary>
        /// 从 i 开始的块注释或字符串：抹掉内容并返回其后的位置；不是或者找不到闭合就返回 i。
        /// </summary>
        private static int TryMaskBlock(string text, char[] chars, int i)
        {
            var ch = text[i];

            // /* 前面必须是空白或标点，免得 ls /tmp/* 这种路径被当成注释开头
            if (ch == '/' && Matches(text, i, "/*") && (i == 0 || !IsPathChar(text[i - 1])))
            {
                return MaskUntil(text, chars, i, 2, "*/");
            }

            if (ch == '<' && Matches(text, i, "<!--"))
            {
                return MaskUntil(text, chars, i, 4, "-->");
            }

            // PowerShell 的块注释
            if (ch == '<' && Matches(text, i, "<#"))
            {
                return MaskUntil(text, chars, i, 2, "#>");
            }

            if ((ch == '"' || ch == '\'') && Matches(text, i, new string(ch, 3)))
            {
                return MaskUntil(text, chars, i, 3, new string(ch, 3));
            }

            if (ch == '"' || ch == '\'' || ch == '`')
            {
                var close = FindClosingQuoteOnLine(text, i + 1, ch);
                if (close > 0)
                {
                    Blank(chars, i + 1, close);
                    return close + 1;
                }
            }

            return i;
        }

        private static int MaskUntil(string text, char[] chars, int start, int openLength, string close)
        {
            var end = text.IndexOf(close, start + openLength, StringComparison.Ordinal);
            if (end < 0)
            {
                return start;
            }

            Blank(chars, start + openLength, end);
            return end + close.Length;
        }

        /// <summary>同一行内配对的引号位置，跳过 \ 转义；找不到返回 -1。</summary>
        private static int FindClosingQuoteOnLine(string text, int from, char quote)
        {
            for (var j = from; j < text.Length; j++)
            {
                var c = text[j];
                if (c == '\n' || c == '\r')
                {
                    return -1;
                }

                if (c == '\\')
                {
                    j++;
                    continue;
                }

                if (c == quote)
                {
                    return j;
                }
            }

            return -1;
        }

        /// <summary>把 [start, end) 里除换行以外的字符换成空格，返回 end。</summary>
        private static int Blank(char[] chars, int start, int end)
        {
            for (var j = start; j < end; j++)
            {
                if (chars[j] != '\n' && chars[j] != '\r')
                {
                    chars[j] = ' ';
                }
            }

            return end;
        }

        private static int LineEnd(string text, int i)
        {
            while (i < text.Length && text[i] != '\n' && text[i] != '\r')
            {
                i++;
            }

            return i;
        }

        private static int RunLength(string text, int i, char c)
        {
            var start = i;
            while (i < text.Length && text[i] == c)
            {
                i++;
            }

            return i - start;
        }

        private static bool Matches(string text, int i, string value)
        {
            return string.CompareOrdinal(text, i, value, 0, value.Length) == 0;
        }

        private static bool IsBlankOrLineEnd(string text, int i)
        {
            return i >= text.Length || char.IsWhiteSpace(text[i]);
        }

        private static bool IsPathChar(char c) => char.IsLetterOrDigit(c) || c == '/' || c == '.' || c == '_' || c == '-' || c == '~';
    }
}
