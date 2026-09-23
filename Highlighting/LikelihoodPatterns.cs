using System;
using System.Collections.Concurrent;
using System.Text.RegularExpressions;

namespace OneNoteCodeHelper.Highlighting
{
    /// <summary>
    /// 自动识别打分用的正则匹配。各语言的 ScoreLikelihood 都经这里跑正则。
    ///
    /// 各语言加起来有上百条模式，远超 Regex 静态方法的缓存上限（默认 15 条）。直接调 Regex.Matches
    /// 的话，每识别一次都要把全部模式重新解析一遍，所以这里自己缓存 Regex 实例。
    ///
    /// 写模式时注意：多行模式下不要用 ^\s*，\s 会匹配换行，每个行首都会把后面连续的空行重扫一遍，
    /// 几千个空行就要跑好几分钟（实测 2 万个空行 253 秒）。行首缩进用 ^[^\S\r\n]*；
    /// 同理，[^;{}]* 这类取反字符类也要把 \r\n 排除掉，否则会一路扫到文件末尾。
    /// 匹配超时是最后一道保险：自动识别只是锦上添花，哪条模式碰上病态输入，宁可按 0 分算也不能卡住。
    /// </summary>
    internal static class LikelihoodPatterns
    {
        private static readonly TimeSpan MatchTimeout = TimeSpan.FromMilliseconds(100);

        private static readonly ConcurrentDictionary<(string Pattern, RegexOptions Options), Regex> Cache =
            new ConcurrentDictionary<(string, RegexOptions), Regex>();

        /// <summary>模式在 source 里出现的次数；超时按 0 算。</summary>
        internal static int Count(string source, string pattern, RegexOptions options)
        {
            try
            {
                return Get(pattern, options).Matches(source).Count;
            }
            catch (RegexMatchTimeoutException)
            {
                return 0;
            }
        }

        /// <summary>模式是否出现在 source 里；超时按没出现算。</summary>
        internal static bool IsMatch(string source, string pattern, RegexOptions options = RegexOptions.None)
        {
            try
            {
                return Get(pattern, options).IsMatch(source);
            }
            catch (RegexMatchTimeoutException)
            {
                return false;
            }
        }

        private static Regex Get(string pattern, RegexOptions options)
        {
            return Cache.GetOrAdd((pattern, options), key => new Regex(key.Pattern, key.Options, MatchTimeout));
        }
    }
}
