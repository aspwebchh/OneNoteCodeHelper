using System.Text.RegularExpressions;

namespace OneNoteCodeHelper.Highlighting.Languages
{
    /// <summary>
    /// Java、C#、C++、JavaScript 共有的自动识别特征。四种语言都调这里，这部分分数完全相同、互相抵消，
    /// 胜负只取决于各自独有的写法；同时这几条又让 C 系整体跟 Python、Lua 这类语言拉开距离。
    ///
    /// 放在一处是为了加一条四种就都加上：以前 Java、C# 各写一份，C++ 少几条，JS 一条没有，
    /// JS 的 class 写法就这样输给了 Java/C#。
    /// </summary>
    internal static class CFamilyFeatures
    {
        /// <summary>
        /// 基本类型做泛型参数：C# 的 List&lt;string&gt;、C++ 的 vector&lt;int&gt;。Java 不能这么写，
        /// 所以不放进 <see cref="Score"/>，由 C# 和 C++ 各自以同样的权重计入、互相抵消。
        /// </summary>
        internal const string PrimitiveTypeArgument =
            @"<(string|int|bool|double|long|object|byte|char|float|decimal)(\[\])?[,>]";

        internal static int Score(DetectionSample sample)
        {
            var code = sample.Code;
            var score = 0;
            score += 3 * Count(code, @"\b(public|private|protected)\s+");
            score += 3 * Count(code, @"\b(class|interface|enum|struct|record)\s+[A-Z]\w*");
            score += 2 * Count(code, @"\bnew\s+[A-Z]\w*\s*\(");
            score += Count(code, @";\s*$");

            // 反证：# 加空白开头的整行注释是 Python、Bash、YAML、PowerShell 的写法（预处理指令 # 后面不带空白）
            score -= 3 * Count(code, @"^[^\S\r\n]*#([^\S\r\n]|$)");

            // 反证：Go、Rust、Kotlin 的函数声明。这几种语言不支持，宁可认不出也不要硬套成 C 系
            score -= 4 * Count(code, @"^[^\S\r\n]*(func|fn|fun)\s+[\w(]");
            return score;
        }

        private static int Count(string source, string pattern)
        {
            return LikelihoodPatterns.Count(source, pattern, RegexOptions.Multiline);
        }
    }
}
