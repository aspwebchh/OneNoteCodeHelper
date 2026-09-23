using System.Collections.Generic;

namespace OneNoteCodeHelper.Highlighting
{
    /// <summary>
    /// 一种语言的词法着色能力。加新语言 = 实现本接口 + 在 LanguageRegistry 里注册一行。
    /// </summary>
    internal interface ILanguage
    {
        /// <summary>稳定标识，用于配置持久化，如 "java"、"lua"。</summary>
        string Id { get; }

        /// <summary>界面上显示的名字。</summary>
        string DisplayName { get; }

        /// <summary>
        /// 把源码切成 token。实现必须保证返回的 token 按起点升序、互不重叠，
        /// 且合起来正好覆盖 [0, source.Length)。
        /// </summary>
        IEnumerable<Token> Tokenize(string source);

        /// <summary>
        /// 给出「这段代码像不像本语言」的打分，供自动识别用。分值只在同一批候选间比较大小。
        /// </summary>
        int ScoreLikelihood(string source);
    }
}
