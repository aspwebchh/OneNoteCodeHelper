using System.Collections.Generic;

namespace OneNoteCodeHelper.Highlighting.Languages
{
    /// <summary>
    /// 纯文本：不着色，只要等宽字体和带底色的代码框。用来放日志、命令输出、配置片段这类东西。
    /// 自动识别永远不会选它，只能手动选。
    /// </summary>
    internal sealed class TextLanguage : ILanguage
    {
        public string Id => "text";

        public string DisplayName => "纯文本";

        public IEnumerable<Token> Tokenize(string source)
        {
            return new LexerCursor(source).Finish();
        }

        public int ScoreLikelihood(string source) => 0;
    }
}
