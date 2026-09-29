using System;
using System.Collections.Generic;
using System.Linq;

namespace OneNoteCodeHelper.Highlighting.Themes
{
    /// <summary>
    /// 内置配色方案。加新主题只需在 All 里加一行。
    /// 两套都接近 GitHub 的代码配色：不加粗、不用斜体，只靠颜色区分，配合无边框的代码框更清爽。
    /// </summary>
    internal static class CodeThemes
    {
        /// <summary>浅色，接近 GitHub 代码块：浅灰底色在白页上不用边框也分得清，打印也正常。</summary>
        internal static CodeTheme Light { get; } = new CodeTheme(
            id: "light",
            displayName: "浅色（GitHub）",
            background: "#F6F8FA",
            border: "#D0D7DE",
            defaultStyle: new TokenStyle("#1F2328"),
            styles: new Dictionary<TokenKind, TokenStyle>
            {
                [TokenKind.Keyword] = new TokenStyle("#CF222E"),
                [TokenKind.Literal] = new TokenStyle("#0550AE"),
                [TokenKind.Type] = new TokenStyle("#953800"),
                [TokenKind.Constant] = new TokenStyle("#0550AE"),
                [TokenKind.Function] = new TokenStyle("#8250DF"),
                [TokenKind.Builtin] = new TokenStyle("#0550AE"),
                [TokenKind.Variable] = new TokenStyle("#953800"),
                [TokenKind.Tag] = new TokenStyle("#116329"),
                [TokenKind.Attribute] = new TokenStyle("#0550AE"),
                [TokenKind.String] = new TokenStyle("#0A3069"),
                [TokenKind.Char] = new TokenStyle("#0A3069"),
                [TokenKind.Number] = new TokenStyle("#0550AE"),
                [TokenKind.Comment] = new TokenStyle("#6E7781"),
                [TokenKind.DocComment] = new TokenStyle("#6E7781"),
                [TokenKind.Annotation] = new TokenStyle("#116329"),
                [TokenKind.Operator] = new TokenStyle("#1F2328"),
                [TokenKind.Punctuation] = new TokenStyle("#1F2328")
            });

        /// <summary>深色，接近 GitHub Dimmed：蓝灰底色比纯黑柔和，放在白页上不刺眼。</summary>
        internal static CodeTheme Dark { get; } = new CodeTheme(
            id: "dark",
            displayName: "深色（GitHub Dimmed）",
            background: "#22272E",
            border: "#444C56",
            defaultStyle: new TokenStyle("#ADBAC7"),
            styles: new Dictionary<TokenKind, TokenStyle>
            {
                [TokenKind.Keyword] = new TokenStyle("#F47067"),
                [TokenKind.Literal] = new TokenStyle("#6CB6FF"),
                [TokenKind.Type] = new TokenStyle("#F69D50"),
                [TokenKind.Constant] = new TokenStyle("#6CB6FF"),
                [TokenKind.Function] = new TokenStyle("#DCBDFB"),
                [TokenKind.Builtin] = new TokenStyle("#6CB6FF"),
                [TokenKind.Variable] = new TokenStyle("#F69D50"),
                [TokenKind.Tag] = new TokenStyle("#8DDB8C"),
                [TokenKind.Attribute] = new TokenStyle("#6CB6FF"),
                [TokenKind.String] = new TokenStyle("#96D0FF"),
                [TokenKind.Char] = new TokenStyle("#96D0FF"),
                [TokenKind.Number] = new TokenStyle("#6CB6FF"),
                [TokenKind.Comment] = new TokenStyle("#768390"),
                [TokenKind.DocComment] = new TokenStyle("#768390"),
                [TokenKind.Annotation] = new TokenStyle("#8DDB8C"),
                [TokenKind.Operator] = new TokenStyle("#ADBAC7"),
                [TokenKind.Punctuation] = new TokenStyle("#ADBAC7")
            });

        internal static IReadOnlyList<CodeTheme> All { get; } = new[] { Light, Dark };

        internal static CodeTheme Find(string id)
        {
            return All.FirstOrDefault(t => string.Equals(t.Id, id, StringComparison.OrdinalIgnoreCase)) ?? Light;
        }
    }
}
