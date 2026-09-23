using System;
using System.Collections.Generic;
using System.Linq;

namespace OneNoteCodeHelper.Highlighting.Themes
{
    /// <summary>内置配色方案。加新主题只需在 All 里加一行。</summary>
    internal static class CodeThemes
    {
        /// <summary>浅色，接近 IntelliJ IDEA 默认配色，和 OneNote 白页最协调，打印也正常。</summary>
        internal static CodeTheme Light { get; } = new CodeTheme(
            id: "light",
            displayName: "浅色（类 IntelliJ）",
            background: "#F7F7F7",
            border: "#D8D8D8",
            defaultStyle: new TokenStyle("#000000"),
            styles: new Dictionary<TokenKind, TokenStyle>
            {
                [TokenKind.Keyword] = new TokenStyle("#7F0055", bold: true),
                [TokenKind.Literal] = new TokenStyle("#7F0055", bold: true),
                [TokenKind.Type] = new TokenStyle("#20999D"),
                [TokenKind.Constant] = new TokenStyle("#660E7A", italic: true),
                [TokenKind.Function] = new TokenStyle("#795E26"),
                [TokenKind.Builtin] = new TokenStyle("#7A3E9D"),
                [TokenKind.String] = new TokenStyle("#008000"),
                [TokenKind.Char] = new TokenStyle("#008000"),
                [TokenKind.Number] = new TokenStyle("#1750EB"),
                [TokenKind.Comment] = new TokenStyle("#808080", italic: true),
                [TokenKind.DocComment] = new TokenStyle("#3F5FBF", italic: true),
                [TokenKind.Annotation] = new TokenStyle("#808000"),
                [TokenKind.Operator] = new TokenStyle("#000000"),
                [TokenKind.Punctuation] = new TokenStyle("#000000")
            });

        /// <summary>深色，接近 VS Code Dark+，和 IDE 里看到的一致。</summary>
        internal static CodeTheme Dark { get; } = new CodeTheme(
            id: "dark",
            displayName: "深色（类 VS Code Dark+）",
            background: "#1E1E1E",
            border: "#3C3C3C",
            defaultStyle: new TokenStyle("#D4D4D4"),
            styles: new Dictionary<TokenKind, TokenStyle>
            {
                [TokenKind.Keyword] = new TokenStyle("#569CD6"),
                [TokenKind.Literal] = new TokenStyle("#569CD6"),
                [TokenKind.Type] = new TokenStyle("#4EC9B0"),
                [TokenKind.Constant] = new TokenStyle("#4FC1FF"),
                [TokenKind.Function] = new TokenStyle("#DCDCAA"),
                [TokenKind.Builtin] = new TokenStyle("#C586C0"),
                [TokenKind.String] = new TokenStyle("#CE9178"),
                [TokenKind.Char] = new TokenStyle("#CE9178"),
                [TokenKind.Number] = new TokenStyle("#B5CEA8"),
                [TokenKind.Comment] = new TokenStyle("#6A9955", italic: true),
                [TokenKind.DocComment] = new TokenStyle("#6A9955", italic: true),
                [TokenKind.Annotation] = new TokenStyle("#DCDCAA"),
                [TokenKind.Operator] = new TokenStyle("#D4D4D4"),
                [TokenKind.Punctuation] = new TokenStyle("#D4D4D4")
            });

        internal static IReadOnlyList<CodeTheme> All { get; } = new[] { Light, Dark };

        internal static CodeTheme Find(string id)
        {
            return All.FirstOrDefault(t => string.Equals(t.Id, id, StringComparison.OrdinalIgnoreCase)) ?? Light;
        }
    }
}
