using System.Collections.Generic;

namespace OneNoteCodeHelper.Highlighting.Themes
{
    /// <summary>一类 token 的呈现方式。颜色是 #RRGGBB 形式，OneNote 和 WPF 都能直接吃。</summary>
    internal sealed class TokenStyle
    {
        internal TokenStyle(string color, bool bold = false, bool italic = false)
        {
            Color = color;
            Bold = bold;
            Italic = italic;
        }

        internal string Color { get; }

        internal bool Bold { get; }

        internal bool Italic { get; }
    }

    /// <summary>
    /// 一套配色。只管颜色，字体与字号属于用户设置（见 <see cref="Services.AddInSettings"/>），
    /// 因为换主题不应该把用户调好的字号也一起换掉。
    /// </summary>
    internal sealed class CodeTheme
    {
        private readonly IReadOnlyDictionary<TokenKind, TokenStyle> _styles;

        internal CodeTheme(
            string id,
            string displayName,
            string background,
            string border,
            TokenStyle defaultStyle,
            IReadOnlyDictionary<TokenKind, TokenStyle> styles)
        {
            Id = id;
            DisplayName = displayName;
            Background = background;
            Border = border;
            DefaultStyle = defaultStyle;
            _styles = styles;
        }

        internal string Id { get; }

        public string DisplayName { get; }

        /// <summary>代码框底色，填到 one:Cell 的 shadingColor。</summary>
        internal string Background { get; }

        /// <summary>预览用的边框色。OneNote 的表格边框颜色不可控，这个值只作用于 WPF 预览。</summary>
        internal string Border { get; }

        /// <summary>未单独配色的 token 用这个。</summary>
        internal TokenStyle DefaultStyle { get; }

        /// <summary>
        /// 默认样式是不是就等于 OneNote 自己的默认文字样式（纯黑、不加粗不斜体）。
        /// 成立时，默认色的文字可以完全不包 span，输出能小一大截。
        /// 深色主题不满足这个条件——那时必须显式上色，否则会变成深底黑字看不见。
        /// </summary>
        internal bool CanOmitDefaultSpans =>
            !DefaultStyle.Bold
            && !DefaultStyle.Italic
            && string.Equals(DefaultStyle.Color, "#000000", System.StringComparison.OrdinalIgnoreCase);

        /// <summary>某个类别是不是可以省掉 span。</summary>
        internal bool IsOmittable(TokenKind kind)
        {
            if (!CanOmitDefaultSpans)
            {
                return false;
            }

            var style = StyleFor(kind);
            return !style.Bold
                   && !style.Italic
                   && string.Equals(style.Color, DefaultStyle.Color, System.StringComparison.OrdinalIgnoreCase);
        }

        internal TokenStyle StyleFor(TokenKind kind)
        {
            return _styles.TryGetValue(kind, out var style) ? style : DefaultStyle;
        }
    }
}
