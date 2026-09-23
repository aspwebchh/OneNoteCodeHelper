namespace OneNoteCodeHelper.Highlighting
{
    /// <summary>
    /// 源码中的一段连续字符及其分类。只存区间不存字符串，避免逐 token 切片。
    /// 词法分析器保证输出的 token 完整且不重叠地覆盖整个源码，渲染方可以直接顺序拼接。
    /// </summary>
    internal readonly struct Token
    {
        internal Token(TokenKind kind, int start, int length)
        {
            Kind = kind;
            Start = start;
            Length = length;
        }

        internal TokenKind Kind { get; }

        internal int Start { get; }

        internal int Length { get; }

        internal int End => Start + Length;
    }
}
