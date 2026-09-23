namespace OneNoteCodeHelper.Highlighting
{
    /// <summary>词法单元的分类。渲染时按这个枚举到主题里查颜色。</summary>
    internal enum TokenKind
    {
        /// <summary>空白、以及未归入其它类别的普通文本。</summary>
        Plain,

        /// <summary>语言关键字。</summary>
        Keyword,

        /// <summary>字面量关键字：Java 的 true/false/null、Lua 的 true/false/nil。</summary>
        Literal,

        /// <summary>类型名（Java 里按大写开头的标识符启发式判定）。</summary>
        Type,

        /// <summary>常量（全大写标识符，如 MAX_SIZE）。</summary>
        Constant,

        /// <summary>函数/方法名（标识符后紧跟左括号）。</summary>
        Function,

        /// <summary>语言内置库与内置函数（主要给 Lua 用）。</summary>
        Builtin,

        /// <summary>字符串字面量，含 Lua 长字符串与 Java 文本块。</summary>
        String,

        /// <summary>Java 的字符字面量。</summary>
        Char,

        /// <summary>数字字面量。</summary>
        Number,

        /// <summary>普通注释。</summary>
        Comment,

        /// <summary>Java 的 /** */ 文档注释。</summary>
        DocComment,

        /// <summary>Java 注解，如 @Override。</summary>
        Annotation,

        /// <summary>运算符。</summary>
        Operator,

        /// <summary>标点：括号、分号、逗号等。</summary>
        Punctuation
    }
}
