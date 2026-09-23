namespace OneNoteCodeHelper.Highlighting
{
    /// <summary>词法单元的分类。渲染时按这个枚举到主题里查颜色。</summary>
    internal enum TokenKind
    {
        /// <summary>空白、以及未归入其它类别的普通文本。</summary>
        Plain,

        /// <summary>语言关键字。</summary>
        Keyword,

        /// <summary>字面量关键字：Java 的 true/false/null、Lua 的 nil、PowerShell 的 $true/$null、YAML 的 yes/no/~ 等。</summary>
        Literal,

        /// <summary>类型名（Java 里按大写开头的标识符启发式判定；PowerShell 的 [type]；CSS 的 .class；YAML 的 !tag）。</summary>
        Type,

        /// <summary>常量（全大写标识符，如 MAX_SIZE；XML/HTML 实体；CSS 的 #id）。</summary>
        Constant,

        /// <summary>函数/方法名（标识符后紧跟左括号；PowerShell 的 Verb-Noun 命令）。</summary>
        Function,

        /// <summary>语言内置库、内置函数与内置命令（Lua 标准库、Bash/Bat 内置命令）。</summary>
        Builtin,

        /// <summary>变量引用：$var、${var}、%VAR%、!VAR!、%%i、CSS 的 --custom-prop、YAML 的 &amp;anchor 与 *alias。</summary>
        Variable,

        /// <summary>XML/HTML 标签名，以及 CSS 的元素选择器。</summary>
        Tag,

        /// <summary>XML/HTML 属性名、CSS 属性名、JSON/YAML 的键、命令行参数与开关（-Path、--flag、/Q）。</summary>
        Attribute,

        /// <summary>字符串字面量，含 Lua 长字符串、Java 文本块、here-string 与 heredoc。</summary>
        String,

        /// <summary>Java 的字符字面量。</summary>
        Char,

        /// <summary>数字字面量。</summary>
        Number,

        /// <summary>普通注释。</summary>
        Comment,

        /// <summary>Java 的 /** */ 文档注释。</summary>
        DocComment,

        /// <summary>Java 注解（@Override）、Bat 标签（:label）、XML 声明（&lt;?xml ?&gt;）、CSS 伪类（:hover）。</summary>
        Annotation,

        /// <summary>运算符。</summary>
        Operator,

        /// <summary>标点：括号、分号、逗号等。</summary>
        Punctuation
    }
}
