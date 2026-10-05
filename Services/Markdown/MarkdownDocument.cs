using System.Collections.Generic;
using System.Linq;
using OneNoteCodeHelper.Highlighting;
using OneNoteCodeHelper.Services.Agent;

namespace OneNoteCodeHelper.Services.Markdown
{
    internal enum MarkdownBlockKind { Paragraph, Heading, Quote, ListItem, Code, Table, Blank }

    /// <summary>一段格式相同的行内文字。Break 表示段内换行（Shift+Enter），这时 Text 为空。</summary>
    internal sealed class MarkdownRun
    {
        internal string Text = string.Empty;
        internal bool Bold;
        internal bool Italic;
        internal bool Strike;
        internal bool Code;
        internal bool Break;
        internal string Href;

        internal bool SameFormat(MarkdownRun other) =>
            !Break && !other.Break && Bold == other.Bold && Italic == other.Italic && Strike == other.Strike &&
            Code == other.Code && Href == other.Href;
    }

    /// <summary>
    /// 一个块：OneNote 里的一段（标题、正文、引用、列表项、空行），或一个代码框、一张表格。
    /// Level 是缩进层级，写到 OneNote 时挂到上一层最近那段的 OEChildren 下面。
    /// </summary>
    internal sealed class MarkdownBlock
    {
        internal MarkdownBlockKind Kind;
        internal int Level;
        /// <summary>标题级别 1–6。</summary>
        internal int HeadingLevel;
        /// <summary>bullet、number，或 null（不是列表、或只是待办）。</summary>
        internal string ListKind;
        /// <summary>编号列表每组第一项的起始编号，其余项为 null。</summary>
        internal int? Number;
        /// <summary>待办：null 不是待办，false 未完成，true 已完成。</summary>
        internal bool? Todo;
        internal List<MarkdownRun> Inline = new List<MarkdownRun>();
        internal string Code;
        internal ILanguage Language;
        /// <summary>表格各行（第一行是标题行），每格一组行内文字。</summary>
        internal List<List<List<MarkdownRun>>> Rows;
        /// <summary>各列对齐：left、center、right，或 null（不指定）。</summary>
        internal List<string> Alignments;

        /// <summary>套用的段落预设。四到六级标题 OneNote 预设里没有对应，用正文加粗。</summary>
        internal string Preset
        {
            get
            {
                if (Kind == MarkdownBlockKind.Quote) return "quote";
                if (Kind != MarkdownBlockKind.Heading) return "body";
                switch (HeadingLevel)
                {
                    case 1: return "heading1";
                    case 2: return "heading2";
                    case 3: return ParagraphStyles.Heading3;
                    default: return "body";
                }
            }
        }

        /// <summary>整段加粗：四到六级标题。</summary>
        internal bool BoldText => Kind == MarkdownBlockKind.Heading && HeadingLevel >= 4;

        internal bool IsText => Kind == MarkdownBlockKind.Paragraph || Kind == MarkdownBlockKind.Heading ||
                                Kind == MarkdownBlockKind.Quote || Kind == MarkdownBlockKind.ListItem;

        internal static string PlainText(IEnumerable<MarkdownRun> runs) => string.Concat(runs.Select(r => r.Break ? "\n" : r.Text));
    }

    /// <summary>解析结果：块、按类别计数的「按原文保留」警告，以及超出上限时的说明。</summary>
    internal sealed class MarkdownDocument
    {
        internal readonly List<MarkdownBlock> Blocks = new List<MarkdownBlock>();

        /// <summary>类别：html（HTML 标签）、image（图片，不下载）、link（相对地址等不支持的链接）。</summary>
        internal readonly Dictionary<string, int> Warnings = new Dictionary<string, int>();

        /// <summary>超出字数或段数上限时的说明；为 null 表示可以插入。</summary>
        internal string LimitError;

        internal int WarningCount => Warnings.Values.Sum();

        /// <summary>写到 OneNote 的段落数：文字段、空行、表格单元格和代码行都算。</summary>
        internal int ParagraphCount => Blocks.Sum(b =>
            b.Kind == MarkdownBlockKind.Code ? CodeLines(b.Code) :
            b.Kind == MarkdownBlockKind.Table ? b.Rows.Count * b.Rows.Max(r => r.Count) : 1);

        internal int TextCount => Blocks.Count(b => b.IsText);

        internal int TableCount => Blocks.Count(b => b.Kind == MarkdownBlockKind.Table);

        internal int CodeCount => Blocks.Count(b => b.Kind == MarkdownBlockKind.Code);

        internal void Warn(string kind)
        {
            Warnings.TryGetValue(kind, out var count);
            Warnings[kind] = count + 1;
        }

        private static int CodeLines(string code) => (code ?? string.Empty).TrimEnd().Count(c => c == '\n') + 1;
    }
}
