using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OneNoteCodeHelper.Services.Agent
{
    /// <summary>表格外观：边框、标题行和首行各单元格的底色。不含单元格文字和行列结构。</summary>
    internal sealed class TableLook
    {
        private static XNamespace One => OneNoteApi.One;

        /// <summary>表头底色：浅蓝、浅灰、浅绿、浅黄。</summary>
        internal static readonly string[] Shadings = { "#DEEAF6", "#F2F2F2", "#E2EFDA", "#FFF2CC" };

        internal bool Borders;
        internal bool HeaderRow;
        /// <summary>首行每个单元格的底色，没有底色记为空串。逐格记录，撤销时能还原各格不同的底色。</summary>
        internal List<string> HeaderShades = new List<string>();

        /// <summary>给模型看的首行底色：各格一致时是颜色或 none，不一致是 mixed。</summary>
        internal string ShadingName
        {
            get
            {
                var distinct = HeaderShades.Distinct().ToList();
                return distinct.Count != 1 ? "mixed" : distinct[0].Length == 0 ? "none" : distinct[0];
            }
        }

        internal static TableLook Read(XElement table) => new TableLook
        {
            Borders = Flag(table, "bordersVisible"), HeaderRow = Flag(table, "hasHeaderRow"),
            HeaderShades = FirstRow(table).Select(c => Shade((string)c.Attribute("shadingColor"))).ToList()
        };

        /// <summary>只写和表格现状不同的项，没改的属性保持 OneNote 原来的写法。</summary>
        internal void Apply(XElement table)
        {
            if (Flag(table, "bordersVisible") != Borders) table.SetAttributeValue("bordersVisible", Borders ? "true" : "false");
            if (Flag(table, "hasHeaderRow") != HeaderRow) table.SetAttributeValue("hasHeaderRow", HeaderRow ? "true" : "false");
            var cells = FirstRow(table).ToList();
            if (cells.Count != HeaderShades.Count) throw new AiException("表格首行已变化。");
            for (var i = 0; i < cells.Count; i++)
                if (Shade((string)cells[i].Attribute("shadingColor")) != HeaderShades[i])
                    cells[i].SetAttributeValue("shadingColor", HeaderShades[i].Length == 0 ? null : HeaderShades[i]);
        }

        /// <summary>按 set_table_style 的参数改出一份新外观，没给的项保持不变。</summary>
        internal TableLook With(IDictionary<string, object> style)
        {
            var look = new TableLook { Borders = Borders, HeaderRow = HeaderRow, HeaderShades = HeaderShades.ToList() };
            if (style.TryGetValue("borders", out var borders)) look.Borders = (bool)borders;
            if (style.TryGetValue("header_row", out var header)) look.HeaderRow = (bool)header;
            if (style.TryGetValue("header_shading", out var shading))
                look.HeaderShades = HeaderShades.Select(_ => (string)shading == "none" ? "" : Shade((string)shading)).ToList();
            return look;
        }

        internal bool SameAs(TableLook other) => other != null && Borders == other.Borders && HeaderRow == other.HeaderRow && HeaderShades.SequenceEqual(other.HeaderShades);

        /// <summary>OneNote 可能省略值为 false 的属性。</summary>
        internal static bool Flag(XElement e, string name)
        {
            var value = ((string)e.Attribute(name) ?? "").Trim().ToLowerInvariant();
            return value == "true" || value == "1";
        }

        /// <summary>没有底色的几种写法（缺省、none、automatic）统一成空串，颜色统一大写。</summary>
        internal static string Shade(string value)
        {
            value = (value ?? "").Trim();
            return value.Length == 0 || value.Equals("none", System.StringComparison.OrdinalIgnoreCase) ||
                value.Equals("automatic", System.StringComparison.OrdinalIgnoreCase) ? "" : value.ToUpperInvariant();
        }

        private static IEnumerable<XElement> FirstRow(XElement table) => table.Element(One + "Row")?.Elements(One + "Cell") ?? Enumerable.Empty<XElement>();
    }

    /// <summary>快照里的一个表格。短 ID 为 t1、t2…；代码框和不支持的文本框里的表格只保护。</summary>
    internal sealed class AgentTable
    {
        internal string Id;
        internal string ObjectId;
        internal string ContainerId;
        internal string ProtectedReason;
        /// <summary>快照时的外观指纹。提交时不一致就按冲突跳过；撤销时是写入后的指纹。</summary>
        internal string Fingerprint;
        /// <summary>首行文字摘录，帮模型认出是哪个表格。</summary>
        internal string Summary;
        internal int Rows;
        internal int Columns;
        internal TableLook Original;
        internal TableLook Draft;
        internal bool Editable => ProtectedReason == null;
        /// <summary>撤销时表格可能已被删除（Original 为 null），仍算作要处理，提交时按冲突跳过。</summary>
        internal bool Changed => Original == null || !Original.SameAs(Draft);

        private static XNamespace One => OneNoteApi.One;

        internal static XElement Find(XElement page, string objectId) =>
            page.Descendants(One + "Table").FirstOrDefault(t => (string)t.Attribute("objectID") == objectId);

        /// <summary>
        /// 表格 ID、外观、行列数和首行单元格的 ID。不含单元格文字：用户在处理期间改表格里的字不算冲突，
        /// 改外观、增删行列才算。
        /// </summary>
        internal static string TakeFingerprint(XElement table)
        {
            var look = TableLook.Read(table);
            var cells = table.Element(One + "Row")?.Elements(One + "Cell").Select(c => (string)c.Attribute("objectID") + "=" + TableLook.Shade((string)c.Attribute("shadingColor")));
            return string.Join("|", (string)table.Attribute("objectID"), look.Borders, look.HeaderRow, table.Elements(One + "Row").Count(),
                table.Element(One + "Columns")?.Elements(One + "Column").Count() ?? 0, string.Join(",", cells ?? Enumerable.Empty<string>()));
        }
    }

    /// <summary>
    /// text_to_table：把用制表符、| 或空格分隔的文字拆成表格。每行文字成为表格的一行，段内换行（Shift+Enter）分开的行也各成一行；
    /// 空行和 Markdown 分隔行（|---|:--:|）去掉。单元格复制源段落的样式和文字，按范围截取，保留加粗、链接等局部格式。
    /// 空格分隔有歧义（单元格里本身可能有空格，标题只拆得出一列），各行列数必须相同；制表符和 | 不足的列补空格。
    /// </summary>
    internal static class AgentTextTable
    {
        internal const int MaxRows = 100;
        internal const int MaxColumns = 10;
        internal const int MaxHeaderChars = 30;
        internal static readonly string[] Delimiters = { "tab", "pipe", "space" };
        private static readonly char[] Breaks = { '\n', '\r', '\t' };
        private static XNamespace One => OneNoteApi.One;

        /// <param name="paragraphs">按页面顺序排好的源段落（含中间空行）。</param>
        /// <param name="shading">首行底色，null 为不设。</param>
        /// <param name="header">在首行前新增的列名，null 为不加；个数须等于列数。</param>
        internal static XElement Build(IEnumerable<XElement> paragraphs, string delimiter, bool headerRow, bool borders, string shading, IList<string> header = null)
        {
            var lines = new List<(XElement Source, string Text, List<(int Start, int Length)> Ranges)>();
            foreach (var oe in paragraphs)
            {
                var text = new AgentRichText(oe).Text;
                // 段内换行（<br>）在文字里是 \n，按它切成行。
                for (var from = 0; ;)
                {
                    var next = text.IndexOf('\n', from);
                    var to = next < 0 ? text.Length : next;
                    var line = text.Substring(from, to - from);
                    var ranges = string.IsNullOrWhiteSpace(line) ? null : Split(text, from, to, delimiter);
                    if (ranges != null) lines.Add((oe, line.Trim(), ranges));
                    if (next < 0) break;
                    from = next + 1;
                }
            }
            if (lines.Count == 0) throw new AiException("目标段落里没有可以转换的行。");
            if (lines.Count > MaxRows) throw new AiException($"表格最多 {MaxRows} 行。");
            var columns = lines.Max(l => l.Ranges.Count);
            if (columns < 2) throw new AiException("按指定的分隔符分不出两列，请确认 delimiter。");
            if (columns > MaxColumns) throw new AiException($"表格最多 {MaxColumns} 列。");
            if (delimiter == "space")
            {
                // 以最常见的列数为准，一样多时取列数多的，标题行就是那个例外。
                var usual = lines.GroupBy(l => l.Ranges.Count).OrderByDescending(g => g.Count()).ThenByDescending(g => g.Key).First().Key;
                var odd = lines.FirstOrDefault(l => l.Ranges.Count != usual);
                if (odd.Source != null)
                    throw new AiException($"按空格拆分后各行列数不一致：「{Excerpt(odd.Text)}」拆出 {odd.Ranges.Count} 列，其他行是 {usual} 列。" +
                        "标题等不含分隔符的段落不要放进 block_ids；单元格里本身有空格时改用 tab 或 pipe。");
            }
            var rows = lines.Select(l => l.Ranges.Select(r => Cell(l.Source, r.Start, r.Length)).ToList()).ToList();
            if (header != null)
            {
                var names = header.Select(h => (h ?? "").Trim()).ToList();
                if (names.Count != columns) throw new AiException($"header 的列数必须等于表格列数 {columns}。");
                if (names.Any(n => n.Length == 0 || n.Length > MaxHeaderChars || n.IndexOfAny(Breaks) >= 0))
                    throw new AiException($"header 的每一项须是 1–{MaxHeaderChars} 字、不含换行和制表符的列名。");
                // 新增的列名按纯文字转义，段落样式随第一行。
                var source = lines[0].Source;
                rows.Insert(0, names.Select(n => new XElement(One + "OE", source.Attribute("style"), source.Attribute("lang"),
                    new XElement(One + "T", new XCData(OneNoteHtmlEncoder.EncodePlainText(n))))).ToList());
            }
            // 未锁定的列宽由 OneNote 按内容自动计算（本机实测），width 只是架构要求的占位值。
            var table = new XElement(One + "Table", new XAttribute("bordersVisible", borders ? "true" : "false"), new XAttribute("hasHeaderRow", headerRow ? "true" : "false"),
                new XElement(One + "Columns", Enumerable.Range(0, columns).Select(i => new XElement(One + "Column", new XAttribute("index", i), new XAttribute("width", 100)))));
            for (var i = 0; i < rows.Count; i++)
            {
                var row = new XElement(One + "Row");
                for (var j = 0; j < columns; j++)
                {
                    var cell = new XElement(One + "Cell", new XElement(One + "OEChildren", j < rows[i].Count ? rows[i][j] : new XElement(One + "OE", new XElement(One + "T", new XCData("")))));
                    if (i == 0 && shading != null) cell.SetAttributeValue("shadingColor", shading);
                    row.Add(cell);
                }
                table.Add(row);
            }
            return table;
        }

        /// <summary>一行文字 [start, end) 里各单元格的范围（已去掉两侧空白）。Markdown 分隔行返回 null。</summary>
        private static List<(int Start, int Length)> Split(string text, int start, int end, string delimiter)
        {
            if (delimiter == "space")
            {
                // 连续的空白（含 &nbsp; 和全角空格）算一个分隔。
                var cells = new List<(int Start, int Length)>();
                for (var i = start; i < end;)
                {
                    while (i < end && char.IsWhiteSpace(text[i])) i++;
                    if (i == end) break;
                    var s = i;
                    while (i < end && !char.IsWhiteSpace(text[i])) i++;
                    cells.Add((s, i - s));
                }
                return cells;
            }
            var separator = delimiter == "tab" ? '\t' : '|';
            if (separator == '|')
            {
                Trim(text, ref start, ref end);
                if (start < end && text[start] == '|') start++;
                if (end > start && text[end - 1] == '|') end--;
            }
            var ranges = new List<(int Start, int Length)>();
            for (var from = start; ;)
            {
                var next = text.IndexOf(separator, from, end - from);
                var to = next < 0 ? end : next;
                int s = from, e = to;
                Trim(text, ref s, ref e);
                ranges.Add((s, e - s));
                if (next < 0) break;
                from = next + 1;
            }
            if (separator == '|' && ranges.All(r => IsRule(text.Substring(r.Start, r.Length)))) return null;
            return ranges;
        }

        private static string Excerpt(string text) => text.Length <= 20 ? text : text.Substring(0, 20) + "…";

        private static bool IsRule(string cell) => cell.Length > 0 && cell.Contains("-") && cell.All(ch => ch == '-' || ch == ':');

        private static void Trim(string text, ref int start, ref int end)
        {
            while (start < end && char.IsWhiteSpace(text[start])) start++;
            while (end > start && char.IsWhiteSpace(text[end - 1])) end--;
        }

        /// <summary>复制源段落的样式和文字，只留下这一格的范围。quickStyleIndex 不带，单元格按正文显示。</summary>
        private static XElement Cell(XElement source, int start, int length)
        {
            var oe = new XElement(One + "OE", source.Attribute("style"), source.Attribute("lang"), source.Elements(One + "T").Select(t => new XElement(t)));
            new AgentRichText(oe).Keep(start, length);
            // 截掉的 T 只剩空内容，留一个就够。
            foreach (var t in oe.Elements(One + "T").Where(t => t.Value.Length == 0).Skip(oe.Elements(One + "T").All(t => t.Value.Length == 0) ? 1 : 0).ToList()) t.Remove();
            return oe;
        }
    }

    /// <summary>表格样式的撤销记录：表格外观没被再改过时换回 Before。</summary>
    internal sealed class AgentTableUndoItem
    {
        internal string ObjectId;
        internal TableLook Before;
        internal string AfterFingerprint;
    }
}
