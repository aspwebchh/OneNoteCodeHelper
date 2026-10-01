using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OneNoteCodeHelper.Services.Agent
{
    /// <summary>表格外观：边框、标题行和各单元格的底色。不含单元格文字和行列结构。</summary>
    internal sealed class TableLook
    {
        private static XNamespace One => OneNoteApi.One;

        /// <summary>表头底色：浅蓝、浅灰、浅绿、浅黄。</summary>
        internal static readonly string[] Shadings = { "#DEEAF6", "#F2F2F2", "#E2EFDA", "#FFF2CC" };

        internal bool Borders;
        internal bool HeaderRow;
        /// <summary>首行每个单元格的底色，没有底色记为空串。逐格记录，撤销时能还原各格不同的底色。</summary>
        internal List<string> HeaderShades = new List<string>();
        /// <summary>首行以外每个单元格的底色，按行、列的先后顺序。同样逐格记录。</summary>
        internal List<string> BodyShades = new List<string>();

        /// <summary>给模型看的首行底色：各格一致时是颜色或 none，不一致是 mixed。</summary>
        internal string ShadingName => NameOf(HeaderShades);
        /// <summary>首行以外的单元格底色，写法同 <see cref="ShadingName"/>；只有一行时为 none。</summary>
        internal string BodyShadingName => BodyShades.Count == 0 ? "none" : NameOf(BodyShades);

        private static string NameOf(List<string> shades)
        {
            var distinct = shades.Distinct().ToList();
            return distinct.Count != 1 ? "mixed" : distinct[0].Length == 0 ? "none" : distinct[0];
        }

        internal static TableLook Read(XElement table) => new TableLook
        {
            Borders = Flag(table, "bordersVisible"), HeaderRow = Flag(table, "hasHeaderRow"),
            HeaderShades = FirstRow(table).Select(c => Shade((string)c.Attribute("shadingColor"))).ToList(),
            BodyShades = BodyCells(table).Select(c => Shade((string)c.Attribute("shadingColor"))).ToList()
        };

        /// <summary>只写和表格现状不同的项，没改的属性保持 OneNote 原来的写法。</summary>
        internal void Apply(XElement table)
        {
            if (Flag(table, "bordersVisible") != Borders) table.SetAttributeValue("bordersVisible", Borders ? "true" : "false");
            if (Flag(table, "hasHeaderRow") != HeaderRow) table.SetAttributeValue("hasHeaderRow", HeaderRow ? "true" : "false");
            var cells = FirstRow(table).ToList();
            if (cells.Count != HeaderShades.Count) throw new AiException("表格首行已变化。");
            var body = BodyCells(table).ToList();
            if (body.Count != BodyShades.Count) throw new AiException("表格行列已变化。");
            foreach (var (cell, shade) in cells.Zip(HeaderShades, (c, s) => (c, s)).Concat(body.Zip(BodyShades, (c, s) => (c, s))))
                if (Shade((string)cell.Attribute("shadingColor")) != shade)
                    cell.SetAttributeValue("shadingColor", shade.Length == 0 ? null : shade);
        }

        /// <summary>
        /// 按 set_table_style 的参数改出一份新外观，没给的项保持不变。cell_shading 为 none 时去掉全部单元格（含首行）的底色，
        /// 同时给了 header_shading 时首行再按它设置。
        /// </summary>
        internal TableLook With(IDictionary<string, object> style)
        {
            var look = new TableLook { Borders = Borders, HeaderRow = HeaderRow, HeaderShades = HeaderShades.ToList(), BodyShades = BodyShades.ToList() };
            if (style.TryGetValue("borders", out var borders)) look.Borders = (bool)borders;
            if (style.TryGetValue("header_row", out var header)) look.HeaderRow = (bool)header;
            if (style.ContainsKey("cell_shading"))
            {
                look.HeaderShades = HeaderShades.Select(_ => "").ToList();
                look.BodyShades = BodyShades.Select(_ => "").ToList();
            }
            if (style.TryGetValue("header_shading", out var shading))
                look.HeaderShades = HeaderShades.Select(_ => (string)shading == "none" ? "" : Shade((string)shading)).ToList();
            return look;
        }

        internal bool SameAs(TableLook other) => other != null && Borders == other.Borders && HeaderRow == other.HeaderRow &&
            HeaderShades.SequenceEqual(other.HeaderShades) && BodyShades.SequenceEqual(other.BodyShades);

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
        internal static IEnumerable<XElement> BodyCells(XElement table) => table.Elements(One + "Row").Skip(1).SelectMany(r => r.Elements(One + "Cell"));
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
        /// 表格 ID、外观、行列数和全部单元格的 ID、底色。不含单元格文字：用户在处理期间改表格里的字不算冲突，
        /// 改外观、增删行列才算。
        /// </summary>
        internal static string TakeFingerprint(XElement table)
        {
            var look = TableLook.Read(table);
            var cells = table.Elements(One + "Row").SelectMany(r => r.Elements(One + "Cell"))
                .Select(c => (string)c.Attribute("objectID") + "=" + TableLook.Shade((string)c.Attribute("shadingColor")));
            return string.Join("|", (string)table.Attribute("objectID"), look.Borders, look.HeaderRow, table.Elements(One + "Row").Count(),
                table.Element(One + "Columns")?.Elements(One + "Column").Count() ?? 0, string.Join(",", cells));
        }
    }

    /// <summary>
    /// text_to_table：把段落拆成表格，有两种切法。
    /// <see cref="Build"/> 按固定规则切：制表符、|、空格，或一条记录分成几行（名称一行、地址一行）。默认每行文字成为表格的一行，
    /// 段内换行（Shift+Enter）分开的行也各成一行；linesPerRow 大于 1 时每几行合成一行，各行的单元格依次排开。空行和 Markdown 分隔行（|---|:--:|）去掉。
    /// <see cref="BuildFromRows"/> 用模型逐格给出的切分结果，按顺序在原文里逐字定位每一格。原文的每个字要么进了单元格，要么是分隔符，
    /// 要么是和该列 header 相同的标签（挪进了表头），否则报错；模型不能借此改写、调换、重复或丢掉文字。模型的字符串只用来定位，不写进页面。
    /// 两种切法的单元格都复制源段落的样式和文字，按范围截取，保留加粗、链接等局部格式。
    /// 文本不规整时不报错：各行（含 header）按最多的列数建表，缺的单元格留空。
    /// </summary>
    internal static class AgentTextTable
    {
        internal const int MaxRows = 100;
        internal const int MaxColumns = 10;
        internal const int MaxHeaderChars = 30;
        internal const int MaxCellChars = 500;
        /// <summary>none 不拆分，整行一格，配合 linesPerRow 使用。</summary>
        internal static readonly string[] Delimiters = { "tab", "pipe", "space", "none" };
        private static readonly char[] Breaks = { '\n', '\r', '\t' };
        /// <summary>rows 的单元格之间可以省略的分隔符；空白另算。</summary>
        private static readonly char[] Separators = { '\t', '|', '｜', ':', '：', ',', '，', ';', '；', '、', '=' };
        private static XNamespace One => OneNoteApi.One;

        /// <param name="paragraphs">按页面顺序排好的源段落（含中间空行）。</param>
        /// <param name="linesPerRow">一条记录占几行（空行不算），1 为每行一行。</param>
        /// <param name="shading">首行底色，null 为不设。</param>
        /// <param name="header">在首行前新增的列名，null 为不加；少于列数时补空，多于列数时按它加列。</param>
        /// <param name="padded">补了空单元格的行数（含 header 行）。</param>
        internal static XElement Build(IEnumerable<XElement> paragraphs, string delimiter, int linesPerRow, bool headerRow, bool borders, string shading,
            IList<string> header, out int padded)
        {
            var names = Names(header);
            var lines = new List<(XElement Source, string Text, List<(int Start, int Length)> Ranges)>();
            foreach (var oe in paragraphs)
            {
                var text = new AgentRichText(oe).Text;
                // 段内换行（<br>）在文字里是 \n，按它切成行。
                for (var from = 0; ;)
                {
                    var next = text.IndexOf('\n', from);
                    var to = next < 0 ? text.Length : next;
                    var ranges = string.IsNullOrWhiteSpace(text.Substring(from, to - from)) ? null : Split(text, from, to, delimiter);
                    if (ranges != null) lines.Add((oe, text, ranges));
                    if (next < 0) break;
                    from = next + 1;
                }
            }
            if (lines.Count == 0) throw new AiException("目标段落里没有可以转换的行。");
            var records = new List<List<(XElement Source, int Start, int Length)>>();
            for (var i = 0; i < lines.Count; i += linesPerRow)
            {
                var record = new List<(XElement Source, int Start, int Length)>();
                for (var k = 0; k < linesPerRow && i + k < lines.Count; k++)
                {
                    var (source, text, ranges) = lines[i + k];
                    foreach (var (start, length) in ranges) record.Add((source, start, length));
                    // 记录里不是最后一行的行尾冒号是「名称：」这类标签的分隔，和制表符、| 一样去掉。
                    if (k < linesPerRow - 1)
                    {
                        var last = record[record.Count - 1];
                        int from = last.Start, to = last.Start + last.Length;
                        if (to > from && (text[to - 1] == ':' || text[to - 1] == '：')) { to--; Trim(text, ref from, ref to); }
                        record[record.Count - 1] = (source, from, to - from);
                    }
                }
                records.Add(record);
            }
            if (records.Count > MaxRows) throw new AiException($"表格最多 {MaxRows} 行。");
            // 两列的检查只看原文：全都只有一列多半是分隔符选错了，补空也不成表。
            if (records.Max(r => r.Count) < 2) throw new AiException("按指定的分隔符分不出两列，请确认 delimiter；一条记录分成几行时用 lines_per_row；分隔不规整时用 rows 逐格给出。");
            return Assemble(records, lines[0].Source, headerRow, borders, shading, names, out padded);
        }

        /// <param name="paragraphs">按页面顺序排好的源段落（含中间空行）。</param>
        /// <param name="rows">模型给出的各行单元格，须逐字取自原文、按原文顺序；空字符串为空单元格。</param>
        internal static XElement BuildFromRows(IReadOnlyList<XElement> paragraphs, IList<IList<string>> rows, bool headerRow, bool borders, string shading,
            IList<string> header, out int padded)
        {
            var names = Names(header);
            // 各段文字用 \n 连成一条，和 read_blocks 给模型看的一致；段内换行（<br>）也是 \n。
            var texts = paragraphs.Select(oe => new AgentRichText(oe).Text).ToList();
            var all = string.Join("\n", texts);
            var starts = new List<int>();
            for (int i = 0, at = 0; i < texts.Count; at += texts[i].Length + 1, i++) starts.Add(at);
            if (rows.Count > MaxRows) throw new AiException($"表格最多 {MaxRows} 行。");
            var records = new List<List<(XElement Source, int Start, int Length)>>();
            var pos = 0;
            for (var i = 0; i < rows.Count; i++)
            {
                var record = new List<(XElement Source, int Start, int Length)>();
                for (var j = 0; j < rows[i].Count; j++)
                {
                    var cell = rows[i][j].Trim();
                    if (cell.Length == 0) { record.Add((null, 0, 0)); continue; }
                    if (cell.IndexOf('\n') >= 0 || cell.IndexOf('\r') >= 0) throw new AiException($"rows 第 {i + 1} 行第 {j + 1} 格含换行；单元格不能跨行，请拆成几格。");
                    var at = Find(all, cell, pos, names != null && j < names.Count ? names[j] : null, out var stop);
                    if (at < 0)
                    {
                        if (!Enumerable.Range(pos, System.Math.Max(0, all.Length - cell.Length - pos + 1)).Any(k => Matches(all, k, cell)))
                            throw new AiException($"rows 第 {i + 1} 行第 {j + 1} 格「{Excerpt(cell, 0)}」没有在原文里按顺序找到：单元格须逐字复制原文、按原文顺序排列，不能改写、调换或重复。");
                        throw Uncovered(all, stop);
                    }
                    var p = starts.FindLastIndex(s => s <= at);
                    record.Add((paragraphs[p], at - starts[p], cell.Length));
                    pos = at + cell.Length;
                }
                records.Add(record);
            }
            // 最后一格之后只能剩分隔符。
            Find(all, null, pos, null, out var rest);
            if (rest >= 0) throw Uncovered(all, rest);
            return Assemble(records, paragraphs[0], headerRow, borders, shading, names, out padded);
        }

        /// <summary>
        /// 从 pos 起找 cell（任何空白彼此相等，换行除外），它前面的文字去掉空白和分隔符后须为空或等于 label。
        /// 找不到时返回 -1，stop 是第一个既不能省略、也没放进单元格的字，没有则为 -1。cell 为 null 时只找 stop。
        /// </summary>
        private static int Find(string text, string cell, int pos, string label, out int stop)
        {
            label = label == null ? null : new string(label.Where(ch => !Skippable(ch)).ToArray());
            var kept = "";
            var first = -1;
            for (var k = pos; k < text.Length; k++)
            {
                if (cell != null && (kept.Length == 0 || kept == label) && Matches(text, k, cell)) { stop = -1; return k; }
                if (Skippable(text[k])) continue;
                // 省略的字只能拼成列名；一旦对不上就不会再对上。
                if (first < 0) first = k;
                kept += text[k];
                if (label == null || !label.StartsWith(kept, System.StringComparison.Ordinal)) { stop = k; return -1; }
            }
            stop = first;
            return -1;
        }

        private static bool Skippable(char ch) => char.IsWhiteSpace(ch) || System.Array.IndexOf(Separators, ch) >= 0;

        private static bool Matches(string text, int start, string cell)
        {
            if (start + cell.Length > text.Length) return false;
            for (var i = 0; i < cell.Length; i++)
            {
                char a = text[start + i], b = cell[i];
                if (a != b && !(a != '\n' && a != '\r' && char.IsWhiteSpace(a) && char.IsWhiteSpace(b))) return false;
            }
            return true;
        }

        private static AiException Uncovered(string text, int start) =>
            new AiException($"原文「{Excerpt(text, start)}」没有放进任何单元格：除分隔符和与该列 header 相同的标签外，所选段落的文字都要逐字放进单元格；标题等不属于表格的段落不要放进 block_ids。");

        /// <summary>从 start 起摘一小段，到换行为止。</summary>
        private static string Excerpt(string text, int start)
        {
            var end = text.IndexOf('\n', start);
            var length = System.Math.Min((end < 0 ? text.Length : end) - start, 20);
            return text.Substring(start, length) + (start + length < text.Length && text[start + length] != '\n' ? "…" : "");
        }

        private static List<string> Names(IList<string> header)
        {
            var names = header?.Select(h => (h ?? "").Trim()).ToList();
            if (names != null && names.Any(n => n.Length == 0 || n.Length > MaxHeaderChars || n.IndexOfAny(Breaks) >= 0))
                throw new AiException($"header 的每一项须是 1–{MaxHeaderChars} 字、不含换行和制表符的列名。");
            return names;
        }

        /// <summary>按最多的列数建表，缺的单元格留空。Source 为 null 的是空单元格。</summary>
        private static XElement Assemble(List<List<(XElement Source, int Start, int Length)>> records, XElement styleSource, bool headerRow, bool borders, string shading,
            List<string> names, out int padded)
        {
            var columns = System.Math.Max(records.Max(r => r.Count), names?.Count ?? 0);
            if (columns > MaxColumns) throw new AiException($"表格最多 {MaxColumns} 列。");
            var rows = records.Select(r => r.Select(c => c.Source == null ? Blank() : Cell(c.Source, c.Start, c.Length)).ToList()).ToList();
            // 新增的列名按纯文字转义，段落样式随第一段。
            if (names != null)
                rows.Insert(0, names.Select(n => new XElement(One + "OE", styleSource.Attribute("style"), styleSource.Attribute("lang"),
                    new XElement(One + "T", new XCData(OneNoteHtmlEncoder.EncodePlainText(n))))).ToList());
            // 未锁定的列宽由 OneNote 按内容自动计算（本机实测），width 只是架构要求的占位值。
            var table = new XElement(One + "Table", new XAttribute("bordersVisible", borders ? "true" : "false"), new XAttribute("hasHeaderRow", headerRow ? "true" : "false"),
                new XElement(One + "Columns", Enumerable.Range(0, columns).Select(i => new XElement(One + "Column", new XAttribute("index", i), new XAttribute("width", 100)))));
            padded = rows.Count(r => r.Count < columns);
            for (var i = 0; i < rows.Count; i++)
            {
                var row = new XElement(One + "Row");
                for (var j = 0; j < columns; j++)
                {
                    var cell = new XElement(One + "Cell", new XElement(One + "OEChildren", j < rows[i].Count ? rows[i][j] : Blank()));
                    if (i == 0 && shading != null) cell.SetAttributeValue("shadingColor", shading);
                    row.Add(cell);
                }
                table.Add(row);
            }
            return table;
        }

        private static XElement Blank() => new XElement(One + "OE", new XElement(One + "T", new XCData("")));

        /// <summary>一行文字 [start, end) 里各单元格的范围（已去掉两侧空白）。Markdown 分隔行返回 null。</summary>
        private static List<(int Start, int Length)> Split(string text, int start, int end, string delimiter)
        {
            if (delimiter == "none")
            {
                Trim(text, ref start, ref end);
                return new List<(int Start, int Length)> { (start, end - start) };
            }
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
