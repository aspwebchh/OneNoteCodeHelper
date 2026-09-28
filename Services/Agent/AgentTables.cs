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

    /// <summary>表格样式的撤销记录：表格外观没被再改过时换回 Before。</summary>
    internal sealed class AgentTableUndoItem
    {
        internal string ObjectId;
        internal TableLook Before;
        internal string AfterFingerprint;
    }
}
