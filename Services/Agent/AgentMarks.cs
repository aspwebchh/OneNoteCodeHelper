using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace OneNoteCodeHelper.Services.Agent
{
    /// <summary>
    /// 段落前的列表符号（one:List）和标记（one:Tag）。两者都是 OE 的子元素，不改正文，
    /// 所以和段落样式一样走草稿、提交、回读核验和撤销。
    /// </summary>
    internal static class AgentMarks
    {
        private static XNamespace One => OneNoteApi.One;

        internal static readonly string[] ListKinds = { "bullet", "number", "none" };
        internal static readonly string[] TagKinds = { "todo", "important", "question" };

        /// <summary>OneNote 默认标记库的前三项：type 是在标记库里的序号，symbol 是图标（3 复选框、13 星号、15 问号）。</summary>
        private static readonly (string Kind, int Type, int Symbol, string Name)[] Presets =
            { ("todo", 0, 3, "待办事项"), ("important", 1, 13, "重要"), ("question", 2, 15, "问题") };

        /// <summary>OE 开头几个子元素按架构的先后顺序，之后才是正文和 OEChildren。</summary>
        private static readonly string[] Leading = { "MediaIndex", "Tag", "OutlookTask", "Meta", "List" };

        internal static string ListKind(XElement oe)
        {
            var list = oe.Element(One + "List");
            if (list == null) return "none";
            return list.Element(One + "Number") != null ? "number" : list.Element(One + "Bullet") != null ? "bullet" : "other";
        }

        /// <summary>同一种列表不动，保留原有样式和起点；恢复 Markdown 的新编号列表可指定起点，不写字体属性。</summary>
        internal static void SetList(XElement oe, string kind, int? start = null)
        {
            if (ListKind(oe) == kind) return;
            if (kind == "none") { oe.Elements(One + "List").Remove(); return; }
            var list = new XElement(One + "List", kind == "bullet"
                ? new XElement(One + "Bullet", new XAttribute("bullet", 2))
                : new XElement(One + "Number", new XAttribute("numberSequence", 0), new XAttribute("numberFormat", "##.")));
            if (kind == "number" && start != null) list.Element(One + "Number").SetAttributeValue("restartNumberingAt", start.Value);
            Replace(oe, "List", new[] { list });
        }

        /// <summary>
        /// kind 为 none 时去掉待办、重要、问题三种标记，其他标记保留。已有同类标记时只改完成状态，不重复添加。
        /// 新标记引用的 TagDef 记在 defs（草稿里的 TagDef）里。
        /// </summary>
        internal static void SetTag(XElement oe, string kind, bool? completed, XElement defs)
        {
            if (kind == "none")
            {
                foreach (var t in oe.Elements(One + "Tag").Where(t => KindOf(defs, t) != null).ToList()) t.Remove();
                return;
            }
            var tag = oe.Elements(One + "Tag").FirstOrDefault(t => KindOf(defs, t) == kind);
            if (tag == null)
            {
                tag = new XElement(One + "Tag", new XAttribute("index", ResolveTagDef(defs, kind)), new XAttribute("completed", "false"),
                    new XAttribute("disabled", "false"), new XAttribute("creationDate", Now()));
                Replace(oe, "Tag", oe.Elements(One + "Tag").Select(t => new XElement(t)).Concat(new[] { tag }).ToList());
                tag = oe.Elements(One + "Tag").Last();
            }
            if (completed == null || IsCompleted(tag) == completed.Value) return;
            tag.SetAttributeValue("completed", completed.Value ? "true" : "false");
            tag.SetAttributeValue("completionDate", completed.Value ? Now() : null);
        }

        /// <summary>草稿里复用同一图标的 TagDef（用户页面上已有的「待办事项」也算），没有才新增。</summary>
        internal static string ResolveTagDef(XElement defs, string kind)
        {
            var preset = Presets.First(p => p.Kind == kind);
            var match = defs.Elements(One + "TagDef").FirstOrDefault(d => (string)d.Attribute("symbol") == Invariant(preset.Symbol));
            if (match != null) return (string)match.Attribute("index");
            var index = NextIndex(defs);
            defs.Add(new XElement(One + "TagDef", new XAttribute("index", index), new XAttribute("type", preset.Type), new XAttribute("symbol", preset.Symbol),
                new XAttribute("fontColor", "automatic"), new XAttribute("highlightColor", "none"), new XAttribute("name", preset.Name)));
            return index;
        }

        /// <summary>提交时在重新读取的页面上找同样的 TagDef，没有就新增并另取编号。TagDef 排在页面最前面、QuickStyleDef 之前。</summary>
        internal static string EnsureTagDefinition(XElement page, XElement definition)
        {
            var all = page.Elements(One + "TagDef").ToList();
            var signature = Signature(definition);
            var match = all.FirstOrDefault(d => Signature(d) == signature);
            if (match != null) return (string)match.Attribute("index");
            var added = new XElement(definition);
            var index = NextIndex(page);
            added.SetAttributeValue("index", index);
            if (all.Count == 0) page.AddFirst(added); else all[all.Count - 1].AddAfterSelf(added);
            return index;
        }

        /// <summary>把 source 的标记和列表复制到 target。两边一样时不动，免得调整了用户段落里这些元素的位置。</summary>
        internal static void CopyMarks(XElement source, XElement target)
        {
            if (Same(source.Elements(One + "Tag"), target.Elements(One + "Tag")) && Same(source.Elements(One + "List"), target.Elements(One + "List"))) return;
            Replace(target, "Tag", source.Elements(One + "Tag").Select(e => new XElement(e)).ToList());
            Replace(target, "List", source.Elements(One + "List").Select(e => new XElement(e)).ToList());
        }

        /// <summary>
        /// 核验用的语义投影：列表看种类及编号样式、起点，标记看图标和完成状态。
        /// OneNote 回存时补上的 fontSize、编号文字、创建时间等不参与比较，TagDef 重新编号也不影响。
        /// </summary>
        internal static string Projection(XElement oe, XElement page) => "list=" + ListKey(oe) + "|tags=" +
            string.Join(",", oe.Elements(One + "Tag").Select(t => ((string)Definition(page, t)?.Attribute("symbol") ?? "?") + ":" + (IsCompleted(t) ? "1" : "0")));

        /// <summary>发布草稿时判断有没有改动：新 TagDef 还不在快照页面里，所以直接比编号。</summary>
        internal static string DraftKey(XElement oe) => ListKey(oe) + "|" +
            string.Join(",", oe.Elements(One + "Tag").Select(t => (string)t.Attribute("index") + ":" + (IsCompleted(t) ? "1" : "0")));

        private static string ListKey(XElement oe)
        {
            var number = oe.Element(One + "List")?.Element(One + "Number");
            return number == null ? ListKind(oe) : "number:" + NumberAttribute(number, "numberSequence") + ":" +
                (string)number.Attribute("numberFormat") + ":" + NumberAttribute(number, "restartNumberingAt");
        }

        private static string NumberAttribute(XElement number, string name)
        {
            var value = (string)number.Attribute(name);
            return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? Invariant(n) : value ?? "";
        }

        /// <summary>给模型看的标记：todo、todo:done、important、question，其他标记记为 other。</summary>
        internal static string[] Describe(XElement oe, XElement defs) => oe.Elements(One + "Tag")
            .Select(t => (KindOf(defs, t) ?? "other") + (IsCompleted(t) ? ":done" : "")).ToArray();

        internal static XElement Definition(XElement defs, XElement tag) =>
            defs.Elements(One + "TagDef").FirstOrDefault(d => (string)d.Attribute("index") == (string)tag.Attribute("index"));

        private static string KindOf(XElement defs, XElement tag)
        {
            var symbol = (string)Definition(defs, tag)?.Attribute("symbol");
            return Presets.Where(p => Invariant(p.Symbol) == symbol).Select(p => p.Kind).FirstOrDefault();
        }

        private static bool IsCompleted(XElement tag) => (string)tag.Attribute("completed") == "true";

        /// <summary>换掉 OE 里所有名为 name 的子元素。原来有就放在原位置，没有就按架构顺序插入。</summary>
        private static void Replace(XElement oe, string name, IList<XElement> items)
        {
            var existing = oe.Elements(One + name).ToList();
            if (existing.Count > 0)
            {
                existing[0].AddBeforeSelf(items);
                foreach (var e in existing) e.Remove();
                return;
            }
            if (items.Count == 0) return;
            var rank = Array.IndexOf(Leading, name);
            var anchor = oe.Elements().TakeWhile(e => e.Name.Namespace == One && Array.IndexOf(Leading, e.Name.LocalName) is var r && r >= 0 && r <= rank).LastOrDefault();
            if (anchor == null) oe.AddFirst(items); else anchor.AddAfterSelf(items);
        }

        private static bool Same(IEnumerable<XElement> left, IEnumerable<XElement> right)
        {
            var a = left.ToList(); var b = right.ToList();
            return a.Count == b.Count && a.Zip(b, (x, y) => XNode.DeepEquals(x, y)).All(x => x);
        }

        /// <summary>TagDef 的内容（不含 index），用于按内容对应标记定义。</summary>
        internal static string Signature(XElement definition)
        {
            var values = definition.Attributes().Where(a => a.Name.LocalName != "index").ToDictionary(a => a.Name.LocalName, a => a.Value.Trim().ToLowerInvariant());
            if (!values.ContainsKey("fontColor")) values["fontColor"] = "automatic";
            if (!values.ContainsKey("highlightColor")) values["highlightColor"] = "none";
            return string.Join(";", values.OrderBy(v => v.Key, StringComparer.Ordinal).Select(v => v.Key + "=" + v.Value));
        }

        private static string NextIndex(XElement parent) => Invariant(parent.Elements(One + "TagDef")
            .Select(d => int.TryParse((string)d.Attribute("index"), out var n) ? n : -1).DefaultIfEmpty(-1).Max() + 1);

        private static string Invariant(int value) => value.ToString(CultureInfo.InvariantCulture);
        private static string Now() => DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.000Z", CultureInfo.InvariantCulture);
    }
}
