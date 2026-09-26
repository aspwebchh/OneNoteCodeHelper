using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace OneNoteCodeHelper.Services.Agent
{
    /// <summary>受限 HTML 的无损语义模型。未知标签/属性不进入格式编辑器。</summary>
    internal sealed class AgentRichText
    {
        private sealed class Piece
        {
            internal int Run;
            internal string Text;
            internal bool Break;
            internal List<XElement> Path;
        }

        private readonly XElement _oe;
        private readonly List<Piece> _pieces = new List<Piece>();
        private static readonly HashSet<string> Tags = new HashSet<string>(
            new[] { "span", "a", "b", "strong", "i", "em", "u", "s", "strike", "sub", "sup", "br" });
        internal string Text => string.Concat(_pieces.Select(p => p.Text));

        internal AgentRichText(XElement oe)
        {
            _oe = oe;
            var run = 0;
            foreach (var t in oe.Elements(OneNoteApi.One + "T"))
            {
                var html = NormalizeTags(t.Value);
                html = Regex.Replace(html, @"&(?:#[xX][0-9a-fA-F]+|#\d+|[a-zA-Z]+);", m =>
                {
                    var decoded = WebUtility.HtmlDecode(m.Value);
                    return System.Security.SecurityElement.Escape(decoded);
                });
                using (var reader = XmlReader.Create(new System.IO.StringReader("<root>" + html + "</root>"),
                    new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 200000 }))
                {
                    Read(XElement.Load(reader, LoadOptions.PreserveWhitespace), run++, new List<XElement>());
                }
            }
        }

        private static string NormalizeTags(string html)
        {
            // OneNote 输出的是 HTML：lang=zh-CN 等无引号属性合法，不能直接按 XML 读取。
            var output = new StringBuilder();
            for (var i = 0; i < html.Length;)
            {
                if (html[i] != '<') { output.Append(html[i++]); continue; }
                var end = i + 1;
                var quote = '\0';
                for (; end < html.Length; end++)
                {
                    var ch = html[end];
                    if (quote != '\0') { if (ch == quote) quote = '\0'; }
                    else if (ch == '\'' || ch == '"') quote = ch;
                    else if (ch == '>') break;
                }
                if (end == html.Length) throw new AiException("HTML 标签未闭合。");
                var inner = html.Substring(i + 1, end - i - 1).Trim();
                var close = inner.StartsWith("/", StringComparison.Ordinal);
                if (close) inner = inner.Substring(1).Trim();
                var match = Regex.Match(inner, @"^[A-Za-z][A-Za-z0-9]*");
                if (!match.Success) throw new AiException("不支持的 HTML 标记。");
                var name = match.Value.ToLowerInvariant();
                if (!Tags.Contains(name)) throw new AiException("不支持的 HTML 标签。");
                var rest = inner.Substring(match.Length).Trim();
                if (close)
                {
                    if (rest.Length != 0) throw new AiException("无效的闭标签。");
                    output.Append("</").Append(name).Append('>');
                }
                else
                {
                    if (rest.EndsWith("/", StringComparison.Ordinal)) rest = rest.Substring(0, rest.Length - 1).TrimEnd();
                    var tag = new XElement(name);
                    while (rest.Length > 0)
                    {
                        var attribute = Regex.Match(rest, "^([A-Za-z][A-Za-z0-9_-]*)\\s*=\\s*(?:\"([^\"]*)\"|'([^']*)'|([^\\s]+))");
                        if (!attribute.Success) throw new AiException("不支持的 HTML 属性语法。");
                        var key = attribute.Groups[1].Value.ToLowerInvariant();
                        var value = attribute.Groups[2].Success ? attribute.Groups[2].Value : attribute.Groups[3].Success ? attribute.Groups[3].Value : attribute.Groups[4].Value;
                        if (tag.Attribute(key) != null) throw new AiException("重复的 HTML 属性。");
                        tag.Add(new XAttribute(key, WebUtility.HtmlDecode(value)));
                        rest = rest.Substring(attribute.Length).TrimStart();
                    }
                    var xml = tag.ToString(SaveOptions.DisableFormatting);
                    output.Append(name == "br" ? xml : xml.Substring(0, xml.Length - 3) + ">");
                }
                i = end + 1;
            }
            return output.ToString();
        }

        private void Read(XElement parent, int run, List<XElement> path)
        {
            foreach (var node in parent.Nodes())
            {
                if (node is XText text)
                {
                    _pieces.Add(new Piece { Run = run, Text = text.Value, Path = path });
                    continue;
                }
                var e = node as XElement;
                if (e == null || e.Name.NamespaceName.Length != 0 || !Tags.Contains(e.Name.LocalName))
                    throw new AiException("该段落包含不支持的 HTML。");
                if (e.Attributes().Any(a => a.Name.NamespaceName.Length != 0 ||
                    !new[] { "style", "href", "lang", "title" }.Contains(a.Name.LocalName)))
                    throw new AiException("该段落包含不支持的 HTML 属性。");
                if (e.Name.LocalName == "br")
                {
                    if (e.HasAttributes || e.HasElements) throw new AiException("不支持的换行格式。");
                    _pieces.Add(new Piece { Run = run, Text = "\n", Break = true, Path = path });
                }
                else
                {
                    if (!e.Nodes().Any()) throw new AiException("包含空的 HTML 标记，保留原样。");
                    var next = new List<XElement>(path) { new XElement(e.Name, e.Attributes()) };
                    Read(e, run, next);
                }
            }
        }

        internal void Format(int start, int length, IDictionary<string, string> properties, bool removeOnly = false)
        {
            if (start < 0 || length < 0 || start + length > Text.Length) throw new AiException("文字范围无效。");
            var boundaries = new HashSet<int>(StringInfo.ParseCombiningCharacters(Text)) { Text.Length };
            if (!boundaries.Contains(start) || !boundaries.Contains(start + length) || !EmojiBoundary(Text, start) || !EmojiBoundary(Text, start + length))
                throw new AiException("文字范围不能切开 Unicode 字符。");
            var output = new List<Piece>();
            var offset = 0;
            foreach (var p in _pieces)
            {
                var from = Math.Max(0, start - offset);
                var to = Math.Min(p.Text.Length, start + length - offset);
                if (from >= to) output.Add(p);
                else
                {
                    if (from > 0) output.Add(Copy(p, p.Text.Substring(0, from), p.Path));
                    var applied = new Dictionary<string, string>(properties);
                    if (applied.TryGetValue("text-decoration", out var decoration))
                    {
                        var strike = p.Path.Any(tag => tag.Name.LocalName == "s" || tag.Name.LocalName == "strike" ||
                            ((string)tag.Attribute("style") ?? "").IndexOf("line-through", StringComparison.OrdinalIgnoreCase) >= 0);
                        if (strike) applied["text-decoration"] = decoration == "none" ? "line-through" : "line-through " + decoration;
                    }
                    var path = p.Path.Select(tag => Rewrite(tag, applied.Keys)).ToList();
                    if (!removeOnly && !p.Break)
                        path.Add(new XElement("span", new XAttribute("style", Css.Write(applied))));
                    output.Add(Copy(p, p.Text.Substring(from, to - from), path));
                    if (to < p.Text.Length) output.Add(Copy(p, p.Text.Substring(to), p.Path));
                }
                offset += p.Text.Length;
            }
            _pieces.Clear();
            _pieces.AddRange(output);
            Save();
        }

        private static bool EmojiBoundary(string text, int index)
        {
            if (index == 0 || index == text.Length) return true;
            if (text[index] == '\u200d' || text[index - 1] == '\u200d') return false;
            var next = char.ConvertToUtf32(text, index);
            if (next >= 0x1f3fb && next <= 0x1f3ff) return false;
            var previousIndex = char.IsLowSurrogate(text[index - 1]) ? index - 2 : index - 1;
            var previous = char.ConvertToUtf32(text, previousIndex);
            return !(previous >= 0x1f1e6 && previous <= 0x1f1ff && next >= 0x1f1e6 && next <= 0x1f1ff);
        }

        private static Piece Copy(Piece p, string text, List<XElement> path) =>
            new Piece { Run = p.Run, Text = text, Break = p.Break, Path = path };

        private static XElement Rewrite(XElement tag, ICollection<string> keys)
        {
            var clone = new XElement(tag);
            var implicitProperty = ImplicitProperty(tag.Name.LocalName);
            if (implicitProperty != null && keys.Contains(implicitProperty)) clone.Name = "span";
            var css = Css.Read((string)clone.Attribute("style"));
            foreach (var key in keys)
            {
                if (key == "font-family" && css.TryGetValue(key, out var font) && (font.Contains("emoji") || font.Contains("symbol"))) continue;
                css.Remove(key);
            }
            clone.SetAttributeValue("style", css.Count == 0 ? null : Css.Write(css));
            return clone;
        }

        private static string ImplicitProperty(string name)
        {
            switch (name)
            {
                case "b": case "strong": return "font-weight";
                case "i": case "em": return "font-style";
                case "u": case "s": case "strike": return "text-decoration";
                case "sub": case "sup": return "vertical-align";
                default: return null;
            }
        }

        private void Save()
        {
            var runs = _oe.Elements(OneNoteApi.One + "T").ToList();
            for (var run = 0; run < runs.Count; run++)
            {
                var root = new XElement("root");
                var open = new List<XElement>();
                var templates = new List<XElement>();
                foreach (var p in _pieces.Where(p => p.Run == run))
                {
                    var common = 0;
                    while (common < templates.Count && common < p.Path.Count && XNode.DeepEquals(templates[common], p.Path[common])) common++;
                    open.RemoveRange(common, open.Count - common);
                    templates.RemoveRange(common, templates.Count - common);
                    for (var i = common; i < p.Path.Count; i++)
                    {
                        var e = new XElement(p.Path[i]);
                        (open.LastOrDefault() ?? root).Add(e);
                        open.Add(e);
                        templates.Add(p.Path[i]);
                    }
                    (open.LastOrDefault() ?? root).Add(p.Break ? (XNode)new XElement("br") : new XText(p.Text));
                }
                var html = string.Concat(root.Nodes().Select(n => n.ToString(SaveOptions.DisableFormatting))).Replace("\u00a0", "&nbsp;");
                runs[run].ReplaceNodes(new XCData(html));
            }
        }

        /// <summary>按字符投影，容忍 OneNote 合并/拆分 span 和 T。用于回读检查。</summary>
        internal string Signature(XElement page, bool includeStyles)
        {
            var result = new StringBuilder();
            var runs = _oe.Elements(OneNoteApi.One + "T").ToList();
            foreach (var p in _pieces)
            {
                var css = PieceStyle(p, runs, page, out var link);
                var style = includeStyles ? Css.Write(css) : "";
                foreach (var ch in p.Text) result.Append((int)ch).Append(':').Append(link.Length).Append(':').Append(link)
                    .Append(':').Append(style.Length).Append(':').Append(style).Append(';');
            }
            return result.ToString();
        }

        /// <summary>非空白字符里有没有、是不是全部是等宽字体。整段等宽是待高亮的代码，部分等宽是行内代码。</summary>
        internal (bool Any, bool All) Monospace(XElement page)
        {
            var runs = _oe.Elements(OneNoteApi.One + "T").ToList();
            bool any = false, all = true;
            foreach (var p in _pieces.Where(p => !string.IsNullOrWhiteSpace(p.Text)))
            {
                var mono = PieceStyle(p, runs, page, out _).TryGetValue("font-family", out var font) && Css.IsMonospace(font);
                any |= mono; all &= mono;
            }
            return (any, any && all);
        }

        private static Dictionary<string, string> PieceStyle(Piece p, List<XElement> runs, XElement page, out string link)
        {
            var css = Css.Effective(runs[p.Run], page);
            link = "";
            foreach (var tag in p.Path)
            {
                // OneNote 的默认链接颜色独立于段落基础色；显式的内层颜色仍然覆盖它。
                if (tag.Name.LocalName == "a") css["color"] = "link-default";
                var key = ImplicitProperty(tag.Name.LocalName);
                if (key != null) css[key] = tag.Name.LocalName == "u" ? "underline" :
                    tag.Name.LocalName == "s" || tag.Name.LocalName == "strike" ? "line-through" :
                    key == "font-weight" ? "bold" : key == "font-style" ? "italic" : tag.Name.LocalName;
                Css.Merge(css, Css.Read((string)tag.Attribute("style")));
                if (tag.Name.LocalName == "a") link = (string)tag.Attribute("href") ?? "";
            }
            css.Remove("text-align"); // 回存时常把 alignment 冗余写进 style，单独校验段落对齐。
            return css;
        }
    }

    internal static class Css
    {
        internal static Dictionary<string, string> Read(string value)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in (value ?? "").Split(';'))
            {
                var colon = item.IndexOf(':');
                if (colon > 0) result[item.Substring(0, colon).Trim().ToLowerInvariant()] = Normalize(item.Substring(colon + 1));
            }
            return result;
        }
        internal static string Normalize(string value)
        {
            value = value.Trim().Trim('\'', '"').ToLowerInvariant();
            if (value == "微软雅黑") return "microsoft yahei";
            if (value.EndsWith("pt", StringComparison.Ordinal) && double.TryParse(value.Substring(0, value.Length - 2), NumberStyles.Float,
                CultureInfo.InvariantCulture, out var number)) return number.ToString("0.###", CultureInfo.InvariantCulture) + "pt";
            return value;
        }
        private static readonly string[] MonospaceFonts =
            { "consolas", "nsimsun", "新宋体", "courier new", "courier", "lucida console", "cascadia code", "cascadia mono" };
        /// <summary>粘贴来的 HTML 可能带字体栈，OneNote 只认第一个字体。</summary>
        internal static bool IsMonospace(string font) => MonospaceFonts.Contains(Normalize((font ?? "").Split(',')[0]));
        internal static string Write(IDictionary<string, string> values) => string.Join(";", values.OrderBy(k => k.Key, StringComparer.Ordinal).Select(k => k.Key + ":" + Normalize(k.Value)));
        internal static void Merge(IDictionary<string, string> target, IDictionary<string, string> source)
        { foreach (var item in source) target[item.Key] = item.Value; }

        internal static Dictionary<string, string> Effective(XElement node, XElement page)
        {
            var result = Read("font-weight:normal;font-style:normal;text-decoration:none;vertical-align:baseline");
            foreach (var e in node.AncestorsAndSelf().Reverse())
            {
                var index = (string)e.Attribute("quickStyleIndex");
                var definition = index == null ? null : page.Elements(OneNoteApi.One + "QuickStyleDef").FirstOrDefault(d => (string)d.Attribute("index") == index);
                if (definition != null)
                {
                    Set(result, "font-family", (string)definition.Attribute("font"));
                    Set(result, "font-size", (string)definition.Attribute("fontSize") + "pt");
                    Set(result, "color", (string)definition.Attribute("fontColor"));
                    result["font-weight"] = (string)definition.Attribute("bold") == "true" ? "bold" : "normal";
                    result["font-style"] = (string)definition.Attribute("italic") == "true" ? "italic" : "normal";
                    result["text-decoration"] = (string)definition.Attribute("underline") == "true" ? "underline" : "none";
                }
                Merge(result, Read((string)e.Attribute("style")));
            }
            return result;
        }
        private static void Set(IDictionary<string, string> css, string key, string value)
        { if (!string.IsNullOrEmpty(value)) css[key] = Normalize(value); }
    }

    internal static class ParagraphStyles
    {
        internal static readonly string[] Ids = { "page_title", "heading1", "heading2", "body", "quote" };
        internal static readonly string[] Colors = { "#1F4E79", "#365F91", "#222222", "#666666" };
        internal static readonly string[] Fonts = { "Microsoft YaHei", "Calibri", "Arial" };

        internal static XElement Definition(string role, AgentOptions options)
        {
            var index = Array.IndexOf(Ids, role);
            var name = role == "heading1" ? "h1" : role == "heading2" ? "h2" : "p";
            return new XElement(OneNoteApi.One + "QuickStyleDef", new XAttribute("name", name),
                new XAttribute("font", options.FontFamily), new XAttribute("fontSize", new[] { 20d, 16, 13.5, 11, 10.5 }[index]),
                // 基础色由段落/文字补丁设置。把颜色放在 QuickStyleDef 会使 OneNote 给链接也补上正文色。
                new XAttribute("fontColor", "automatic"),
                new XAttribute("highlightColor", "automatic"), new XAttribute("bold", index <= 2 ? "true" : "false"),
                new XAttribute("spaceBefore", "0"), new XAttribute("spaceAfter", "0"));
        }

        internal static string EnsureDefinition(XElement parent, XElement definition)
        {
            var all = parent.Elements(OneNoteApi.One + "QuickStyleDef").ToList();
            var signature = DefinitionSignature(definition);
            var match = all.FirstOrDefault(d => DefinitionSignature(d) == signature);
            if (match != null) return (string)match.Attribute("index");
            var index = all.Select(d => int.TryParse((string)d.Attribute("index"), out var n) ? n : -1).DefaultIfEmpty(-1).Max() + 1;
            var added = new XElement(definition);
            added.SetAttributeValue("index", index);
            // TagDef 在 QuickStyleDef 前；其余 Page 子元素在后。
            var preceding = parent.Elements().LastOrDefault(e => e.Name == OneNoteApi.One + "TagDef" || e.Name == OneNoteApi.One + "QuickStyleDef");
            if (preceding == null) parent.AddFirst(added); else preceding.AddAfterSelf(added);
            return index.ToString(CultureInfo.InvariantCulture);
        }

        private static string DefinitionSignature(XElement e)
        {
            var values = e.Attributes().Where(a => a.Name.LocalName != "index").ToDictionary(a => a.Name.LocalName, a => Css.Normalize(a.Value));
            foreach (var flag in new[] { "bold", "italic", "underline", "strikethrough", "superscript", "subscript" })
                if (!values.ContainsKey(flag)) values[flag] = "false";
            foreach (var color in new[] { "fontColor", "highlightColor" }) if (!values.ContainsKey(color)) values[color] = "automatic";
            foreach (var number in new[] { "fontSize", "spaceBefore", "spaceAfter" })
                values[number] = double.TryParse((string)e.Attribute(number), NumberStyles.Float, CultureInfo.InvariantCulture, out var n)
                    ? n.ToString("0.###", CultureInfo.InvariantCulture) : "0";
            return string.Join(";", values.OrderBy(a => a.Key).Select(a => a.Key + "=" + a.Value));
        }
        internal static void Apply(XElement oe, string preset, IDictionary<string, object> overrides, AgentOptions options)
        {
            var index = Array.IndexOf(Ids, preset);
            if (index < 0) throw new AiException("未知样式。");
            var size = new[] { 20d, 16, 13.5, 11, 10.5 }[index];
            var color = new[] { Colors[0], Colors[0], Colors[1], Colors[2], Colors[3] }[index];
            var css = Css.Read((string)oe.Attribute("style"));
            css["font-family"] = options.FontFamily;
            css["font-size"] = size.ToString(CultureInfo.InvariantCulture) + "pt";
            css["color"] = color;
            oe.SetAttributeValue("alignment", "left");
            if (options.EnableParagraphSpacing)
            {
                oe.SetAttributeValue("spaceBefore", new[] { 0, 12, 8, 0, 4 }[index]);
                oe.SetAttributeValue("spaceAfter", new[] { 12, 6, 4, 4, 4 }[index]);
            }
            foreach (var item in overrides)
            {
                switch (item.Key)
                {
                    case "font_family": css["font-family"] = (string)item.Value; break;
                    case "font_size_pt": css["font-size"] = Convert.ToDouble(item.Value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture) + "pt"; break;
                    case "color": css["color"] = (string)item.Value; break;
                    case "alignment": oe.SetAttributeValue("alignment", item.Value); break;
                    case "space_before_pt": oe.SetAttributeValue("spaceBefore", item.Value); break;
                    case "space_after_pt": oe.SetAttributeValue("spaceAfter", item.Value); break;
                }
            }
            var hasChildren = oe.Elements(OneNoteApi.One + "OEChildren").Any();
            if (!hasChildren) oe.SetAttributeValue("style", Css.Write(css));
            var remove = new Dictionary<string, string> { ["font-family"] = "", ["font-size"] = "" };
            foreach (var t in oe.Elements(OneNoteApi.One + "T"))
            {
                var s = Css.Read((string)t.Attribute("style"));
                foreach (var key in remove.Keys) s.Remove(key);
                if (hasChildren)
                {
                    s["font-family"] = css["font-family"];
                    s["font-size"] = css["font-size"];
                    s["color"] = css["color"];
                }
                t.SetAttributeValue("style", s.Count == 0 ? null : Css.Write(s));
            }
            var rich = new AgentRichText(oe);
            rich.Format(0, rich.Text.Length, remove, true);
            if (index <= 2)
            {
                rich = new AgentRichText(oe);
                rich.Format(0, rich.Text.Length, new Dictionary<string, string> { ["font-weight"] = "bold" });
            }
        }
    }
}
