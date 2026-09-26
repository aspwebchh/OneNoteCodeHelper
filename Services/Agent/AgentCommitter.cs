using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;

namespace OneNoteCodeHelper.Services.Agent
{
    internal interface IOneNotePageAccess
    {
        string GetPageContent(string pageId, PageInfo info);
        void UpdatePageContent(string xml, DateTime expectedLastModified);
    }

    internal sealed class AgentUndoItem
    {
        internal string ObjectId;
        internal XElement Before;
        internal string AfterFingerprint;
    }

    internal sealed class AgentReport
    {
        internal string Status;
        internal string Message;
        internal int Applied;
        internal int Conflicts;
        internal int Unverified;
        internal int Protected;
        internal readonly List<AgentUndoItem> Undo = new List<AgentUndoItem>();
        internal readonly List<string> ConflictIds = new List<string>();
        internal object ToToolResult() => new { status = Status, applied = Applied, skipped_conflict = ConflictIds,
            unverified = Unverified, protected_count = Protected, message = Message };
    }

    internal sealed class AgentCommitter
    {
        private readonly IOneNotePageAccess _api;
        private static XNamespace One => OneNoteApi.One;
        internal AgentCommitter(IOneNotePageAccess api) { _api = api; }

        internal AgentReport Commit(AgentPageSnapshot snapshot, CancellationToken cancellation)
        {
            lock (PageEditCoordinator.ForPage(snapshot.PageId))
            {
                for (var attempt = 0; ; attempt++)
                {
                    cancellation.ThrowIfCancellationRequested();
                    var page = AgentPageSnapshot.ParsePage(_api.GetPageContent(snapshot.PageId, PageInfo.piBasic));
                    var report = new AgentReport { Protected = snapshot.Blocks.Count(b => !b.Editable) };
                    var untouched = new Dictionary<string, string>();
                    var fingerprints = new Dictionary<string, string>();
                    foreach (var oe in page.Descendants(One + "OE").Where(e => e.Elements(One + "T").Any()))
                    {
                        var id = (string)oe.Attribute("objectID");
                        if (id == null) continue;
                        fingerprints[id] = AgentPageSnapshot.Fingerprint(oe, page);
                        try { untouched[id] = AgentPageSnapshot.SemanticFormat(oe, page); }
                        catch (Exception) { /* 不支持的 HTML 由内容不变量检查保留。 */ }
                    }
                    var planned = new List<(AgentBlock Block, XElement Before, XElement Desired)>();
                    var containers = new HashSet<XElement>();
                    foreach (var block in snapshot.Blocks.Where(b => b.Changed))
                    {
                        var target = Find(page, block.ObjectId);
                        if (target == null || !fingerprints.TryGetValue(block.ObjectId, out var fingerprint) || fingerprint != block.Fingerprint)
                        { report.ConflictIds.Add(block.Id); continue; }
                        var before = new XElement(target);
                        var originalContent = new AgentRichText(target).Signature(page, false);
                        AgentPageSnapshot.CopyFormat(block.Draft, target);
                        var styleId = (string)block.Draft.Attribute("quickStyleIndex");
                        var definition = styleId == null ? null : snapshot.DraftStyles.Elements(One + "QuickStyleDef").FirstOrDefault(d => (string)d.Attribute("index") == styleId);
                        if (definition != null) target.SetAttributeValue("quickStyleIndex", ParagraphStyles.EnsureDefinition(page, definition));
                        if (new AgentRichText(target).Signature(page, false) != originalContent)
                            throw new AiException("格式修改改变了正文或链接，已阻止写入。");
                        planned.Add((block, before, new XElement(target)));
                        untouched.Remove(block.ObjectId);
                        containers.Add(target.Ancestors().First(e => e.Parent == page));
                    }
                    report.Conflicts = report.ConflictIds.Count;
                    if (planned.Count == 0)
                    {
                        report.Status = "NoChange";
                        report.Message = report.Conflicts > 0 ? $"没有写入：{report.Conflicts} 段在处理期间发生变化。" : "没有需要写入的格式修改。";
                        return report;
                    }
                    var timestamp = AgentPageSnapshot.Modified(page);
                    if (!UntouchedPreserved(untouched, page)) throw new AiException("这组格式会影响未指定的嵌套段落，已阻止写入。请分别选择段落处理。");
                    var xml = PageEditor.BuildPageChanges(snapshot.PageId, page.Elements().Where(e => containers.Contains(e) || e.Name == One + "QuickStyleDef").Select(e => new XElement(e)).ToArray());
                    cancellation.ThrowIfCancellationRequested();
                    var uncertain = false;
                    try { _api.UpdatePageContent(xml, timestamp); }
                    catch (COMException ex) when (ex.ErrorCode == unchecked((int)0x80042010) && attempt < 2) { continue; }
                    catch (COMException ex) when (ex.ErrorCode == unchecked((int)0x80042010))
                    { throw new AiException("页面持续变化，未提交；请停止编辑后重试。"); }
                    catch (Exception) { uncertain = true; }
                    // COM 已经开始后，取消也必须先确定实际写入结果。
                    XElement actual;
                    try { actual = AgentPageSnapshot.ParsePage(_api.GetPageContent(snapshot.PageId, PageInfo.piBasic)); }
                    catch (Exception)
                    {
                        report.Status = "CommitOutcomeUnknown";
                        report.Unverified = planned.Count;
                        report.Message = "已尝试写入，但无法回读确认。请查看目标页面，不要立即重复执行。";
                        return report;
                    }
                    foreach (var item in planned)
                    {
                        var written = Find(actual, item.Block.ObjectId);
                        try
                        {
                            if (written == null || AgentPageSnapshot.SemanticFormat(written, actual) != DesiredSignature(item.Desired, page, item.Block.ObjectId))
                            { report.Unverified++; continue; }
                            report.Applied++;
                            report.Undo.Add(new AgentUndoItem { ObjectId = item.Block.ObjectId, Before = item.Before,
                                AfterFingerprint = AgentPageSnapshot.Fingerprint(written, actual) });
                        }
                        catch (Exception) { report.Unverified++; }
                    }
                    // 原页面所有文字、链接、段落顺序必须保留，包括没有交给模型的对象。
                    if (!ContentPreserved(page, actual) || !UntouchedPreserved(untouched, actual))
                    {
                        report.Status = "CommitOutcomeUnknown";
                        report.Message = "写入后页面结构或内容与预期不一致，请检查目标页面。";
                        return report;
                    }
                    report.Status = report.Unverified > 0 || report.Conflicts > 0 ? "PartiallyApplied" : "Verified";
                    report.Message = $"已验证修改 {report.Applied} 段；冲突跳过 {report.Conflicts} 段；未验证 {report.Unverified} 段；保护 {report.Protected} 段。";
                    if (uncertain && report.Applied == 0) report.Message = "写回未得到确认，请检查页面。" + report.Message;
                    return report;
                }
            }
        }

        private static string DesiredSignature(XElement desired, XElement page, string objectId)
        {
            var copy = new XElement(page);
            var target = Find(copy, objectId);
            AgentPageSnapshot.CopyFormat(desired, target);
            return AgentPageSnapshot.SemanticFormat(target, copy);
        }

        internal AgentReport Undo(string pageId, AgentReport previous, AgentOptions options, CancellationToken cancellation)
        {
            lock (PageEditCoordinator.ForPage(pageId))
            {
                cancellation.ThrowIfCancellationRequested();
                var xml = _api.GetPageContent(pageId, PageInfo.piBasic);
                var snapshot = new AgentPageSnapshot(xml, new HashSet<string>(previous.Undo.Select(i => i.ObjectId)), options);
                var skipped = new List<string>();
                foreach (var item in previous.Undo)
                {
                    var block = snapshot.Blocks.FirstOrDefault(b => b.ObjectId == item.ObjectId);
                    if (block == null || !block.Editable || block.Fingerprint != item.AfterFingerprint)
                    { skipped.Add(item.ObjectId); continue; }
                    AgentPageSnapshot.CopyFormat(item.Before, block.Draft);
                }
                var report = Commit(snapshot, cancellation);
                report.Conflicts += skipped.Count;
                report.ConflictIds.AddRange(skipped);
                if (report.Status == "Verified" && report.Conflicts > 0) report.Status = "PartiallyApplied";
                report.Message = $"撤销已验证恢复 {report.Applied} 段；跳过 {report.Conflicts} 段；未验证 {report.Unverified} 段。";
                return report;
            }
        }

        internal static XElement Find(XElement page, string id) => page.Descendants(One + "OE").SingleOrDefault(e => (string)e.Attribute("objectID") == id);

        private static bool UntouchedPreserved(Dictionary<string, string> expected, XElement page)
        {
            foreach (var item in expected)
            {
                var oe = Find(page, item.Key);
                try { if (oe == null || AgentPageSnapshot.SemanticFormat(oe, page) != item.Value) return false; }
                catch (Exception) { return false; }
            }
            return true;
        }

        private static bool ContentPreserved(XElement expected, XElement actual)
        {
            if (Topology(expected) != Topology(actual)) return false;
            var before = expected.Descendants(One + "OE").ToList();
            var after = actual.Descendants(One + "OE").ToList();
            if (before.Count != after.Count) return false;
            for (var i = 0; i < before.Count; i++)
            {
                if (StableIdentity(before[i]) != StableIdentity(after[i])) return false;
                if (before[i].Ancestors(One + "OE").Count() != after[i].Ancestors(One + "OE").Count()) return false;
                if (before[i].Elements(One + "T").Any())
                {
                    try { if (new AgentRichText(before[i]).Signature(expected, false) != new AgentRichText(after[i]).Signature(actual, false)) return false; }
                    catch (Exception) { if (!before[i].Elements(One + "T").Select(t => t.Value).SequenceEqual(after[i].Elements(One + "T").Select(t => t.Value))) return false; }
                }
            }
            // 二进制对象的内容引用和表格布局不能消失。
            foreach (var name in new[] { "Image", "InkDrawing", "InsertedFile", "MediaFile", "FutureObject", "Tag", "List" })
            {
                var left = expected.Descendants(One + name).Select(StableXml).ToArray();
                var right = actual.Descendants(One + name).Select(StableXml).ToArray();
                if (!left.SequenceEqual(right)) return false;
            }
            if (!expected.Descendants(One + "Columns").Select(StableColumns).SequenceEqual(actual.Descendants(One + "Columns").Select(StableColumns))) return false;
            foreach (var name in new[] { "Table", "Cell" })
                if (!expected.Descendants(One + name).Select(TableAppearance).SequenceEqual(actual.Descendants(One + name).Select(TableAppearance))) return false;
            return true;
        }
        private static string StableColumns(XElement columns)
        {
            var copy = new XElement(columns);
            foreach (var column in copy.Elements(One + "Column"))
            {
                // Office 会随字体重算未锁定列的宽度。保留其自动宽度语义，不替用户锁列。
                var locked = (string)column.Attribute("isLocked");
                if (locked == "true" || locked == "1") column.SetAttributeValue("isLocked", "true");
                else { column.Attribute("width")?.Remove(); column.SetAttributeValue("isLocked", "false"); }
            }
            return StableXml(copy);
        }
        private static string TableAppearance(XElement element) => string.Join(";", new[] { "bordersVisible", "hasHeaderRow", "shadingColor", "alignment" }
            .Select(name => name + "=" + ((string)element.Attribute(name) ?? "").ToLowerInvariant()));
        private static string StableIdentity(XElement element)
        {
            // Office 回存表格时会重新生成无正文的外层 OE，内部 Table/Row/Cell/OE 的 ID 保持不变。
            // 只对纯表格包装节点使用内部 Table 的身份；普通段落仍严格检查 objectID。
            if (element.Name == One + "OE" && !element.Elements(One + "T").Any() && element.Elements().Count() == 1 && element.Element(One + "Table") != null)
                return "table-wrapper:" + (string)element.Element(One + "Table").Attribute("objectID");
            return (string)element.Attribute("objectID") ?? "";
        }
        private static string Topology(XElement page)
        {
            var names = new HashSet<string>(new[] { "Title", "Outline", "OEChildren", "OE", "Table", "Row", "Cell", "Image", "InkDrawing", "InsertedFile", "MediaFile" });
            return string.Join("\n", page.Descendants().Where(e => names.Contains(e.Name.LocalName)).Select(e =>
                string.Join("/", e.AncestorsAndSelf().Where(p => p != page).Reverse().Select(p =>
                    p.Name.LocalName + "#" + StableIdentity(p) + "@" + p.ElementsBeforeSelf(p.Name).Count()))));
        }
        private static string StableXml(XElement e)
        {
            var copy = new XElement(e);
            foreach (var a in copy.DescendantsAndSelf().Attributes().Where(a => a.Name.LocalName == "selected" || a.Name.LocalName == "lastModifiedTime").ToList()) a.Remove();
            foreach (var node in copy.DescendantNodes().OfType<XText>().Where(t => t.Parent.HasElements && string.IsNullOrWhiteSpace(t.Value)).ToList()) node.Remove();
            foreach (var element in copy.DescendantsAndSelf()) element.ReplaceAttributes(element.Attributes().OrderBy(a => a.Name.ToString(), StringComparer.Ordinal).ToArray());
            return copy.ToString(SaveOptions.DisableFormatting);
        }
    }
}
