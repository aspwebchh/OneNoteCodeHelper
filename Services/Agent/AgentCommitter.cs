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
        /// <summary>这段写入时修正的文字；撤销时文字也一起还原。</summary>
        internal List<string> TextFixes = new List<string>();
    }

    internal sealed class AgentReport
    {
        internal string Status;
        internal string Message;
        internal int Applied;
        internal int Conflicts;
        internal int Unverified;
        internal int Protected;
        /// <summary>已核验的代码框：执行时是新建的代码框数，撤销时是换回原段落的代码框数。</summary>
        internal int CodeBlocks;
        internal readonly List<AgentUndoItem> Undo = new List<AgentUndoItem>();
        internal readonly List<AgentCodeUndoItem> CodeUndo = new List<AgentCodeUndoItem>();
        internal readonly List<string> ConflictIds = new List<string>();
        /// <summary>已核验写入的文字修正，每项形如「原文」→「改后」。含笔记正文，只在窗口里显示，不写日志。</summary>
        internal readonly List<string> TextFixes = new List<string>();
        internal bool CanUndo => Undo.Count + CodeUndo.Count > 0;
        internal object ToToolResult() => new { status = Status, applied = Applied, text_fixes = TextFixes.Count, code_blocks = CodeBlocks, skipped_conflict = ConflictIds,
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
                    var report = new AgentReport { Protected = snapshot.Blocks.Count(b => !b.Editable && b.Conversion == null) };
                    // 写入前页面上已有的对象。OneNote 给新代码框、还原段落分配的 ID 都不在其中，核验时按新建对象比对。
                    var known = new HashSet<string>(page.Descendants().Attributes("objectID").Select(a => a.Value));
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
                        // 只有 fix_text 排过修正的段落可以改文字，而且只能改成草稿里的样子。
                        var expectedContent = new AgentRichText(block.Draft).Signature(page, false);
                        if (block.TextFixes.Count == 0 && expectedContent != new AgentRichText(target).Signature(page, false))
                            throw new AiException("格式修改改变了正文或链接，已阻止写入。");
                        AgentPageSnapshot.CopyFormat(block.Draft, target);
                        var styleId = (string)block.Draft.Attribute("quickStyleIndex");
                        var definition = styleId == null ? null : snapshot.DraftStyles.Elements(One + "QuickStyleDef").FirstOrDefault(d => (string)d.Attribute("index") == styleId);
                        if (definition != null) target.SetAttributeValue("quickStyleIndex", ParagraphStyles.EnsureDefinition(page, definition));
                        if (new AgentRichText(target).Signature(page, false) != expectedContent)
                            throw new AiException("写入的正文与草稿不一致，已阻止写入。");
                        planned.Add((block, before, new XElement(target)));
                        untouched.Remove(block.ObjectId);
                        containers.Add(target.Ancestors().First(e => e.Parent == page));
                    }
                    var codes = new List<(XElement Box, AgentCodeUndoItem Undo)>();
                    foreach (var conversion in snapshot.CodeConversions)
                    {
                        // 源段落有任何变化（改字、改格式、移动、插入下级段落）都整个跳过，不在别人的内容上重放。
                        var ids = conversion.Blocks.Select(b => b.ObjectId).ToList();
                        CodeSelection selection = null;
                        if (conversion.Blocks.All(b => fingerprints.TryGetValue(b.ObjectId, out var f) && f == b.Fingerprint))
                            try { selection = AgentCode.Select(page, ids); } catch (AiException) { }
                        if (selection == null || selection.Code != conversion.Code)
                        { report.ConflictIds.AddRange(conversion.Blocks.Select(b => b.Id)); continue; }
                        var undo = AgentCode.CaptureOriginals(selection, page);
                        containers.Add(selection.Outline);
                        codes.Add((selection.ReplaceWith(new XElement(conversion.Table)), undo));
                        foreach (var id in ids) untouched.Remove(id);
                    }
                    var restores = new List<List<XElement>>();
                    foreach (var item in snapshot.CodeRestores)
                    {
                        var box = AgentCode.FindCodeBox(page, item.TableId);
                        if (box == null || AgentCode.Fingerprint(box) != item.Fingerprint) { report.ConflictIds.Add(item.TableId); continue; }
                        foreach (var id in box.Descendants(One + "OE").Select(e => (string)e.Attribute("objectID")).Where(id => id != null).ToList()) untouched.Remove(id);
                        containers.Add(box.Ancestors().First(e => e.Parent == page));
                        restores.Add(AgentCode.Restore(box, item, page));
                    }
                    report.Conflicts = report.ConflictIds.Count;
                    if (planned.Count == 0 && codes.Count == 0 && restores.Count == 0)
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
                        report.Unverified = planned.Count + codes.Count + restores.Count;
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
                            report.TextFixes.AddRange(item.Block.TextFixes);
                            report.Undo.Add(new AgentUndoItem { ObjectId = item.Block.ObjectId, Before = item.Before,
                                AfterFingerprint = AgentPageSnapshot.Fingerprint(written, actual), TextFixes = item.Block.TextFixes.ToList() });
                        }
                        catch (Exception) { report.Unverified++; }
                    }
                    // 原页面所有文字、链接、段落顺序必须保留，包括没有交给模型的对象。
                    if (!ContentPreserved(page, actual, known) || !UntouchedPreserved(untouched, actual))
                    {
                        report.Status = "CommitOutcomeUnknown";
                        report.Message = "写入后页面结构或内容与预期不一致，请检查目标页面。";
                        return report;
                    }
                    // 内容核验保证两边的段落一一对应，按位置找到 OneNote 新建的代码框和还原段落。
                    var expectedLines = page.Descendants(One + "OE").ToList();
                    var actualLines = actual.Descendants(One + "OE").ToList();
                    foreach (var code in codes)
                    {
                        var written = actualLines[expectedLines.IndexOf(code.Box)];
                        var tableId = (string)written.Element(One + "Table")?.Attribute("objectID");
                        if (string.IsNullOrEmpty(tableId)) { report.Unverified++; continue; }
                        code.Undo.TableId = tableId;
                        code.Undo.Fingerprint = AgentCode.Fingerprint(written);
                        report.CodeBlocks++;
                        report.CodeUndo.Add(code.Undo);
                    }
                    foreach (var restored in restores)
                    {
                        var same = restored.SelectMany(r => r.DescendantsAndSelf(One + "OE")).Where(e => e.Elements(One + "T").Any()).All(e =>
                        {
                            try { return AgentPageSnapshot.SemanticFormat(e, page) == AgentPageSnapshot.SemanticFormat(actualLines[expectedLines.IndexOf(e)], actual); }
                            catch (Exception) { return false; }
                        });
                        if (same) report.CodeBlocks++; else report.Unverified++;
                    }
                    report.Status = report.Unverified > 0 || report.Conflicts > 0 ? "PartiallyApplied" : "Verified";
                    report.Message = $"已验证修改 {report.Applied} 段；" + (report.TextFixes.Count > 0 ? $"修正文字 {report.TextFixes.Count} 处；" : "") +
                        (codes.Count > 0 ? $"高亮代码 {report.CodeBlocks} 处；" : "") +
                        $"冲突跳过 {report.Conflicts} 段；未验证 {report.Unverified} 段；保护 {report.Protected} 段。";
                    if (uncertain && report.Applied + report.CodeBlocks == 0) report.Message = "写回未得到确认，请检查页面。" + report.Message;
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
                snapshot.CodeRestores.AddRange(previous.CodeUndo);
                var skipped = new List<string>();
                foreach (var item in previous.Undo)
                {
                    var block = snapshot.Blocks.FirstOrDefault(b => b.ObjectId == item.ObjectId);
                    if (block == null || !block.Editable || block.Fingerprint != item.AfterFingerprint)
                    { skipped.Add(item.ObjectId); continue; }
                    AgentPageSnapshot.CopyFormat(item.Before, block.Draft);
                    // 把修正过的文字改回去，提交时按改文字的段落核验。
                    block.TextFixes.AddRange(item.TextFixes);
                }
                var report = Commit(snapshot, cancellation);
                report.Conflicts += skipped.Count;
                report.ConflictIds.AddRange(skipped);
                if (report.Status == "Verified" && report.Conflicts > 0) report.Status = "PartiallyApplied";
                report.Message = $"撤销已验证恢复 {report.Applied} 段；" + (report.TextFixes.Count > 0 ? $"还原文字 {report.TextFixes.Count} 处；" : "") +
                    (previous.CodeUndo.Count > 0 ? $"恢复代码 {report.CodeBlocks} 处；" : "") +
                    $"跳过 {report.Conflicts} 处；未验证 {report.Unverified} 处。";
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

        private static bool ContentPreserved(XElement expected, XElement actual, ISet<string> known)
        {
            if (Topology(expected, known) != Topology(actual, known)) return false;
            var before = expected.Descendants(One + "OE").ToList();
            var after = actual.Descendants(One + "OE").ToList();
            if (before.Count != after.Count) return false;
            for (var i = 0; i < before.Count; i++)
            {
                if (StableIdentity(before[i], known) != StableIdentity(after[i], known)) return false;
                if (before[i].Ancestors(One + "OE").Count() != after[i].Ancestors(One + "OE").Count()) return false;
                // 新建的代码行和还原的段落只比文字：OneNote 会改写代码行的 span 和硬空格，还原段落的格式另行核验。
                if (before[i].Elements(One + "T").Any() && !known.Contains((string)before[i].Attribute("objectID") ?? ""))
                {
                    if (AgentCode.PlainText(before[i]) != AgentCode.PlainText(after[i])) return false;
                }
                else if (before[i].Elements(One + "T").Any())
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
        private static string StableIdentity(XElement element, ISet<string> known)
        {
            // Office 回存表格时会重新生成无正文的外层 OE，内部 Table/Row/Cell/OE 的 ID 保持不变。
            // 只对纯表格包装节点使用内部 Table 的身份；普通段落仍严格检查 objectID。
            if (element.Name == One + "OE" && !element.Elements(One + "T").Any() && element.Elements().Count() == 1 && element.Element(One + "Table") != null)
                return "table-wrapper:" + Known((string)element.Element(One + "Table").Attribute("objectID"), known);
            return Known((string)element.Attribute("objectID"), known);
        }
        /// <summary>写入前不存在的 ID 是这次新建的对象：期望页上它们还没有 ID，回读页上是 OneNote 新分配的，两边都记成空。</summary>
        private static string Known(string id, ISet<string> known) => id != null && known.Contains(id) ? id : "";
        private static string Topology(XElement page, ISet<string> known)
        {
            var names = new HashSet<string>(new[] { "Title", "Outline", "OEChildren", "OE", "Table", "Row", "Cell", "Image", "InkDrawing", "InsertedFile", "MediaFile" });
            return string.Join("\n", page.Descendants().Where(e => names.Contains(e.Name.LocalName)).Select(e =>
                string.Join("/", e.AncestorsAndSelf().Where(p => p != page).Reverse().Select(p =>
                    p.Name.LocalName + "#" + StableIdentity(p, known) + "@" + p.ElementsBeforeSelf(p.Name).Count()))));
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
