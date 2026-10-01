using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;

namespace OneNoteCodeHelper.Services.Agent
{
    internal interface IOneNotePageAccess
    {
        string GetPageContent(string pageId, PageInfo info);
        void UpdatePageContent(string xml, DateTime expectedLastModified);
        /// <summary>删除页面上的对象。只用来删掉合并后还留着一行占位空白的文本框（OneNote 通常在写入时已经删掉）。</summary>
        void DeletePageContent(string pageId, string objectId, DateTime expectedLastModified);
    }

    internal sealed class AgentUndoItem
    {
        internal string ObjectId;
        internal XElement Before;
        internal string AfterFingerprint;
        /// <summary>这段写入时修正的文字；撤销时文字也一起还原。</summary>
        internal List<string> TextFixes = new List<string>();
        /// <summary>这段写入时去掉的 Markdown 标记处数；撤销时文字也一起还原。</summary>
        internal int MarkdownMarks;
        /// <summary>Before 引用的 QuickStyleDef、TagDef 副本。OneNote 回存后会重新编号，撤销时按内容对应。</summary>
        internal List<XElement> Styles = new List<XElement>();
        internal List<XElement> Tags = new List<XElement>();
    }

    internal sealed class AgentReport
    {
        internal string Status;
        internal string Message;
        internal int Applied;
        internal int Conflicts;
        internal int Unverified;
        /// <summary>执行结束时尚未读取的段落数；撤销不沿用此计数。</summary>
        internal int UnreadCount;
        internal int Protected;
        /// <summary>已核验的代码框：执行时是新建的代码框数，撤销时是换回原段落的代码框数。</summary>
        internal int CodeBlocks;
        /// <summary>已核验的表格样式修改（撤销时是恢复的表格数）。</summary>
        internal int Tables;
        /// <summary>text_to_table 转成的表格（撤销时是换回段落的表格数）。</summary>
        internal int TextTables;
        /// <summary>结构改动：删掉的空行、移动、调整缩进和插入的段落数，合并掉的文本框数。</summary>
        internal int Removed, Moved, Indented, Inserted, Merged;
        internal int RemovedSoftLines, InsertedBlankLines;
        internal object[] SpacingSkipped = new object[0];
        /// <summary>整框写入的文本框，含合并删掉的（撤销时是整框恢复、重建的文本框数）。</summary>
        internal int Outlines;
        /// <summary>合并后没能删掉、留着一行空白的文本框。</summary>
        internal int Leftover;
        internal readonly List<AgentUndoItem> Undo = new List<AgentUndoItem>();
        internal readonly List<AgentCodeUndoItem> CodeUndo = new List<AgentCodeUndoItem>();
        internal readonly List<AgentTableUndoItem> TableUndo = new List<AgentTableUndoItem>();
        internal readonly List<AgentOutlineUndoItem> OutlineUndo = new List<AgentOutlineUndoItem>();
        internal readonly List<string> ConflictIds = new List<string>();
        /// <summary>已核验写入的文字修正，每项形如「原文」→「改后」。含笔记正文，只在窗口里显示，不写日志。</summary>
        internal readonly List<string> TextFixes = new List<string>();
        /// <summary>已核验去掉的 Markdown 标记处数，含整段删掉的围栏、分隔线（撤销时是还原的处数）。</summary>
        internal int MarkdownMarks;
        /// <summary>已核验写入、为保留下级格式而只设置外观的段落。</summary>
        internal readonly List<string> AppearanceOnly = new List<string>();
        internal bool CanUndo => Undo.Count + CodeUndo.Count + TableUndo.Count + OutlineUndo.Count > 0;
        /// <summary>撤销之后不再把撤销的逆操作当作可撤销。</summary>
        internal void ClearUndo() { Undo.Clear(); CodeUndo.Clear(); TableUndo.Clear(); OutlineUndo.Clear(); }
        internal object ToToolResult() => new { status = Status, applied = Applied, text_fixes = TextFixes.Count, markdown_marks = MarkdownMarks, code_blocks = CodeBlocks, tables = Tables, text_tables = TextTables,
            removed = Removed, moved = Moved, indented = Indented, inserted = Inserted, merged = Merged,
            removed_soft_lines = RemovedSoftLines, inserted_blank_lines = InsertedBlankLines, spacing_skipped = SpacingSkipped,
            skipped_conflict = ConflictIds, unverified = Unverified, unread_count = UnreadCount, protected_count = Protected, appearance_only = AppearanceOnly, message = Message };
        /// <summary>结果消息里的结构改动部分。</summary>
        internal string LayoutSummary => (Removed > 0 ? $"删除空行 {Removed} 行；" : "") +
            (RemovedSoftLines > 0 ? $"删除段内空行 {RemovedSoftLines} 行；" : "") + (InsertedBlankLines > 0 ? $"补空行 {InsertedBlankLines} 行；" : "") + (Moved > 0 ? $"移动 {Moved} 段；" : "") +
            (Indented > 0 ? $"调整缩进 {Indented} 段；" : "") + (Inserted > 0 ? $"插入 {Inserted} 段；" : "") + (Merged > 0 ? $"合并文本框 {Merged} 个；" : "") +
            (Leftover > 0 ? $"{Leftover} 个文本框合并后没能删掉，留下一行空白；" : "");
    }

    internal sealed class AgentCommitter
    {
        private readonly IOneNotePageAccess _api;
        private static XNamespace One => OneNoteApi.One;
        internal AgentCommitter(IOneNotePageAccess api) { _api = api; }

        // 只输出对象标识和固定原因。即使页面 XML 的标识异常，也不把任意字符串写进日志。
        internal static string VerificationDiagnostic(string kind, string id, string reason)
        {
            string Identifier(string value) => !string.IsNullOrEmpty(value) && value.Length <= 128 &&
                value.All(c => c < 128 && (char.IsLetterOrDigit(c) || "{}-_.".Contains(c))) ? value : "-";
            return $"Agent 核验失败：kind={Identifier(kind)} id={Identifier(id)} reason={Identifier(reason)}。";
        }

        private static void NoteUnverified(AgentReport report, string kind, string id, string reason)
        {
            report.Unverified++;
            AddInLog.Info(VerificationDiagnostic(kind, id, reason));
        }

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
                    foreach (var group in page.Descendants(One + "OE").Where(e => e.Attribute("objectID") != null).GroupBy(e => (string)e.Attribute("objectID")))
                    {
                        var lines = group.Where(e => e.Elements(One + "T").Any()).ToList();
                        if (lines.Count == 0) continue;
                        foreach (var oe in lines) fingerprints[group.Key] = AgentPageSnapshot.Fingerprint(oe, page);
                        try { untouched[group.Key] = UntouchedFormat(group, page); }
                        catch (Exception) { /* 不支持的 HTML 由内容不变量检查保留。 */ }
                    }
                    var planned = new List<(AgentBlock Block, XElement Before, XElement Desired, XElement Target)>();
                    var containers = new HashSet<XElement>();
                    var tagDefinitions = page.Elements(One + "TagDef").Count();
                    // 有结构改动的文本框整框替换，跨框移动、合并连起来的文本框整组替换。处理期间组里有文本框被改过就整组不替换：
                    // 结构改动按冲突跳过，框里其余改动照常逐项提交。
                    // 写入前的整页用来判断对象原来在哪个文本框、原来有没有列表；图片数据只在要跨框搬图片时才读。
                    var original = new XElement(page);
                    var homes = AgentLayout.Homes(original);
                    XElement binary = null;
                    XElement Binary() => binary ?? (binary = AgentPageSnapshot.ParsePage(_api.GetPageContent(snapshot.PageId, PageInfo.piBinaryData)));
                    var formatted = new HashSet<string>();
                    var edits = new List<OutlineEdit>();
                    var draft = snapshot.LayoutChanges.Count > 0 ? snapshot.CreateDraftPage() : null;
                    foreach (var group in snapshot.LayoutGroups)
                    {
                        if (!group.All(id => Outline(page, id) is XElement current && Outline(snapshot.Page, id) is XElement before &&
                            AgentLayout.OutlineFingerprint(current, page) == AgentLayout.OutlineFingerprint(before, snapshot.Page)))
                        {
                            report.ConflictIds.AddRange(snapshot.LayoutChanges.Where(c => group.Contains(c.OutlineId))
                                .SelectMany(c => c.Ids.Length > 0 ? c.Ids : new[] { c.From ?? c.OutlineId }).Distinct());
                            continue;
                        }
                        var key = Guid.NewGuid().ToString("N");
                        foreach (var id in group)
                        {
                            var current = Outline(page, id);
                            var edit = Outline(draft, id) == null ? DeleteOutline(page, current) : ReplaceOutline(snapshot, draft, page, original, homes, current, formatted, Binary);
                            edit.Group = key;
                            // 撤销时图片要放回别的文本框（或重建的文本框），先记下图片数据。
                            if (group.Count > 1) AgentLayout.FillImageData(edit.Before, Binary);
                            edits.Add(edit);
                        }
                    }
                    // 撤销：文本框写入后没被改过、合并删掉的文本框仍不在页面上，才整组换回。
                    foreach (var group in snapshot.OutlineRestores.GroupBy(i => i.Group ?? i.OutlineId))
                    {
                        if (!group.All(item => item.Deleted ? Outline(page, item.OutlineId) == null
                            : Outline(page, item.OutlineId) is XElement current && AgentLayout.OutlineFingerprint(current, page) == item.AfterFingerprint))
                        { report.ConflictIds.AddRange(group.Select(i => i.OutlineId)); continue; }
                        foreach (var item in group) edits.Add(RestoreOutline(item, page, original, homes, untouched, formatted));
                    }
                    var replaced = new HashSet<string>(edits.Select(e => e.Id));
                    foreach (var edit in edits) containers.Add(edit.Written);
                    // 删掉、转换、改过格式、跨框重建的段落不再按「未指定段落」核对；只是移动、调整缩进的段落仍要求格式不变。
                    var present = new HashSet<string>(page.Descendants().Attributes("objectID").Select(a => a.Value));
                    foreach (var edit in edits.Where(e => !e.Restore))
                        foreach (var oid in edit.Before.Descendants(One + "OE").Select(e => (string)e.Attribute("objectID")).Where(x => x != null))
                            if (!present.Contains(oid) || formatted.Contains(oid)) untouched.Remove(oid);
                    foreach (var block in snapshot.Blocks.Where(b => b.Changed && !replaced.Contains(b.ContainerId)))
                    {
                        var target = Find(page, block.ObjectId);
                        if (target == null || !fingerprints.TryGetValue(block.ObjectId, out var fingerprint) || fingerprint != block.Fingerprint)
                        { report.ConflictIds.Add(block.Id); continue; }
                        // 新恢复的连续编号依赖前一项；前项冲突时后续项也跳过，避免少了起点却接到上一组编号。
                        var number = block.Draft.Element(One + "List")?.Element(One + "Number");
                        if (block.MarkdownList?.Number != null && AgentMarks.ListKind(block.Original) != "number" &&
                            number != null && number.Attribute("restartNumberingAt") == null)
                        {
                            var priorNode = AgentLayout.Find(snapshot.Layout, block.Id)?.ElementsBeforeSelf(One + "OE").LastOrDefault();
                            var prior = snapshot.Blocks.FirstOrDefault(b => b.Id == (priorNode == null ? null : AgentLayout.KeyOf(priorNode)));
                            if (prior != null && (report.ConflictIds.Contains(prior.Id) ||
                                !fingerprints.TryGetValue(prior.ObjectId, out var priorFingerprint) || priorFingerprint != prior.Fingerprint))
                            { report.ConflictIds.Add(block.Id); continue; }
                        }
                        var before = new XElement(target);
                        // 只有 fix_text、strip_markdown 改过文字的段落可以改文字，而且只能改成草稿里的样子。
                        var expectedContent = new AgentRichText(block.Draft).Signature(page, false);
                        if (!block.TextEdited && expectedContent != new AgentRichText(target).Signature(page, false))
                            throw new AiException("格式修改改变了正文或链接，已阻止写入。");
                        var hadList = target.Element(One + "List") != null;
                        AgentPageSnapshot.CopyFormat(block.Draft, target);
                        // OneNote 不会因为 OE 里少了 List 就去掉列表（本机实测），只能让它把这一段当新段落重建。
                        // 下级段落保持原 ID，这一段得到新的 objectID，回读时按位置对应。
                        if (hadList && target.Element(One + "List") == null) AgentCode.StripIdentity(target);
                        var styleId = (string)block.Draft.Attribute("quickStyleIndex");
                        var definition = styleId == null ? null : snapshot.DraftStyles.Elements(One + "QuickStyleDef").FirstOrDefault(d => (string)d.Attribute("index") == styleId);
                        if (definition != null) target.SetAttributeValue("quickStyleIndex", ParagraphStyles.EnsureDefinition(page, definition));
                        // 标记引用的 TagDef 同样按内容对应到重新读取的页面；CopyFormat 保持了标记的先后顺序。
                        var draftTags = block.Draft.Elements(One + "Tag").ToList();
                        var targetTags = target.Elements(One + "Tag").ToList();
                        for (var i = 0; i < draftTags.Count && i < targetTags.Count; i++)
                        {
                            var tagDefinition = AgentMarks.Definition(snapshot.DraftTags, draftTags[i]);
                            if (tagDefinition != null) targetTags[i].SetAttributeValue("index", AgentMarks.EnsureTagDefinition(page, tagDefinition));
                        }
                        if (new AgentRichText(target).Signature(page, false) != expectedContent)
                            throw new AiException("写入的正文与草稿不一致，已阻止写入。");
                        planned.Add((block, before, new XElement(target), target));
                        untouched.Remove(block.ObjectId);
                        containers.Add(target.Ancestors().First(e => e.Parent == page));
                    }
                    var codes = new List<(XElement Box, AgentCodeUndoItem Undo, AgentCodeConversion Conversion)>();
                    foreach (var conversion in snapshot.CodeConversions.Where(c => !replaced.Contains(c.Blocks[0].ContainerId)))
                    {
                        // 源段落有任何变化（改字、改格式、移动、插入下级段落）都整个跳过，不在别人的内容上重放。
                        var ids = conversion.Blocks.Select(b => b.ObjectId).ToList();
                        CodeSelection selection = null;
                        if (conversion.Blocks.All(b => fingerprints.TryGetValue(b.ObjectId, out var f) && f == b.Fingerprint))
                            try { selection = AgentCode.Select(page, ids); } catch (AiException) { }
                        if (selection == null || selection.Code != conversion.Code)
                        { report.ConflictIds.AddRange(conversion.Blocks.Select(b => b.Id)); continue; }
                        var undo = AgentCode.CaptureOriginals(selection, page);
                        undo.TextTable = conversion.TextTable;
                        undo.MarkdownMarks = conversion.MarkdownMarks;
                        undo.TextFixes = conversion.TextFixes.ToList();
                        containers.Add(selection.Outline);
                        codes.Add((selection.ReplaceWith(new XElement(conversion.Table)), undo, conversion));
                        foreach (var id in ids) untouched.Remove(id);
                    }
                    var restores = new List<(List<XElement> Lines, AgentCodeUndoItem Item)>();
                    foreach (var item in snapshot.CodeRestores)
                    {
                        var box = AgentCode.FindCodeBox(page, item.TableId);
                        // 同批逐段撤销可能先恢复祖先格式或改变前面的段落数，冲突必须对照本轮写入前的状态。
                        var beforeBox = AgentCode.FindCodeBox(original, item.TableId);
                        if (box == null || beforeBox == null || AgentCode.Fingerprint(beforeBox, original) != item.Fingerprint) { report.ConflictIds.Add(item.TableId); continue; }
                        foreach (var id in box.Descendants(One + "OE").Select(e => (string)e.Attribute("objectID")).Where(id => id != null).ToList()) untouched.Remove(id);
                        containers.Add(box.Ancestors().First(e => e.Parent == page));
                        restores.Add((AgentCode.Restore(box, item, page), item));
                    }
                    var tables = new List<(AgentTable Table, TableLook Before)>();
                    foreach (var table in snapshot.Tables.Where(t => t.Changed && !replaced.Contains(t.ContainerId)))
                    {
                        // 外观、行列或首行被改过就跳过，不在别人改过的表格上重放。单元格文字的修改不影响。
                        var target = AgentTable.Find(page, table.ObjectId);
                        if (target == null || AgentTable.TakeFingerprint(target) != table.Fingerprint) { report.ConflictIds.Add(table.Id); continue; }
                        tables.Add((table, TableLook.Read(target)));
                        table.Draft.Apply(target);
                        containers.Add(target.Ancestors().First(e => e.Parent == page));
                    }
                    report.Conflicts = report.ConflictIds.Count;
                    if (planned.Count == 0 && codes.Count == 0 && restores.Count == 0 && tables.Count == 0 && edits.Count == 0)
                    {
                        report.Status = "NoChange";
                        report.Message = report.Conflicts > 0 ? $"没有写入：{report.Conflicts} 段在处理期间发生变化。" : "没有需要写入的格式修改。";
                        return report;
                    }
                    var timestamp = AgentPageSnapshot.Modified(page);
                    if (!UntouchedPreserved(untouched, page)) throw new AiException("这组格式会影响未指定的嵌套段落，已阻止写入。请分别选择段落处理。");
                    // 只有新增了 TagDef 才连同 TagDef 一起提交；复用页面已有的定义时不碰它们。
                    var sendTags = page.Elements(One + "TagDef").Count() != tagDefinitions;
                    var xml = PageEditor.BuildPageChanges(snapshot.PageId, page.Elements().Where(e => containers.Contains(e) || e.Name == One + "QuickStyleDef" ||
                        (sendTags && e.Name == One + "TagDef")).Select(e => new XElement(e)).ToArray());
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
                        report.Unverified = planned.Count + codes.Count + restores.Count + tables.Count + edits.Count;
                        AddInLog.Info(VerificationDiagnostic("page", snapshot.PageId, "readback_failed"));
                        report.Message = "已尝试写入，但无法回读确认。请查看目标页面，不要立即重复执行。";
                        return report;
                    }
                    // 合并后空了的文本框写成一行空白，OneNote 收到后直接删掉它（本机实测）。万一还留着占位的空白，
                    // 先不算进核对，核验通过后再用 DeletePageContent 删。
                    var placeholders = new Dictionary<OutlineEdit, string>();
                    foreach (var edit in edits.Where(e => e.Delete))
                    {
                        edit.Written.Remove();
                        var left = Outline(actual, edit.Id);
                        var lines = left?.Descendants(One + "OE").ToList();
                        if (lines == null || lines.Count != 1 || AgentCode.PlainText(lines[0]).Length > 0) continue;
                        placeholders[edit] = AgentLayout.OutlineFingerprint(left, actual);
                        left.Remove();
                    }
                    // OneNote 按位置把重建的文本框排进页面 XML，核对前先把它挪到期望页面里的顺序。
                    AlignNewOutlines(page, actual, known);
                    foreach (var item in planned)
                    {
                        // 重建的段落没有原 ID，按在页面上的位置找；位置对不上的情况由下面的内容核验兜底。
                        var recreated = item.Target.Attribute("objectID") == null;
                        var written = recreated ? AtSamePosition(page, actual, item.Target) : Find(actual, item.Block.ObjectId);
                        try
                        {
                            var desired = recreated ? AgentPageSnapshot.SemanticFormat(item.Target, page) : DesiredSignature(item.Desired, page, item.Block.ObjectId);
                            if (written == null || string.IsNullOrEmpty((string)written.Attribute("objectID")))
                            { NoteUnverified(report, "paragraph", item.Block.ObjectId, "missing_object"); continue; }
                            if (AgentPageSnapshot.SemanticFormat(written, actual) != desired)
                            { NoteUnverified(report, "paragraph", item.Block.ObjectId, "semantic_format_mismatch"); continue; }
                            report.Applied++;
                            report.TextFixes.AddRange(item.Block.TextFixes);
                            report.MarkdownMarks += item.Block.MarkdownMarks;
                            if (item.Block.AppearanceOnly) report.AppearanceOnly.Add(item.Block.Id);
                            report.Undo.Add(new AgentUndoItem { ObjectId = (string)written.Attribute("objectID"), Before = item.Before,
                                AfterFingerprint = AgentPageSnapshot.Fingerprint(written, actual), TextFixes = item.Block.TextFixes.ToList(), MarkdownMarks = item.Block.MarkdownMarks,
                                Styles = AgentLayout.Styles(item.Before, original), Tags = AgentLayout.Tags(item.Before, original) });
                        }
                        catch (Exception) { NoteUnverified(report, "paragraph", item.Block.ObjectId, "format_unreadable"); }
                    }
                    // 原页面所有文字、链接、段落顺序必须保留，包括没有交给模型的对象。
                    formatted.UnionWith(planned.Select(p => p.Block.ObjectId));
                    // 只有本次专用工具补入、或删段内空行后剩下的单行空段落，允许 OneNote 回存为空 T。
                    // 按期望节点记录，身份、层级和位置仍由内容核验先行检查。
                    var spacingBlanks = new HashSet<XElement>(edits.SelectMany(e => e.BlankLines.Concat(e.SoftLines.Select(s => s.Node))).Where(IsNormalizableBlank));
                    if (!ContentPreserved(page, actual, known, formatted, spacingBlanks) || !UntouchedPreserved(untouched, actual))
                    {
                        report.Status = "CommitOutcomeUnknown";
                        AddInLog.Info(VerificationDiagnostic("page", snapshot.PageId, "content_or_untouched_mismatch"));
                        report.Message = "写入后页面结构或内容与预期不一致，请检查目标页面。";
                        return report;
                    }
                    // 内容核验保证两边的段落一一对应，按位置找到 OneNote 新建的代码框和还原段落。
                    var expectedLines = page.Descendants(One + "OE").ToList();
                    var actualLines = actual.Descendants(One + "OE").ToList();
                    foreach (var code in codes)
                    {
                        var written = actualLines[expectedLines.IndexOf(code.Box)];
                        if (!ConversionWritten(code.Box, written, code.Conversion, page, actual))
                        { NoteUnverified(report, code.Conversion.TextTable ? "table" : "code", code.Conversion.Blocks[0].ObjectId, "conversion_mismatch"); continue; }
                        code.Undo.TableId = (string)written.Element(One + "Table").Attribute("objectID");
                        code.Undo.Fingerprint = AgentCode.Fingerprint(written, actual);
                        if (code.Conversion.TextTable) report.TextTables++; else report.CodeBlocks++;
                        report.MarkdownMarks += code.Conversion.MarkdownMarks;
                        report.TextFixes.AddRange(code.Conversion.TextFixes);
                        report.CodeUndo.Add(code.Undo);
                    }
                    foreach (var restored in restores)
                    {
                        var same = restored.Lines.SelectMany(r => r.DescendantsAndSelf(One + "OE")).Where(e => e.Elements(One + "T").Any()).All(e =>
                        {
                            try { return AgentPageSnapshot.SemanticFormat(e, page) == AgentPageSnapshot.SemanticFormat(actualLines[expectedLines.IndexOf(e)], actual); }
                            catch (Exception) { return false; }
                        });
                        if (!same) NoteUnverified(report, "restore", restored.Item.TableId, "semantic_format_mismatch");
                        else
                        {
                            if (restored.Item.TextTable) report.TextTables++; else report.CodeBlocks++;
                            report.MarkdownMarks += restored.Item.MarkdownMarks;
                            report.TextFixes.AddRange(restored.Item.TextFixes);
                        }
                    }
                    foreach (var edit in edits) VerifyOutline(edit, page, actual, expectedLines, actualLines, report);
                    // 重建的段落里的图片（跨框移动、撤销重建）只带着数据写入，piBasic 看不出坏图，按二进制数据核对。
                    var images = page.Descendants(One + "Image").Where(i => i.Parent?.Name == One + "OE" && !known.Contains((string)i.Parent.Attribute("objectID") ?? "")).ToList();
                    if (images.Count > 0 && !ImagesWritten(snapshot.PageId, page, images, known))
                        foreach (var image in images) NoteUnverified(report, "image", (string)image.Parent?.Attribute("objectID"), "image_data_mismatch");
                    // 还留着占位空白的文本框：写入核验通过后再删掉，占位的空白之后被人改过就留着。
                    foreach (var edit in edits.Where(e => e.Delete))
                    {
                        var deleted = !placeholders.TryGetValue(edit, out var placeholder) || TryDelete(snapshot.PageId, edit.Id, placeholder);
                        if (!deleted) report.Leftover++;
                        report.OutlineUndo.Add(new AgentOutlineUndoItem { OutlineId = edit.Id, Before = edit.Before, Styles = edit.Styles, Tags = edit.Tags, Group = edit.Group,
                            Deleted = deleted, AfterFingerprint = deleted ? null : placeholder });
                    }
                    foreach (var (table, before) in tables)
                    {
                        var written = AgentTable.Find(actual, table.ObjectId);
                        if (written == null || !TableLook.Read(written).SameAs(table.Draft))
                        { NoteUnverified(report, "table", table.ObjectId, written == null ? "missing_object" : "table_style_mismatch"); continue; }
                        report.Tables++;
                        report.TableUndo.Add(new AgentTableUndoItem { ObjectId = table.ObjectId, Before = before, AfterFingerprint = AgentTable.TakeFingerprint(written) });
                    }
                    report.Status = report.Unverified > 0 || report.Conflicts > 0 || report.Leftover > 0 ? "PartiallyApplied" : "Verified";
                    var conversions = codes.Select(c => c.Conversion).Concat(edits.SelectMany(e => e.Boxes.Select(b => b.Conversion))).ToList();
                    report.Message = $"已验证修改 {report.Applied} 段；" + (report.TextFixes.Count > 0 ? $"修正文字 {report.TextFixes.Count} 处；" : "") +
                        (report.MarkdownMarks > 0 ? $"去除 Markdown 符号 {report.MarkdownMarks} 处；" : "") + report.LayoutSummary +
                        (conversions.Any(c => !c.TextTable) ? $"高亮代码 {report.CodeBlocks} 处；" : "") + (conversions.Any(c => c.TextTable) ? $"转换表格 {report.TextTables} 个；" : "") +
                        (tables.Count + edits.Sum(e => e.Tables.Count) > 0 ? $"表格样式 {report.Tables} 个；" : "") +
                        $"冲突跳过 {report.Conflicts} 处；未验证 {report.Unverified} 处；保护 {report.Protected} 段。";
                    if (report.AppearanceOnly.Count > 0) report.Message += $"\n为保留下级段落格式，有 {report.AppearanceOnly.Count} 段仅设置外观，保留原有标题层级。";
                    if (uncertain && report.Applied + report.CodeBlocks + report.Tables + report.TextTables + report.Outlines == 0) report.Message = "写回未得到确认，请检查页面。" + report.Message;
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
                snapshot.OutlineRestores.AddRange(previous.OutlineUndo);
                // 表格外观和写入后一致时换回原外观；指纹用写入后的，表格之后又被改过就按冲突跳过。
                var restoredTables = new HashSet<string>(previous.TableUndo.Select(t => t.ObjectId));
                snapshot.Tables.RemoveAll(t => restoredTables.Contains(t.ObjectId));
                foreach (var item in previous.TableUndo)
                {
                    var current = AgentTable.Find(snapshot.Page, item.ObjectId);
                    snapshot.Tables.Add(new AgentTable { Id = "t" + (snapshot.Tables.Count + 1), ObjectId = item.ObjectId, Fingerprint = item.AfterFingerprint,
                        Original = current == null ? null : TableLook.Read(current), Draft = item.Before });
                }
                var skipped = new List<string>();
                foreach (var item in previous.Undo)
                {
                    var block = snapshot.Blocks.FirstOrDefault(b => b.ObjectId == item.ObjectId);
                    // strip_markdown 清空的围栏、分隔线现在是受保护空段，只允许凭已核验的撤销记录恢复。
                    var emptiedMarkdown = block?.ProtectedReason == "empty" && item.MarkdownMarks > 0;
                    if (block == null || !(block.Editable || emptiedMarkdown) || block.Fingerprint != item.AfterFingerprint)
                    { skipped.Add(item.ObjectId); continue; }
                    if (emptiedMarkdown) block.ProtectedReason = null;
                    AgentPageSnapshot.CopyFormat(Renumbered(item, snapshot), block.Draft);
                    // 把修正过的文字改回去，提交时按改文字的段落核验。
                    block.TextFixes.AddRange(item.TextFixes);
                    block.MarkdownMarks = item.MarkdownMarks;
                }
                var report = Commit(snapshot, cancellation);
                report.Conflicts += skipped.Count;
                report.ConflictIds.AddRange(skipped);
                if (report.Status == "Verified" && report.Conflicts > 0) report.Status = "PartiallyApplied";
                report.Message = $"撤销已验证恢复 {report.Applied} 段；" + (report.TextFixes.Count > 0 ? $"还原文字 {report.TextFixes.Count} 处；" : "") +
                    (report.MarkdownMarks > 0 ? $"还原 Markdown 符号 {report.MarkdownMarks} 处；" : "") +
                    (previous.OutlineUndo.Count > 0 ? $"恢复文本框结构 {report.Outlines} 个；" : "") +
                    (previous.CodeUndo.Any(u => !u.TextTable) ? $"恢复代码 {report.CodeBlocks} 处；" : "") +
                    (previous.CodeUndo.Any(u => u.TextTable) ? $"表格换回段落 {report.TextTables} 个；" : "") +
                    (previous.TableUndo.Count > 0 ? $"恢复表格 {report.Tables} 个；" : "") +
                    $"跳过 {report.Conflicts} 处；未验证 {report.Unverified} 处。";
                return report;
            }
        }

        /// <summary>
        /// 撤销项里的段落改用当前页面的样式、标记编号。OneNote 回存后会给 QuickStyleDef、TagDef 重新编号，
        /// 照搬执行前的编号会指到别的定义（比如正文的编号成了标题，撤销后整段变粗）。当前页面没有的定义加进草稿，提交时再对应到页面。
        /// </summary>
        private static XElement Renumbered(AgentUndoItem item, AgentPageSnapshot snapshot)
        {
            var before = new XElement(item.Before);
            var styles = item.Styles.ToDictionary(d => (string)d.Attribute("index"), d => ParagraphStyles.EnsureDefinition(snapshot.DraftStyles, d));
            var tags = item.Tags.ToDictionary(d => (string)d.Attribute("index"), d => AgentMarks.EnsureTagDefinition(snapshot.DraftTags, d));
            foreach (var e in before.DescendantsAndSelf().Where(e => e.Attribute("quickStyleIndex") != null))
                if (styles.TryGetValue((string)e.Attribute("quickStyleIndex"), out var mapped)) e.SetAttributeValue("quickStyleIndex", mapped);
            foreach (var tag in before.Descendants(One + "Tag"))
                if (tags.TryGetValue((string)tag.Attribute("index"), out var mapped)) tag.SetAttributeValue("index", mapped);
            return before;
        }

        internal static XElement Find(XElement page, string id) => page.Descendants(One + "OE").SingleOrDefault(e => (string)e.Attribute("objectID") == id);
        private static XElement Outline(XElement page, string id) => page.Elements(One + "Outline").FirstOrDefault(o => (string)o.Attribute("objectID") == id);

        /// <summary>一个整框写入的文本框：写入前后的样子，以及框里要核验、计数的改动。</summary>
        private sealed class OutlineEdit
        {
            internal string Id;
            internal bool Restore;
            /// <summary>合并后空了的文本框：写成一行空白，OneNote 收到后直接删掉；还留着时核验通过后再删。</summary>
            internal bool Delete;
            internal string Group;
            /// <summary>写入前的文本框；撤销重建的文本框原来不在页面上，为 null。</summary>
            internal XElement Before;
            /// <summary>页面里替换上去的文本框，写入后按位置和回读页面对应。</summary>
            internal XElement Written;
            internal List<XElement> Styles = new List<XElement>();
            internal List<XElement> Tags = new List<XElement>();
            internal readonly List<(AgentBlock Block, XElement Node)> Changed = new List<(AgentBlock, XElement)>();
            internal readonly List<(XElement Node, int Lines)> SoftLines = new List<(XElement, int)>();
            internal readonly HashSet<XElement> BlankLines = new HashSet<XElement>();
            internal readonly List<(XElement Box, AgentCodeConversion Conversion)> Boxes = new List<(XElement, AgentCodeConversion)>();
            /// <summary>改了外观的表格和它在写入内容里的元素；跨框移过来的表格没有原 ID，按位置核对。</summary>
            internal readonly List<(AgentTable Table, XElement Target)> Tables = new List<(AgentTable, XElement)>();
            internal readonly List<AgentLayoutChange> Changes = new List<AgentLayoutChange>();
            internal int RestoredMarkdownMarks;
            internal List<string> RestoredTextFixes = new List<string>();
        }

        /// <summary>
        /// 用结构草稿替换整个文本框：套上格式草稿、代码框和表格转换、表格外观，样式和标记编号按内容对应到重新读取的页面。
        /// 改动按段落、表格现在所在的文本框取（跨框移动后不在原来的框里）。跨框移过来的对象去掉 ID，图片补上数据。
        /// </summary>
        private static OutlineEdit ReplaceOutline(AgentPageSnapshot snapshot, XElement draft, XElement page, XElement original, Dictionary<string, string> homes,
            XElement current, ISet<string> formatted, Func<XElement> binary)
        {
            var id = (string)current.Attribute("objectID");
            var edit = new OutlineEdit { Id = id, Before = new XElement(current), Styles = AgentLayout.Styles(current, page), Tags = AgentLayout.Tags(current, page) };
            var written = Outline(draft, id);
            foreach (var inserted in snapshot.LayoutChanges.Where(c => c.Kind == "inserted_blank").SelectMany(c => c.Ids).Distinct())
                if (AgentLayout.Find(written, inserted) is XElement blank) edit.BlankLines.Add(blank);
            var converted = new HashSet<string>(snapshot.CodeConversions.SelectMany(c => c.Blocks).Select(b => b.Id));
            foreach (var node in written.Descendants(One + "OE").Where(AgentCodeSpacing.HasTrim).Where(n => !converted.Contains(AgentLayout.KeyOf(n) ?? "")))
            {
                edit.SoftLines.Add((node, AgentCodeSpacing.TrimCount(node)));
                if ((string)node.Attribute("objectID") is string objectId) formatted.Add(objectId);
            }
            foreach (var block in snapshot.Blocks.Where(b => b.Changed))
            {
                var node = AgentLayout.Find(written, block.Id);
                if (node == null) continue;
                // fix_text、strip_markdown 的文字草稿，以及已记录的完整空白行删除，才允许正文变化。
                var unchangedText = new XElement(block.Original);
                if (!block.TextEdited && block.Conversion == null) AgentCodeSpacing.Apply(unchangedText, node);
                if (!block.TextEdited && new AgentRichText(node).Signature(draft, false) != new AgentRichText(unchangedText).Signature(snapshot.Page, false))
                    throw new AiException("格式修改改变了正文或链接，已阻止写入。");
                edit.Changed.Add((block, node));
                formatted.Add(block.ObjectId);
            }
            foreach (var conversion in snapshot.CodeConversions.Where(c => AgentLayout.Find(written, c.Blocks[0].Id) != null))
            {
                var selection = AgentCode.Select(draft, conversion.Blocks.Select(b => b.ObjectId).ToList());
                if (selection.Code != conversion.Code) throw new AiException("结构草稿与代码框或表格转换不一致，已阻止写入。");
                edit.Boxes.Add((selection.ReplaceWith(new XElement(conversion.Table)), conversion));
            }
            foreach (var table in snapshot.Tables.Where(t => t.Changed))
            {
                var target = AgentTable.Find(written, table.ObjectId);
                if (target == null) continue;
                table.Draft.Apply(target);
                edit.Tables.Add((table, target));
            }
            AgentLayout.RebuildDroppedLists(written, original);
            foreach (var e in written.DescendantsAndSelf().Where(e => e.Attribute("quickStyleIndex") != null))
            {
                var definition = snapshot.DraftStyles.Elements(One + "QuickStyleDef").FirstOrDefault(d => (string)d.Attribute("index") == (string)e.Attribute("quickStyleIndex"));
                if (definition != null) e.SetAttributeValue("quickStyleIndex", ParagraphStyles.EnsureDefinition(page, definition));
            }
            foreach (var tag in written.Descendants(One + "Tag"))
            {
                var definition = AgentMarks.Definition(snapshot.DraftTags, tag);
                if (definition != null) tag.SetAttributeValue("index", AgentMarks.EnsureTagDefinition(page, definition));
            }
            AgentLayout.Strip(written);
            // OneNote 把移到另一个文本框的对象一律当新对象建立，带着原 ID 也一样；图片只带 CallbackID 会建成坏图。
            var foreign = AgentLayout.StripForeign(written, id, homes);
            AgentLayout.FillImageData(written, binary, foreign);
            AgentLayout.PrepareImages(written);
            written.Remove();
            current.ReplaceWith(written);
            edit.Written = written;
            edit.Changes.AddRange(snapshot.LayoutChanges.Where(c => c.OutlineId == id));
            return edit;
        }

        /// <summary>
        /// 合并后空了的文本框：OneNote 不接受没有段落的 Outline，写成一行空白；OneNote 收到只剩一行空白的文本框会直接删掉它。
        /// 期望页面里不算这个文本框，写入后还留着时由 <see cref="TryDelete"/> 删。
        /// </summary>
        private static OutlineEdit DeleteOutline(XElement page, XElement current)
        {
            var edit = new OutlineEdit { Id = (string)current.Attribute("objectID"), Delete = true, Before = new XElement(current),
                Styles = AgentLayout.Styles(current, page), Tags = AgentLayout.Tags(current, page) };
            var written = new XElement(current);
            written.Elements(One + "OEChildren").Remove();
            written.Add(new XElement(One + "OEChildren", new XElement(One + "OE", new XElement(One + "T", new XCData("")))));
            current.ReplaceWith(written);
            edit.Written = written;
            return edit;
        }

        /// <summary>删掉合并后还留着占位空白的文本框。占位的文本框之后被改过就不删；页面在删之前变了会重新读取再试。</summary>
        private bool TryDelete(string pageId, string outlineId, string expected)
        {
            for (var attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    var page = AgentPageSnapshot.ParsePage(_api.GetPageContent(pageId, PageInfo.piBasic));
                    var outline = Outline(page, outlineId);
                    if (outline == null) return true;
                    if (AgentLayout.OutlineFingerprint(outline, page) != expected) return false;
                    _api.DeletePageContent(pageId, outlineId, AgentPageSnapshot.Modified(page));
                    return true;
                }
                catch (Exception) { /* 时间戳冲突或结果不确定：重新读取后再判断。 */ }
            }
            try { return Outline(AgentPageSnapshot.ParsePage(_api.GetPageContent(pageId, PageInfo.piBasic)), outlineId) == null; }
            catch (Exception) { return false; }
        }

        /// <summary>
        /// 整框撤销：换回写入前的文本框，合并删掉的文本框按原来的位置和内容重新建立（得到新的 ID）。
        /// 原来就在这个框里、现在还在的对象保留 ID；已经不在的（删掉的空行、转换前的段落、表格外层、跨框移动后重建的对象）去掉 ID 让 OneNote 重建，
        /// 其中的图片用执行时记下的数据。列表要去掉的段落同样重建；样式和标记定义按当前页面重新对应。插入的段落和新建的表格不在原样里，写回时一并删除。
        /// </summary>
        private static OutlineEdit RestoreOutline(AgentOutlineUndoItem item, XElement page, XElement original, Dictionary<string, string> homes,
            Dictionary<string, string> untouched, ISet<string> formatted)
        {
            var current = item.Deleted ? null : Outline(page, item.OutlineId);
            var edit = new OutlineEdit { Id = item.OutlineId, Restore = true, Before = current == null ? null : new XElement(current),
                RestoredMarkdownMarks = item.MarkdownMarks, RestoredTextFixes = item.TextFixes };
            var written = new XElement(item.Before);
            if (current == null) AgentCode.StripIdentity(written);
            AgentLayout.StripForeign(written, item.OutlineId, homes);
            AgentLayout.PrepareImages(written);
            AgentLayout.RebuildDroppedLists(written, original);
            var styles = item.Styles.ToDictionary(d => (string)d.Attribute("index"), d => ParagraphStyles.EnsureDefinition(page, d));
            var tags = item.Tags.ToDictionary(d => (string)d.Attribute("index"), d => AgentMarks.EnsureTagDefinition(page, d));
            foreach (var e in written.DescendantsAndSelf().Where(e => e.Attribute("quickStyleIndex") != null))
                if (styles.TryGetValue((string)e.Attribute("quickStyleIndex"), out var mapped)) e.SetAttributeValue("quickStyleIndex", mapped);
            foreach (var tag in written.Descendants(One + "Tag"))
                if (tags.TryGetValue((string)tag.Attribute("index"), out var mapped)) tag.SetAttributeValue("index", mapped);
            if (current != null) current.ReplaceWith(written);
            // 重建的文本框先排在最后；OneNote 会按位置排，核对前由 AlignNewOutlines 对齐。
            else if (page.Elements(One + "Outline").LastOrDefault() is XElement last) last.AddAfterSelf(written);
            else page.Add(written);
            edit.Written = written;
            if (edit.Before != null)
                foreach (var oid in edit.Before.Descendants(One + "OE").Select(e => (string)e.Attribute("objectID")).Where(x => x != null)) untouched.Remove(oid);
            formatted.UnionWith(written.Descendants(One + "OE").Select(e => (string)e.Attribute("objectID")).Where(x => x != null));
            return edit;
        }

        /// <summary>
        /// 整框写入的核验：框里每个文字段落按位置比较格式（期望一侧解析不了的段落已由内容核验比过原文）；
        /// 新代码框只比文字，已在内容核验里做过；新表格的单元格另比正文和链接。执行时记下整框撤销。
        /// </summary>
        private static void VerifyOutline(OutlineEdit edit, XElement page, XElement actual, List<XElement> expectedLines, List<XElement> actualLines, AgentReport report)
        {
            var marksBefore = report.MarkdownMarks;
            var fixesBefore = report.TextFixes.Count;
            var boxLines = new HashSet<XElement>(edit.Boxes.SelectMany(b => b.Box.Descendants(One + "OE")));
            var failed = new HashSet<XElement>();
            // 合并删掉的文本框不在期望页面里，没有要按位置核对的段落。
            foreach (var line in edit.Delete ? Enumerable.Empty<XElement>() : edit.Written.Descendants(One + "OE").Where(e => e.Elements(One + "T").Any() && !boxLines.Contains(e)))
            {
                string expected;
                var blankLine = edit.BlankLines.Contains(line) || edit.SoftLines.Any(s => s.Node == line) && string.IsNullOrWhiteSpace(new AgentRichText(line).Text);
                try { expected = blankLine ? AgentPageSnapshot.BlankFormat(line, page) : AgentPageSnapshot.SemanticFormat(line, page); } catch (Exception) { continue; }
                try
                {
                    var written = actualLines[expectedLines.IndexOf(line)];
                    if ((blankLine ? AgentPageSnapshot.BlankFormat(written, actual) : AgentPageSnapshot.SemanticFormat(written, actual)) != expected)
                    {
                        failed.Add(line);
                        AddInLog.Info(VerificationDiagnostic("outline_paragraph", (string)line.Attribute("objectID") ?? AgentLayout.KeyOf(line), "semantic_format_mismatch"));
                    }
                }
                catch (Exception)
                {
                    failed.Add(line);
                    AddInLog.Info(VerificationDiagnostic("outline_paragraph", (string)line.Attribute("objectID") ?? AgentLayout.KeyOf(line), "format_unreadable"));
                }
            }
            report.Unverified += failed.Count;
            foreach (var (block, node) in edit.Changed.Where(c => !failed.Contains(c.Node)))
            {
                report.Applied++;
                report.TextFixes.AddRange(block.TextFixes);
                report.MarkdownMarks += block.MarkdownMarks;
                if (block.AppearanceOnly) report.AppearanceOnly.Add(block.Id);
            }
            foreach (var (box, conversion) in edit.Boxes)
            {
                if (!ConversionWritten(box, actualLines[expectedLines.IndexOf(box)], conversion, page, actual))
                { NoteUnverified(report, conversion.TextTable ? "table" : "code", conversion.Blocks[0].ObjectId, "conversion_mismatch"); continue; }
                if (conversion.TextTable) report.TextTables++; else report.CodeBlocks++;
                report.MarkdownMarks += conversion.MarkdownMarks;
                report.TextFixes.AddRange(conversion.TextFixes);
            }
            foreach (var (table, target) in edit.Tables)
            {
                var index = expectedLines.IndexOf(target.Parent);
                var written = index < 0 ? null : actualLines[index].Element(One + "Table");
                if (written == null || !TableLook.Read(written).SameAs(table.Draft))
                { NoteUnverified(report, "table", table.ObjectId, written == null ? "missing_object" : "table_style_mismatch"); continue; }
                report.Tables++;
            }
            string[] Ids(string kind) => edit.Changes.Where(c => c.Kind == kind).SelectMany(c => c.Ids).Distinct().ToArray();
            report.Removed += Ids("removed").Length;
            report.RemovedSoftLines += edit.SoftLines.Where(s => !failed.Contains(s.Node)).Sum(s => s.Lines);
            report.InsertedBlankLines += edit.BlankLines.Count(b => !failed.Contains(b));
            report.MarkdownMarks += Ids("markdown").Length;
            report.Moved += Ids("moved").Length;
            report.Indented += Ids("indented").Length;
            report.Inserted += Ids("inserted").Length;
            report.Merged += edit.Changes.Count(c => c.Kind == "merged");
            if (edit.Restore && failed.Count == 0)
            {
                report.MarkdownMarks += edit.RestoredMarkdownMarks;
                report.TextFixes.AddRange(edit.RestoredTextFixes);
            }
            report.Outlines++;
            // 合并删掉的文本框在删除之后才记撤销。
            var after = edit.Restore || edit.Delete ? null : Outline(actual, edit.Id);
            if (after != null)
                report.OutlineUndo.Add(new AgentOutlineUndoItem { OutlineId = edit.Id, Before = edit.Before, Styles = edit.Styles, Tags = edit.Tags, Group = edit.Group,
                    AfterFingerprint = AgentLayout.OutlineFingerprint(after, actual), MarkdownMarks = report.MarkdownMarks - marksBefore,
                    TextFixes = report.TextFixes.Skip(fixesBefore).ToList() });
        }

        /// <summary>
        /// 撤销重建的文本框没有 ID，期望页面里排在最后；OneNote 按位置把它排进页面 XML。把回读页面里新出现的文本框挪到期望页面的顺序，
        /// 其他元素不动。数量对不上时不调整，交给内容核验报错。
        /// </summary>
        private static void AlignNewOutlines(XElement expected, XElement actual, ISet<string> known)
        {
            var fresh = expected.Elements(One + "Outline").Select(o => !known.Contains((string)o.Attribute("objectID") ?? "")).ToList();
            if (!fresh.Contains(true)) return;
            var outlines = actual.Elements(One + "Outline").ToList();
            var added = outlines.Where(o => !known.Contains((string)o.Attribute("objectID") ?? "")).ToList();
            var kept = outlines.Except(added).ToList();
            if (outlines.Count != fresh.Count || added.Count != fresh.Count(f => f)) return;
            var slots = outlines.Select(o => { var slot = new XElement("slot"); o.AddBeforeSelf(slot); return slot; }).ToList();
            foreach (var o in outlines) o.Remove();
            int a = 0, k = 0;
            for (var i = 0; i < slots.Count; i++) slots[i].ReplaceWith(fresh[i] ? added[a++] : kept[k++]);
        }

        /// <summary>重建的段落里的图片：回读二进制数据，数量和位置对得上，带着数据写入的图片数据不变。坏图在二进制读取里会整个消失。</summary>
        private bool ImagesWritten(string pageId, XElement expected, List<XElement> images, ISet<string> known)
        {
            try
            {
                var binary = AgentPageSnapshot.ParsePage(_api.GetPageContent(pageId, PageInfo.piBinaryData));
                AlignNewOutlines(expected, binary, known);
                var before = expected.Descendants(One + "Image").ToList();
                var after = binary.Descendants(One + "Image").ToList();
                if (before.Count != after.Count) return false;
                string Data(XElement image) => Regex.Replace((string)image.Element(One + "Data") ?? "", @"\s", "");
                return images.All(i => { var written = Data(after[before.IndexOf(i)]); return written.Length > 0 && (i.Element(One + "Data") == null || written == Data(i)); });
            }
            catch (Exception) { return false; }
        }

        /// <summary>转换出的代码框或表格已由 OneNote 建立；表格另外核对每个单元格的正文和链接。</summary>
        private static bool ConversionWritten(XElement box, XElement written, AgentCodeConversion conversion, XElement page, XElement actual)
        {
            if (string.IsNullOrEmpty((string)written.Element(One + "Table")?.Attribute("objectID"))) return false;
            if (!conversion.TextTable) return true;
            var left = box.Descendants(One + "OE").Where(e => e.Elements(One + "T").Any()).ToList();
            var right = written.Descendants(One + "OE").Where(e => e.Elements(One + "T").Any()).ToList();
            // 清理草稿转来的表格还要验证保留的行内格式；其他转换沿用原来的正文、链接核验。
            var styles = conversion.MarkdownMarks > 0;
            try { return left.Count == right.Count && left.Zip(right, (a, b) => new AgentRichText(a).Signature(page, styles) == new AgentRichText(b).Signature(actual, styles)).All(x => x); }
            catch (Exception) { return false; }
        }

        private static bool UntouchedPreserved(Dictionary<string, string> expected, XElement page)
        {
            foreach (var item in expected)
            {
                var paragraphs = page.Descendants(One + "OE").Where(e => (string)e.Attribute("objectID") == item.Key).ToList();
                try { if (paragraphs.Count == 0 || UntouchedFormat(paragraphs, page) != item.Value) return false; }
                catch (Exception) { return false; }
            }
            return true;
        }

        /// <summary>同一 objectID 的全部段落按页面顺序的语义格式。页面上可能有重复 ID，逐个核对，不能只取其中一个。</summary>
        private static string UntouchedFormat(IEnumerable<XElement> paragraphs, XElement page) =>
            string.Join("\n", paragraphs.Select(oe => AgentPageSnapshot.SemanticFormat(oe, page)));

        private static bool IsNormalizableBlank(XElement oe)
        {
            try { return new AgentRichText(oe).IsSingleBlank; }
            catch (Exception ex) when (ex is AiException || ex is System.Xml.XmlException || ex is ArgumentException) { return false; }
        }

        /// <param name="formatted">本次写入格式的段落。它们的列表和标记已按语义核验，这里不再逐字比 XML（OneNote 会补上字号、编号文字、时间）。</param>
        /// <param name="spacingBlanks">本次代码框间隔调整得到的单行空段落；只允许无链接、无换行的空白与空 T 等价，外观另行核验。</param>
        private static bool ContentPreserved(XElement expected, XElement actual, ISet<string> known, ISet<string> formatted, ISet<XElement> spacingBlanks)
        {
            if (Topology(expected, known) != Topology(actual, known)) return false;
            var before = expected.Descendants(One + "OE").ToList();
            var after = actual.Descendants(One + "OE").ToList();
            if (before.Count != after.Count) return false;
            for (var i = 0; i < before.Count; i++)
            {
                if (StableIdentity(before[i], known) != StableIdentity(after[i], known)) return false;
                if (before[i].Ancestors(One + "OE").Count() != after[i].Ancestors(One + "OE").Count()) return false;
                if (spacingBlanks.Contains(before[i]))
                {
                    if (!IsNormalizableBlank(after[i])) return false;
                }
                // 新建的代码行和还原的段落只比文字：OneNote 会改写代码行的 span 和硬空格，还原段落的格式另行核验。
                else if (before[i].Elements(One + "T").Any() && !known.Contains((string)before[i].Attribute("objectID") ?? ""))
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
                var left = expected.Descendants(One + name).Where(e => !IsFormattedMark(e, formatted, known)).Select(e => StableObject(e, expected, known)).ToArray();
                var right = actual.Descendants(One + name).Where(e => !IsFormattedMark(e, formatted, known)).Select(e => StableObject(e, actual, known)).ToArray();
                if (!left.SequenceEqual(right)) return false;
            }
            if (!expected.Descendants(One + "Columns").Select(StableColumns).SequenceEqual(actual.Descendants(One + "Columns").Select(StableColumns))) return false;
            foreach (var name in new[] { "Table", "Cell" })
                if (!expected.Descendants(One + name).Select(TableAppearance).SequenceEqual(actual.Descendants(One + name).Select(TableAppearance))) return false;
            return true;
        }
        /// <summary>本次写入格式的段落和新建（重建）段落上的标记、列表：已按语义核验，不逐字比 XML。</summary>
        private static bool IsFormattedMark(XElement e, ISet<string> formatted, ISet<string> known)
        {
            if (e.Name != One + "Tag" && e.Name != One + "List") return false;
            var id = (string)e.Parent?.Attribute("objectID") ?? "";
            return formatted.Contains(id) || !known.Contains(id);
        }

        /// <summary>写入前后页面的段落一一对应（由内容核验保证），按序号找回 OneNote 新建的段落。</summary>
        private static XElement AtSamePosition(XElement expected, XElement actual, XElement element)
        {
            var before = expected.Descendants(One + "OE").ToList();
            var after = actual.Descendants(One + "OE").ToList();
            var index = before.IndexOf(element);
            return index >= 0 && before.Count == after.Count ? after[index] : null;
        }

        /// <summary>
        /// 标记按引用的 TagDef 内容比较，不看编号：新增 TagDef 后 OneNote 重新编号也不算内容变化。
        /// 新建段落里的图片带着数据写入、回读只有新的 CallbackID，这里不比数据引用，另由 <see cref="ImagesWritten"/> 按二进制核对。
        /// </summary>
        private static string StableObject(XElement e, XElement page, ISet<string> known)
        {
            if (e.Name == One + "Image" && e.Parent?.Name == One + "OE" && !known.Contains((string)e.Parent.Attribute("objectID") ?? ""))
            {
                var image = new XElement(e);
                image.Elements(One + "CallbackID").Remove();
                image.Elements(One + "Data").Remove();
                return StableXml(image);
            }
            if (e.Name != One + "Tag") return StableXml(e);
            var tag = new XElement(e);
            tag.Attribute("index")?.Remove();
            var definition = AgentMarks.Definition(page, e);
            if (definition == null) return StableXml(tag) + "|missing";
            definition = new XElement(definition);
            definition.Attribute("index")?.Remove();
            return StableXml(tag) + "|" + StableXml(definition);
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
        /// <summary>OneNote 可能省略值为 false 的开关和空底色，按 <see cref="TableLook"/> 的规则归一后比较。</summary>
        private static string TableAppearance(XElement element) => string.Join(";",
            "bordersVisible=" + TableLook.Flag(element, "bordersVisible"), "hasHeaderRow=" + TableLook.Flag(element, "hasHeaderRow"),
            "shadingColor=" + TableLook.Shade((string)element.Attribute("shadingColor")), "alignment=" + ((string)element.Attribute("alignment") ?? "").ToLowerInvariant());
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
