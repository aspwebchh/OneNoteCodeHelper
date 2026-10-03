using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using OneNoteCodeHelper.Services.Agent;
using OneNoteCodeHelper.Services.Mcp;

internal static partial class Program
{
    private static XElement McpFidelityObject(XElement page) => page.Descendants(One + "OE").Single(o => o.Element(One + "Tag") != null);
    private static XElement McpFidelityDefinition(XElement page) => AgentMarks.Definition(page, McpFidelityObject(page).Element(One + "Tag"));
    private static string[] McpTagNameVariants() => new[] { "Custom tag", "CUSTOM TAG", " Custom tag", "Custom tag " };
    private static XElement McpNamedTagsPage(bool text)
    {
        var names = McpTagNameVariants();
        var paragraphs = Enumerable.Range(0, names.Length + 1).Select(i =>
        {
            var oe = new XElement(One + "OE", new XAttribute("objectID", "named-tag-" + i), Tag((i % names.Length).ToString()));
            if (text) oe.Add(new XElement(One + "T", new XCData("标记段落 " + i)));
            return oe;
        });
        var page = Page(paragraphs.ToArray());
        page.AddFirst(names.Select((name, i) => TagDef(i.ToString(), 3, name)));
        return page;
    }
    private static void CheckMcpTagNames(XElement page)
    {
        var names = McpTagNameVariants();
        var tags = page.Descendants(One + "Tag").ToList();
        Equal(names.Length, page.Elements(One + "TagDef").Count()); Equal(names.Length + 1, tags.Count);
        for (var i = 0; i < tags.Count; i++)
            Equal(names[i % names.Length], (string)AgentMarks.Definition(page, tags[i])?.Attribute("name"));
        Equal(names.Length, tags.Take(names.Length).Select(t => (string)t.Attribute("index")).Distinct().Count());
        Equal((string)tags[0].Attribute("index"), (string)tags.Last().Attribute("index"));
    }
    private static XElement McpFidelityPage(bool text)
    {
        var page = McpMarkedObjectPage("empty"); var oe = McpFidelityObject(page);
        oe.Element(One + "List").ReplaceWith(new XElement(One + "List", new XElement(One + "Bullet", new XAttribute("bullet", "2"))));
        oe.Element(One + "List").AddBeforeSelf(new XElement(One + "Meta", new XAttribute("name", "001"), new XAttribute("content", "001")));
        if (text) oe.Add(new XElement(One + "T", new XCData("待办正文")));
        var definition = McpFidelityDefinition(page);
        definition.SetAttributeValue("name", "Custom tag"); definition.SetAttributeValue("fontColor", "#123456"); definition.SetAttributeValue("highlightColor", "#ABCDEF");
        return page;
    }
    private static Dictionary<string, Action<XElement>> McpFidelityChanges() => new Dictionary<string, Action<XElement>>
    {
        ["tag name"] = p => McpFidelityDefinition(p).SetAttributeValue("name", "Changed tag"),
        ["tag name case"] = p => McpFidelityDefinition(p).SetAttributeValue("name", "CUSTOM TAG"),
        ["tag name leading space"] = p => McpFidelityDefinition(p).SetAttributeValue("name", " Custom tag"),
        ["tag name trailing space"] = p => McpFidelityDefinition(p).SetAttributeValue("name", "Custom tag "),
        ["tag text color"] = p => McpFidelityDefinition(p).SetAttributeValue("fontColor", "#FF0000"),
        ["tag highlight"] = p => McpFidelityDefinition(p).SetAttributeValue("highlightColor", "#FFFF00"),
        ["tag type"] = p => McpFidelityDefinition(p).SetAttributeValue("type", "1"),
        ["tag disabled"] = p => McpFidelityObject(p).Element(One + "Tag").SetAttributeValue("disabled", "true"),
        ["bullet symbol"] = p => McpFidelityObject(p).Element(One + "List").Element(One + "Bullet").SetAttributeValue("bullet", "1"),
        ["metadata name"] = p => McpFidelityObject(p).Element(One + "Meta").SetAttributeValue("name", "1"),
        ["metadata value"] = p => McpFidelityObject(p).Element(One + "Meta").SetAttributeValue("content", "1")
    };
    private static void TestMcpWriteVerification()
    {
        Test("shared tag mapping preserves exact names and reuses equivalent renumbered definitions", () =>
        {
            var page = new XElement(One + "Page");
            foreach (var name in McpTagNameVariants()) AgentMarks.EnsureTagDefinition(page, TagDef("99", 3, name));
            Equal(4, page.Elements(One + "TagDef").Count());
            foreach (var definition in page.Elements(One + "TagDef").ToList())
            {
                var equivalent = new XElement(definition); equivalent.SetAttributeValue("index", "99");
                equivalent.SetAttributeValue("fontColor", null); equivalent.SetAttributeValue("highlightColor", null);
                equivalent.ReplaceAttributes(equivalent.Attributes().Reverse().ToArray());
                Equal(AgentMarks.Signature(definition), AgentMarks.Signature(equivalent));
                Equal((string)definition.Attribute("index"), AgentMarks.EnsureTagDefinition(page, equivalent));
            }
            var colored = TagDef("99", 3, "Custom tag"); colored.SetAttributeValue("fontColor", "#ABCDEF");
            var mapped = AgentMarks.EnsureTagDefinition(page, colored);
            colored.SetAttributeValue("fontColor", "#abcdef");
            Equal(mapped, AgentMarks.EnsureTagDefinition(page, colored)); Equal(5, page.Elements(One + "TagDef").Count());
        });
        Test("MCP copy move and compensation preserve distinct names and shared tag references", () =>
        {
            foreach (var text in new[] { false, true }) foreach (var operation in new[] { "copy_page", "move_page" })
            {
                var api = new FakeWorkspace(); api.Add("page", McpNamedTagsPage(text));
                using (var s = WorkspaceService(api))
                {
                    var id = WorkspaceBegin(s); McpCall(s, operation, new { snapshot_id = id, page_id = "page", section_id = "s2" });
                    var result = McpFinish(s, id, 1); Equal("Verified", result["status"]);
                    var target = (string)result["page_id"]; CheckMcpTagNames(api.Pages[target].Page);
                    Equal(operation == "move_page", result["source_recycled"]);
                    var undo = McpCall(s, "undo_edit", new { snapshot_id = id }); Equal("Verified", undo["status"]);
                    True(!api.Pages.ContainsKey(target));
                    CheckMcpTagNames(api.Pages[operation == "move_page" ? (string)undo["restored_page_id"] : "page"].Page);
                    Equal(operation == "move_page" ? 2 : 1, api.Creates);
                    Equal(operation == "move_page" ? 2 : 1, api.Recycles);
                }
            }
        });
        Test("MCP source tag name edits during planning prevent copy and move writes", () =>
        {
            foreach (var text in new[] { false, true }) foreach (var operation in new[] { "copy_page", "move_page" }) foreach (var name in McpTagNameVariants().Skip(1))
            {
                var api = new FakeWorkspace(); api.Add("page", McpFidelityPage(text));
                using (var s = WorkspaceService(api))
                {
                    var id = WorkspaceBegin(s); McpCall(s, operation, new { snapshot_id = id, page_id = "page", section_id = "s2" });
                    McpFidelityDefinition(api.Pages["page"].Page).SetAttributeValue("name", name);
                    var edited = api.Pages["page"].Page.ToString(SaveOptions.DisableFormatting);
                    Equal("NoChange", McpFinish(s, id, 1)["status"]); Equal(0, api.Creates); Equal(0, api.Recycles);
                    Equal(0, api.Pages["page"].Writes); Equal(edited, api.Pages["page"].Page.ToString(SaveOptions.DisableFormatting));
                }
            }
        });
        Test("MCP move preserves source tag name edits made while writing the target", () =>
        {
            foreach (var text in new[] { false, true }) foreach (var name in McpTagNameVariants().Skip(1))
            {
                var api = new FakeWorkspace(); api.Add("page", McpFidelityPage(text));
                using (var s = WorkspaceService(api))
                {
                    var id = WorkspaceBegin(s); McpCall(s, "move_page", new { snapshot_id = id, page_id = "page", section_id = "s2" });
                    api.AfterUpdate = () => McpFidelityDefinition(api.Pages["page"].Page).SetAttributeValue("name", name);
                    var result = McpFinish(s, id, 1); Equal("PartiallyApplied", result["status"]); Equal(false, result["source_recycled"]);
                    var target = (string)result["page_id"]; Equal(1, api.Creates); Equal(0, api.Recycles);
                    Equal("Custom tag", (string)McpFidelityDefinition(api.Pages[target].Page).Attribute("name"));
                    Equal("Verified", McpCall(s, "undo_edit", new { snapshot_id = id })["status"]);
                    True(!api.Pages.ContainsKey(target)); Equal(name, (string)McpFidelityDefinition(api.Pages["page"].Page).Attribute("name"));
                }
            }
        });
        Test("MCP undo preserves target tag name edits after copy and move", () =>
        {
            foreach (var text in new[] { false, true }) foreach (var operation in new[] { "copy_page", "move_page" }) foreach (var name in McpTagNameVariants().Skip(1))
            {
                var api = new FakeWorkspace(); api.Add("page", McpFidelityPage(text));
                using (var s = WorkspaceService(api))
                {
                    var id = WorkspaceBegin(s); McpCall(s, operation, new { snapshot_id = id, page_id = "page", section_id = "s2" });
                    var result = McpFinish(s, id, 1); Equal("Verified", result["status"]); var target = (string)result["page_id"];
                    McpFidelityDefinition(api.Pages[target].Page).SetAttributeValue("name", name);
                    var edited = api.Pages[target].Page.ToString(SaveOptions.DisableFormatting);
                    Equal("NoChange", McpCall(s, "undo_edit", new { snapshot_id = id })["status"]);
                    Equal(1, api.Creates); Equal(operation == "move_page" ? 1 : 0, api.Recycles); Equal(1, api.Pages[target].Writes);
                    Equal(edited, api.Pages[target].Page.ToString(SaveOptions.DisableFormatting));
                }
            }
        });
        Test("Agent conversion fingerprints detect tag name case and whitespace edits", () =>
        {
            foreach (var name in McpTagNameVariants().Skip(1))
            {
                var page = McpFidelityPage(true); var before = AgentCode.Fingerprint(McpFidelityObject(page), page);
                McpFidelityDefinition(page).SetAttributeValue("name", name);
                True(before != AgentCode.Fingerprint(McpFidelityObject(page), page));
            }
        });
        Test("MCP write signatures distinguish custom tag definitions bullet symbols and string metadata", () =>
        {
            foreach (var text in new[] { false, true })
            {
                var page = McpFidelityPage(text); var signature = McpPageModel.WriteSignature(page);
                foreach (var change in McpFidelityChanges())
                {
                    var actual = new XElement(page); change.Value(actual);
                    if (signature == McpPageModel.WriteSignature(actual)) throw new Exception(change.Key + ": changed content was accepted (text=" + text + ")");
                }
            }
        });
        Test("MCP mark fidelity still tolerates native definition renumbering defaults and numeric spellings", () =>
        {
            foreach (var text in new[] { false, true })
            {
                var expected = McpFidelityPage(text); var actual = new XElement(expected);
                var definition = McpFidelityDefinition(actual); definition.SetAttributeValue("index", "40");
                definition.SetAttributeValue("type", "00");
                definition.SetAttributeValue("fontColor", "#123456"); definition.SetAttributeValue("highlightColor", "#abcdef");
                definition.ReplaceAttributes(definition.Attributes().Reverse().ToArray());
                var tag = McpFidelityObject(actual).Element(One + "Tag"); tag.SetAttributeValue("index", "40"); tag.SetAttributeValue("disabled", null);
                tag.SetAttributeValue("creationDate", "2026-10-03T01:00:00Z");
                McpFidelityObject(actual).Element(One + "List").Element(One + "Bullet").SetAttributeValue("bullet", "02");
                Equal(McpPageModel.WriteSignature(expected), McpPageModel.WriteSignature(actual));
                var defaults = McpFidelityPage(text); var omitted = new XElement(defaults);
                McpFidelityDefinition(defaults).SetAttributeValue("fontColor", "automatic"); McpFidelityDefinition(defaults).SetAttributeValue("highlightColor", "none");
                McpFidelityDefinition(omitted).SetAttributeValue("fontColor", null); McpFidelityDefinition(omitted).SetAttributeValue("highlightColor", null);
                Equal(McpPageModel.WriteSignature(defaults), McpPageModel.WriteSignature(omitted));
            }
        });
        Test("MCP numeric normalization is limited to known geometry and spacing fields", () =>
        {
            var expected = McpFidelityPage(false); var outline = expected.Element(One + "Outline");
            outline.AddFirst(new XElement(One + "Position", new XAttribute("x", "36"), new XAttribute("y", "72"), new XAttribute("z", "0")));
            outline.Add(new XElement(One + "Size", new XAttribute("width", "420"), new XAttribute("isSetByUser", "true")));
            McpFidelityObject(expected).SetAttributeValue("spaceBefore", "12"); McpFidelityObject(expected).SetAttributeValue("spaceAfter", "6");
            var actual = new XElement(expected);
            foreach (var a in actual.Descendants(One + "Position").Attributes()) a.Value += ".0";
            actual.Element(One + "Outline").Element(One + "Size").SetAttributeValue("width", "420.000");
            McpFidelityObject(actual).SetAttributeValue("spaceBefore", "12.0"); McpFidelityObject(actual).SetAttributeValue("spaceAfter", "6.000");
            Equal(McpPageModel.WriteSignature(expected), McpPageModel.WriteSignature(actual));
            var changed = new XElement(actual); McpFidelityObject(changed).Element(One + "Meta").SetAttributeValue("content", "1");
            True(McpPageModel.WriteSignature(expected) != McpPageModel.WriteSignature(changed));
            var locked = McpMarkedObjectPage("table"); var column = locked.Descendants(One + "Column").Single(); column.SetAttributeValue("isLocked", "true");
            var decimalWidth = new XElement(locked); decimalWidth.Descendants(One + "Column").Single().SetAttributeValue("width", "120.0");
            Equal(McpPageModel.WriteSignature(locked), McpPageModel.WriteSignature(decimalWidth));
        });
        Test("MCP damaged copy and move results keep the source and freeze without replay", () =>
        {
            foreach (var text in new[] { false, true }) foreach (var operation in new[] { "copy_page", "move_page" }) foreach (var change in McpFidelityChanges())
            {
                var api = new FakeWorkspace(); api.Add("page", McpFidelityPage(text));
                using (var s = WorkspaceService(api))
                {
                    var id = WorkspaceBegin(s); McpCall(s, operation, new { snapshot_id = id, page_id = "page", section_id = "s2" });
                    api.AfterUpdate = () => change.Value(api.Pages.Single(p => p.Key != "page").Value.Page);
                    var result = McpFinish(s, id, 1);
                    if ((string)result["status"] != "CommitOutcomeUnknown") throw new Exception(operation + ": " + change.Key + " was accepted (text=" + text + ")");
                    Equal(false, result["source_recycled"]); Equal(0, api.Recycles); True(api.Pages.ContainsKey("page")); True(api.Pages.ContainsKey((string)result["page_id"]));
                    Equal(false, McpCall(s, "get_edit_status", new { snapshot_id = id })["can_undo"]);
                    var writes = api.Pages.Values.Sum(p => p.Writes);
                    Equal(McpJson.Serialize(result), McpJson.Serialize(McpFinish(s, id, 1))); Equal(1, api.Creates); Equal(writes, api.Pages.Values.Sum(p => p.Writes));
                }
            }
        });
        Test("MCP damaged move compensation keeps the verified target and cannot replay undo", () =>
        {
            foreach (var text in new[] { false, true }) foreach (var change in McpFidelityChanges())
            {
                var api = new FakeWorkspace(); api.Add("page", McpFidelityPage(text));
                using (var s = WorkspaceService(api))
                {
                    var id = WorkspaceBegin(s); McpCall(s, "move_page", new { snapshot_id = id, page_id = "page", section_id = "s2" });
                    var moved = McpFinish(s, id, 1); Equal("Verified", moved["status"]); var target = (string)moved["page_id"];
                    api.AfterUpdate = () => change.Value(api.Pages.Single(p => p.Key != target).Value.Page);
                    var undo = McpCall(s, "undo_edit", new { snapshot_id = id });
                    if ((string)undo["status"] != "CommitOutcomeUnknown") throw new Exception("compensation: " + change.Key + " was accepted (text=" + text + ")");
                    True(api.Pages.ContainsKey(target)); True(api.Pages.ContainsKey((string)undo["restored_page_id"])); Equal(1, api.Recycles);
                    Equal(false, McpCall(s, "get_edit_status", new { snapshot_id = id })["can_undo"]);
                    var writes = api.Pages.Values.Sum(p => p.Writes);
                    Equal(McpJson.Serialize(undo), McpJson.Serialize(McpCall(s, "undo_edit", new { snapshot_id = id })));
                    Equal(2, api.Creates); Equal(1, api.Recycles); Equal(writes, api.Pages.Values.Sum(p => p.Writes));
                }
            }
        });
    }
}
