using System.Collections.Generic;
using System.Linq;
using OneNoteCodeHelper.Mcp;
using OneNoteCodeHelper.Services.Agent;

namespace OneNoteCodeHelper.Services.Mcp
{
    internal static class McpExtensionCatalog
    {
        internal static readonly Dictionary<string, AgentSchema> Schemas = new Dictionary<string, AgentSchema>();
        private static readonly Dictionary<string, string> Descriptions = new Dictionary<string, string>();
        internal static readonly string[] ContentTools = { "append_content", "insert_content", "replace_text" };
        internal static readonly string[] WorkspaceTools = { "create_page", "copy_page", "move_page", "create_section" };
        internal static readonly string[] ReadTools = { "list_nodes", "search_pages", "read_page", "read_selection", "find_tasks", "get_link" };
        static McpExtensionCatalog()
        {
            var paging = new Dictionary<string, AgentSchema> { ["cursor"] = AgentSchema.Str(), ["limit"] = AgentSchema.Num(1, 100, true) };
            Add("list_nodes", "逐层浏览已打开笔记本、分区组、分区及页面。默认 50 项，最多 100 项；游标绑定客户端并保留 5 分钟。", Fields(paging, "parent_id"));
            Add("search_pages", "使用 OneNote 原生搜索，不改变当前界面。返回路径、摘要和原文链接；root_id 可限定范围。", Fields(paging, "query", "root_id"), "query");
            Add("read_page", "独立读取指定页面，包含受保护代码、表格、链接和图片 OCR。blocks 或 markdown 输出；分页保持同一只读快照。", Format(Fields(paging, "page_id"), "blocks", "markdown"), "page_id");
            Add("read_selection", "独立读取当前页的完整选中段落，包含选中范围内空行；没有选区明确报错。", Format(new Dictionary<string, AgentSchema>(paging), "blocks", "markdown"));
            var tasks = Fields(paging, "root_id"); tasks["completed"] = new AgentSchema { Type = "boolean" };
            Add("find_tasks", "分批扫描原生待办标记，默认未完成。返回真实对象 ID、页面路径及链接；跳过页面明确说明。", tasks);
            Add("begin_workspace_edit", "建立页面或分区操作草稿。每个草稿只规划一项操作，检查 get_pending_changes 后 finish_edit 才写入。", new Dictionary<string, AgentSchema>());
            Add("append_content", "在页面下方追加独立文本框，支持空页；只修改草稿。plain/markdown，最多 5000 字和 50 段，服从插入能力开关。", Format(Fields(null, "snapshot_id", "content"), "plain", "markdown"), "snapshot_id", "content", "format");
            Add("insert_content", "在已读取的普通段落前后插入 plain/markdown 内容，检查选区边界，整框核验和撤销。", Format(Position(Fields(null, "snapshot_id", "target_id", "content")), "plain", "markdown"), "snapshot_id", "target_id", "position", "content", "format");
            var replace = Fields(null, "snapshot_id", "block_id", "quote", "replacement"); replace["replacement"].AllowEmpty = true;
            replace["occurrence"] = AgentSchema.Num(1, 100000, true);
            Add("replace_text", "精确更新普通正文，不跨换行，不增删段落。必须先 read_blocks；未修改字保留格式和链接，只改草稿。", replace, "snapshot_id", "block_id", "quote", "occurrence", "replacement");
            var create = Format(Fields(null, "snapshot_id", "section_id", "title", "content"), "plain", "markdown"); create["content"].AllowEmpty = true;
            Add("create_page", "在工作区草稿里规划新页面；section_id 必须是真实分区 ID，title 为原生标题，content 可选。提交后才返回真实新页面 ID。", create, "snapshot_id", "section_id", "title");
            foreach (var tool in new[] { "copy_page", "move_page" })
                Add(tool, tool == "copy_page" ? "规划单页保真复制，复杂对象或坏图提交前拒绝；不包含子页。" : "规划单页移动：复制并核验后源页入回收站。带子页父页面拒绝；返回新 ID 和链接，撤销通过恢复副本完成。",
                    Fields(null, "snapshot_id", "page_id", "section_id"), "snapshot_id", "page_id", "section_id");
            Add("create_section", "在工作区草稿里规划新分区。parent_id 是已打开笔记本或分区组，name 为分区名。相同名称已存在则拒绝。", Fields(null, "snapshot_id", "parent_id", "name"), "snapshot_id", "parent_id", "name");
            Add("get_link", "生成真实页面、分区或页面对象的 OneNote 链接，不能传入 p1 等草稿短 ID。", Fields(null, "page_id", "object_id"), "page_id");
            Add("navigate_to", "定位真实页面或页面对象；仅改变 OneNote 当前视图。", Fields(null, "page_id", "object_id"), "page_id");
            Add("export_page", "导出已保存页面为 Markdown 或 PDF。output_path 为绝对路径且不能已存在；先生成临时文件，再交付。", Format(Fields(null, "page_id", "output_path"), "markdown", "pdf"), "page_id", "format", "output_path");
        }
        private static Dictionary<string, AgentSchema> Fields(Dictionary<string, AgentSchema> basis, params string[] fields)
        { var map = basis == null ? new Dictionary<string, AgentSchema>() : new Dictionary<string, AgentSchema>(basis); foreach (var f in fields) map[f] = AgentSchema.Str(); return map; }
        private static Dictionary<string, AgentSchema> Format(Dictionary<string, AgentSchema> map, params string[] formats) { map["format"] = AgentSchema.Str(formats); return map; }
        private static Dictionary<string, AgentSchema> Position(Dictionary<string, AgentSchema> map) { map["position"] = AgentSchema.Str("before", "after"); return map; }
        private static void Add(string name, string description, Dictionary<string, AgentSchema> fields, params string[] required)
        { Schemas[name] = AgentSchema.Obj(fields, required); Descriptions[name] = description; }
        internal static IEnumerable<McpToolDefinition> Definitions => Schemas.Select(p => new McpToolDefinition
        {
            Name = p.Key, Description = Descriptions[p.Key], InputSchemaJson = McpJson.Serialize(p.Value.Json()),
            OutputSchemaJson = OutputSchema(p.Key),
            ReadOnly = ReadTools.Contains(p.Key), Destructive = false
        });
        private static string OutputSchema(string name)
        {
            object String(bool nullable = false) => new { type = nullable ? (object)new[] { "string", "null" } : "string" };
            var fields = new Dictionary<string, object> { ["ok"] = new { type = "boolean" }, ["error_code"] = String(), ["error"] = String() };
            if (ContentTools.Contains(name) || WorkspaceTools.Contains(name))
            { fields["draft_revision"] = new { type = "integer" }; fields["changed"] = new { type = "boolean" }; fields["warnings"] = new { type = "array", items = String() }; }
            else if (name == "begin_workspace_edit")
            { fields["snapshot_id"] = String(); fields["instance_id"] = String(); fields["scope"] = String(); fields["draft_revision"] = new { type = "integer" }; fields["available_tools"] = new { type = "array", items = String() }; }
            else if (name == "get_link") { fields["page_id"] = String(); fields["link"] = String(); }
            else if (name == "export_page")
            { fields["output_path"] = String(); fields["format"] = String(); fields["bytes"] = new { type = "integer" }; fields["last_modified"] = String(true); fields["complete"] = new { type = "boolean" }; }
            else if (name != "navigate_to")
            {
                fields["read_id"] = String(); fields["items"] = new { type = "array", items = new { type = "object", additionalProperties = true } };
                fields["next_cursor"] = String(true); fields["total"] = new { type = "integer" }; fields["expires_at"] = String();
                fields["complete"] = new { type = "boolean" }; fields["page_id"] = String(true); fields["markdown"] = String(true);
                fields["skipped"] = new { type = "array", items = new { type = "object", additionalProperties = true } };
            }
            return McpJson.Serialize(new { type = "object", properties = fields, additionalProperties = true });
        }
    }
}
