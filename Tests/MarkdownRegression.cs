using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using OneNoteCodeHelper.Highlighting;
using OneNoteCodeHelper.Highlighting.Themes;
using OneNoteCodeHelper.Services;
using OneNoteCodeHelper.Services.Agent;
using OneNoteCodeHelper.Services.Markdown;

// 「插入代码」窗口的 Markdown 转换：解析、行内格式、生成的 OneNote XML 和插入页面。不调用 OneNote。
internal static partial class Program
{
    private static string MdText(MarkdownBlock block) => MarkdownBlock.PlainText(block.Inline);

    private static string MdText(IEnumerable<MarkdownRun> runs) => MarkdownBlock.PlainText(runs);

    private static void TestMarkdown()
    {
        Test("Markdown headings map to presets, setext headings and closing hashes", () =>
        {
            var b = MarkdownParser.Parse("# 一级 #\n## 二级\n### 三级\n#### 四级\n标题\n===\n副标题\n---\n#没有空格").Blocks;
            Equal(7, b.Count);
            Equal("heading1", b[0].Preset); Equal("一级", MdText(b[0]));
            Equal("heading2", b[1].Preset);
            Equal(ParagraphStyles.Heading3, b[2].Preset);
            Equal("body", b[3].Preset); True(b[3].BoldText);
            Equal(1, b[4].HeadingLevel); Equal("标题", MdText(b[4]));
            Equal(2, b[5].HeadingLevel); Equal("副标题", MdText(b[5]));
            Equal(MarkdownBlockKind.Paragraph, b[6].Kind); Equal("#没有空格", MdText(b[6]));
        });

        Test("Markdown lists nest by indent, number groups keep only their start, todo has no bullet", () =>
        {
            var b = MarkdownParser.Parse("1. 第一\n1. 第二\n   - 子项\n     更多说明\n1. 第三\n\n- [ ] 待办\n- [x] 完成\n\n5) 从五开始\n6) 六\n\t- Tab 缩进").Blocks;
            Equal("第一|第二|子项|更多说明|第三||待办|完成||从五开始|六|Tab 缩进", string.Join("|", b.Select(MdText)));
            Equal(1, b[0].Number); Equal("number", b[0].ListKind); Equal(0, b[0].Level);
            Equal(null, b[1].Number); Equal("number", b[1].ListKind);
            Equal("bullet", b[2].ListKind); Equal(1, b[2].Level);
            Equal(MarkdownBlockKind.Paragraph, b[3].Kind); Equal(2, b[3].Level);
            Equal(null, b[4].Number); Equal(0, b[4].Level);
            Equal(MarkdownBlockKind.Blank, b[5].Kind);
            Equal(null, b[6].ListKind); Equal(false, b[6].Todo);
            Equal(true, b[7].Todo);
            Equal(MarkdownBlockKind.Blank, b[8].Kind);
            Equal(5, b[9].Number); Equal(null, b[10].Number);
            Equal(1, b[11].Level); Equal("bullet", b[11].ListKind);
        });

        Test("Markdown blank lines collapse, quotes nest, rules become one blank line", () =>
        {
            var doc = MarkdownParser.Parse("\n\n> 引用\n> > 二层\n>\n> 第三行\n\n---\n\n\n正文  \n\n\n");
            Equal("Quote,Quote,Blank,Quote,Blank,Paragraph", string.Join(",", doc.Blocks.Select(x => x.Kind)));
            Equal(0, doc.Blocks[0].Level); Equal(1, doc.Blocks[1].Level); Equal("二层", MdText(doc.Blocks[1]));
            Equal("正文", MdText(doc.Blocks[5]));
            // 列表项之间的空行、紧跟列表项的续行
            var list = MarkdownParser.Parse("- 甲\n\n- 乙\n续行\n\n正文").Blocks;
            Equal("ListItem,ListItem,Paragraph,Blank,Paragraph", string.Join(",", list.Select(x => x.Kind)));
            Equal(1, list[2].Level); Equal(0, list[4].Level);
        });

        Test("Markdown fences resolve aliases, keep inner blank lines, run to the end when unclosed", () =>
        {
            var b = MarkdownParser.Parse("```js\nconst a = 1;\n\n  b();\n```\n~~~\nplain\n~~~\n```mermaid\ngraph TD\n```\n- 步骤\n  ```bash\n  ls -la\n  ```\n```python\nprint(1)").Blocks;
            Equal("Code,Code,Code,ListItem,Code,Code", string.Join(",", b.Select(x => x.Kind)));
            Equal("javascript", b[0].Language.Id); Equal("const a = 1;\n\n  b();", b[0].Code);
            Equal("text", b[1].Language.Id); Equal("plain", b[1].Code);
            Equal("text", b[2].Language.Id);
            Equal(1, b[4].Level); Equal("bash", b[4].Language.Id); Equal("ls -la", b[4].Code);
            Equal("python", b[5].Language.Id); Equal("print(1)", b[5].Code);
        });

        Test("Markdown indented code needs a blank line before it", () =>
        {
            var b = MarkdownParser.Parse("说明\n\n    int x = 1;\n    x++;\n\n结束\n    不是代码").Blocks;
            Equal("Paragraph,Blank,Code,Blank,Paragraph,Paragraph", string.Join(",", b.Select(x => x.Kind)));
            Equal("int x = 1;\nx++;", b[2].Code);
            Equal("不是代码", MdText(b[5]));
        });

        Test("Markdown tables keep alignment, escaped pipes and ragged rows", () =>
        {
            var doc = MarkdownParser.Parse("| 名称 | 数量 | 备注 |\n| :--- | :---: | ---: |\n| A \\| B | **2** |\n| C | 3 | 多 | 余 |\n\na | b\n不是分隔行");
            var table = doc.Blocks[0];
            Equal(MarkdownBlockKind.Table, table.Kind);
            Equal(3, table.Rows.Count); Equal(4, table.Rows.Max(r => r.Count)); True(table.Rows.All(r => r.Count == 4));
            Equal("left,center,right,", string.Join(",", table.Alignments));
            Equal("A | B", MdText(table.Rows[1][0]));
            True(table.Rows[1][1].Single().Bold);
            Equal("余", MdText(table.Rows[2][3]));
            Equal("Table,Blank,Paragraph,Paragraph", string.Join(",", doc.Blocks.Select(x => x.Kind)));
        });

        Test("Markdown inline emphasis follows strip_markdown pairing rules", () =>
        {
            var doc = new MarkdownDocument();
            var runs = MarkdownInline.Parse("**粗** *斜* ***两者*** ~~删~~ snake_case 2 * 3 a*b*c", doc);
            Equal("粗 斜 两者 删 snake_case 2 * 3 a*b*c", MdText(runs));
            True(runs.Single(r => r.Text == "粗").Bold);
            True(runs.Single(r => r.Text == "斜").Italic);
            var both = runs.Single(r => r.Text == "两者"); True(both.Bold && both.Italic);
            True(runs.Single(r => r.Text == "删").Strike);
            True(runs.Where(r => r.Text.Contains("snake_case")).All(r => !r.Bold && !r.Italic));
            Equal(0, doc.WarningCount);
        });

        Test("Markdown inline code, escapes and HTML stay literal", () =>
        {
            var doc = new MarkdownDocument();
            var runs = MarkdownInline.Parse("`a*b*` 和 `` x ` y `` \\*不是强调\\* 第一行<br>第二行 <b>粗</b>", doc);
            Equal("a*b* 和 x ` y *不是强调* 第一行\n第二行 <b>粗</b>", MdText(runs));
            True(runs.First().Code); Equal("a*b*", runs.First().Text);
            True(runs.Single(r => r.Text == "x ` y").Code);
            True(runs.All(r => !r.Italic && !r.Bold));
            Equal(1, runs.Count(r => r.Break));
            Equal(2, doc.Warnings["html"]);
            var html = MarkdownInline.ToHtml(MarkdownInline.Parse("a < b & c", doc), "Consolas");
            Equal("a &lt; b &amp; c", html);
        });

        Test("Markdown links accept only safe schemes, bare URLs stop at Chinese punctuation, images are not downloaded", () =>
        {
            var doc = new MarkdownDocument();
            var runs = MarkdownInline.Parse("[官网](https://example.com) [相对](./a.md) [脚本](javascript:alert(1)) <https://x.org> <a@b.com>", doc);
            Equal("官网 相对 脚本 https://x.org a@b.com", MdText(runs));
            Equal("https://example.com", runs.Single(r => r.Text == "官网").Href);
            True(runs.Where(r => r.Text.Contains("相对") || r.Text.Contains("脚本")).All(r => r.Href == null));
            Equal("https://x.org", runs.Single(r => r.Text == "https://x.org").Href);
            Equal("mailto:a@b.com", runs.Single(r => r.Text == "a@b.com").Href);
            Equal(2, doc.Warnings["link"]);

            runs = MarkdownInline.Parse("见https://example.com/a_b_c。下一句 (https://x.org/p).", doc);
            Equal("https://example.com/a_b_c", runs.Single(r => r.Text == "https://example.com/a_b_c").Href);
            Equal("https://x.org/p", runs.Single(r => r.Text == "https://x.org/p").Href);
            True(runs.All(r => !r.Italic));

            runs = MarkdownInline.Parse("[`run()` 的**说明**](https://x.org/doc_a) [https://a.com](https://a.com)", doc);
            Equal("run() 的说明 https://a.com", MdText(runs));
            var codeLink = runs.Single(r => r.Text == "run()"); True(codeLink.Code); Equal("https://x.org/doc_a", codeLink.Href);
            var boldLink = runs.Single(r => r.Text == "说明"); True(boldLink.Bold); Equal("https://x.org/doc_a", boldLink.Href);
            Equal("https://a.com", runs.Single(r => r.Text == "https://a.com").Href);

            runs = MarkdownInline.Parse("![图](https://img.example/x.png) ![](./y.png)", doc);
            Equal("图 ./y.png", MdText(runs));
            Equal("https://img.example/x.png", runs.Single(r => r.Text == "图").Href);
            Equal(2, doc.Warnings["image"]);
        });

        Test("Markdown badges link the image text to the outer link, brackets in link text pair up", () =>
        {
            var doc = new MarkdownDocument();
            var runs = MarkdownInline.Parse("[![build](https://img.shields.io/b.svg)](https://github.com/x/y) [![](https://img.example/c.svg)](https://c.org) [a [b] c](https://d.org) [x] 任务 [![徽章](./相对.png) 项目](./readme.md)", doc);
            Equal("build https://img.example/c.svg a [b] c [x] 任务 徽章 项目", MdText(runs));
            Equal("https://github.com/x/y", runs.Single(r => r.Text == "build").Href);
            Equal("https://c.org", runs.Single(r => r.Text == "https://img.example/c.svg").Href);
            Equal("https://d.org", runs.Single(r => r.Text == "a [b] c").Href);
            True(runs.Where(r => r.Text.Contains("徽章")).All(r => r.Href == null));
            Equal(3, doc.Warnings["image"]); Equal(1, doc.Warnings["link"]);

            // 图片说明里不再认链接；满行的 [ 不会拖慢
            runs = MarkdownInline.Parse("![a [b](https://e.org) c](https://img.example/d.png)", doc);
            Equal("a [b](https://e.org) c", MdText(runs));
            True(runs.All(r => r.Href == "https://img.example/d.png"));
            var watch = System.Diagnostics.Stopwatch.StartNew();
            MarkdownInline.Parse(new string('[', 50000), doc);
            True(watch.ElapsedMilliseconds < 5000);
        });

        Test("Markdown decodes HTML entities outside code and keeps unknown ones", () =>
        {
            var doc = new MarkdownDocument();
            var runs = MarkdownInline.Parse("AT&amp;T &lt;div&gt; a&nbsp;b &#169; &#x1F600; &#42;不强调&#42; &foo; &#0; `&amp;` \\&amp; [链](https://a.com/?x=1&amp;y=2)", doc);
            Equal("AT&T <div> a b © \U0001F600 *不强调* &foo; &#0; &amp; &amp; 链", MdText(runs));
            True(runs.All(r => !r.Italic));
            Equal("https://a.com/?x=1&y=2", runs.Single(r => r.Text == "链").Href);
            Equal(0, doc.WarningCount);
            Equal("&lt;div&gt; &amp;", MarkdownInline.ToHtml(MarkdownInline.Parse("&lt;div&gt; &amp;", doc), "Consolas"));
        });

        Test("Markdown trailing backslash before another line is a line break, not text", () =>
        {
            var b = MarkdownParser.Parse("甲\\\n乙\\\n\n- 项\\\n  续\n> 引\\\n> 用\\\n>\n双\\\\\n尾").Blocks;
            Equal("Paragraph,Paragraph,Blank,ListItem,Paragraph,Quote,Quote,Blank,Paragraph,Paragraph", string.Join(",", b.Select(x => x.Kind)));
            Equal("甲|乙\\||项|续|引|用\\||双\\|尾", string.Join("|", b.Select(MdText)));
        });

        Test("Markdown front matter becomes a YAML code box instead of rules and a setext heading", () =>
        {
            var b = MarkdownParser.Parse("\n---\ntitle: 笔记\ntags:\n  - a\n# 注释\n---\n\n# 正文标题").Blocks;
            Equal("Code,Blank,Heading", string.Join(",", b.Select(x => x.Kind)));
            Equal("yaml", b[0].Language.Id); Equal("title: 笔记\ntags:\n  - a\n# 注释", b[0].Code);
            Equal("正文标题", MdText(b[2]));

            // 不像 YAML 的、没有结尾的、不在开头的不算：照旧是分隔线、Setext 标题和正文
            Equal(MarkdownBlockKind.Heading, MarkdownParser.Parse("---\n普通一段话\n---").Blocks.Single().Kind);
            Equal(MarkdownBlockKind.Paragraph, MarkdownParser.Parse("---\ntitle: 没有结尾").Blocks.Single().Kind);
            Equal("Paragraph,Blank,Heading", string.Join(",", MarkdownParser.Parse("正文\n\n---\ntitle: x\n---").Blocks.Select(x => x.Kind)));
        });

        Test("Markdown fences keep the written language: unsupported names stay plain text, only bare fences are detected", () =>
        {
            var b = MarkdownParser.Parse("```diff\n- var a = 1;\n+ const a = 1;\n```\n```ruby\ndef greet(name)\n  puts \"Hello\"\nend\n```\n```{.python}\nprint(1)\n```\n```js {1,3}\nconst a = 1;\n```\n```\nimport os\n\ndef main():\n    print(os.getcwd())\n```").Blocks;
            Equal("text,text,python,javascript,python", string.Join(",", b.Select(x => x.Language.Id)));
            Equal("xml", LanguageRegistry.FindByName("xaml").Id);
            Equal("css", LanguageRegistry.FindByName("SCSS").Id);
            Equal("sql", LanguageRegistry.FindByName("postgresql").Id);
        });

        Test("Markdown writer builds styled paragraphs, lists, todo, nested code box and table", () =>
        {
            var doc = MarkdownParser.Parse("# 标题\n- [ ] 待办 **重点**\n1. 一\n   ```python\n   print(1)\n   ```\n\n| a | b |\n|---|:-:|\n| 1 | `2` |");
            var page = Page(Paragraph("a", "已有内容"));
            var styles = new XElement("styles", page.Elements(One + "QuickStyleDef"));
            var tags = new XElement("tags", page.Elements(One + "TagDef"));
            var settings = new AddInSettings();
            var nodes = MarkdownWriter.Write(doc, settings, new AgentOptions(), styles, tags);
            Equal(5, nodes.Count);
            var all = nodes.SelectMany(n => n.DescendantsAndSelf(One + "OE")).ToList();
            foreach (var oe in all.Where(e => e.Attribute("quickStyleIndex") != null))
                True(styles.Elements(One + "QuickStyleDef").Any(d => (string)d.Attribute("index") == (string)oe.Attribute("quickStyleIndex")));
            True(all.All(oe => oe.Attribute(AgentLayout.Key) == null));
            Equal("h1", (string)styles.Elements(One + "QuickStyleDef").Single(d => (string)d.Attribute("index") == (string)nodes[0].Attribute("quickStyleIndex")).Attribute("name"));

            var todo = nodes[1];
            Equal(null, todo.Element(One + "List"));
            var tag = todo.Element(One + "Tag");
            True(tags.Elements(One + "TagDef").Any(d => (string)d.Attribute("index") == (string)tag.Attribute("index")));
            True(todo.Element(One + "T").Value.Contains("font-weight:bold"));
            Equal("待办 重点", new AgentRichText(todo).Text);

            var numbered = nodes[2];
            Equal("1", (string)numbered.Element(One + "List").Element(One + "Number").Attribute("restartNumberingAt"));
            var code = numbered.Element(One + "OEChildren").Element(One + "OE").Element(One + "Table");
            True(XNode.DeepEquals(CodeBlockBuilder.BuildTable("print(1)", LanguageRegistry.Find("python"), settings.Theme, settings), code));

            Equal(MarkdownBlockKind.Blank, doc.Blocks[4].Kind);
            var table = nodes[4].Element(One + "Table");
            Equal("true", (string)table.Attribute("hasHeaderRow"));
            Equal(2, table.Element(One + "Columns").Elements().Count());
            var rows = table.Elements(One + "Row").ToList();
            True(rows[0].Elements(One + "Cell").All(c => (string)c.Attribute("shadingColor") == "#F2F2F2"));
            Equal("center", (string)rows[1].Elements(One + "Cell").ElementAt(1).Descendants(One + "OE").Single().Attribute("alignment"));
            True(rows[1].Elements(One + "Cell").ElementAt(1).Descendants(One + "T").Single().Value.Contains("font-family:" + settings.FontFamily));
            foreach (var oe in all.Where(e => e.Element(One + "T") != null)) new AgentRichText(oe);
        });

        Test("Markdown insert appends a new outline with tag definitions and keeps existing content", () =>
        {
            var fake = new FakePage(Page(Paragraph("a", "已有内容")));
            var result = MarkdownWriter.Insert(fake, "page", "# 新标题\n- [ ] 事项\n\n正文", new AddInSettings(), new AgentOptions(), out var verified);
            True(result.Success); True(verified);
            Equal(1, fake.Writes);
            Equal(2, fake.Page.Elements(One + "Outline").Count());
            Equal("已有内容", PageEditor.ExtractPlainText(fake.Page.Descendants(One + "OE").First(e => (string)e.Attribute("objectID") == "a")));
            True(XElement.Parse(fake.LastXml).Elements(One + "TagDef").Any());
            True(XElement.Parse(fake.LastXml).Elements(One + "Outline").Single().Attribute("objectID") == null);
        });

        Test("Markdown insert rebuilds after a timestamp conflict and gives up without writing", () =>
        {
            var fake = new FakePage(Page(Paragraph("a", "已有内容"))) { ConflictsRemaining = 1 };
            var ok = MarkdownWriter.Insert(fake, "page", "正文", new AddInSettings(), new AgentOptions(), out _);
            True(ok.Success); Equal(2, fake.Attempts); Equal(1, fake.Writes);

            var busy = new FakePage(Page(Paragraph("a", "已有内容"))) { ConflictsRemaining = 10 };
            var failed = MarkdownWriter.Insert(busy, "page", "正文", new AddInSettings(), new AgentOptions(), out _);
            True(!failed.Success); Equal(3, busy.Attempts); Equal(0, busy.Writes);
        });

        Test("Markdown insert rejects empty and oversized content", () =>
        {
            var fake = new FakePage(Page(Paragraph("a", "已有内容")));
            True(!MarkdownWriter.Insert(fake, "page", "\n  \n\n", new AddInSettings(), new AgentOptions(), out _).Success);
            var tooLong = MarkdownWriter.Insert(fake, "page", new string('字', MarkdownParser.MaxChars + 1), new AddInSettings(), new AgentOptions(), out _);
            True(!tooLong.Success && tooLong.Message.Contains("内容太长"));
            var tooMany = MarkdownWriter.Insert(fake, "page", string.Join("\n", Enumerable.Repeat("段", MarkdownParser.MaxParagraphs + 1)), new AddInSettings(), new AgentOptions(), out _);
            True(!tooMany.Success && tooMany.Message.Contains("内容太长"));
            Equal(0, fake.Writes);
        });

        Test("Markdown hint, fence aliases and heading3 preset stay out of Agent tools", () =>
        {
            True(MarkdownParser.LooksLikeMarkdown("## 标题\n- 项目"));
            True(MarkdownParser.LooksLikeMarkdown("```js\nx\n```\n说明"));
            True(!MarkdownParser.LooksLikeMarkdown("# 注释\nkey: value\n- item"));
            True(!MarkdownParser.LooksLikeMarkdown("# comment\ndef f():\n    return 1"));
            Equal("javascript", LanguageRegistry.FindByName("js").Id);
            Equal("csharp", LanguageRegistry.FindByName("C#").Id);
            Equal("python", LanguageRegistry.FindByName("Python").Id);
            Equal("text", LanguageRegistry.FindByName("mermaid").Id);
            Equal(null, LanguageRegistry.FindByName("foo"));
            Equal(null, LanguageRegistry.FindByName(null));
            Equal(null, LanguageRegistry.Find("js"));
            True(!ParagraphStyles.Ids.Contains(ParagraphStyles.Heading3));
            var h3 = ParagraphStyles.Definition(ParagraphStyles.Heading3, new AgentOptions());
            Equal("h3", (string)h3.Attribute("name")); Equal("true", (string)h3.Attribute("bold"));
            Equal(11d, ParagraphStyles.Appearance("body").Size);
        });
    }
}
