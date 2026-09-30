using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace OneNoteCodeHelper.Services
{
    /// <summary>Agent 窗口「模型」下拉里的一项。下拉里直接显示模型 id，也就是发给接口的 model 参数。</summary>
    internal sealed class AiModel
    {
        internal AiModel(string id)
        {
            Id = id;
        }

        internal string Id { get; }
    }

    /// <summary>Agent 窗口「功能」下拉里的一项文字功能：显示名 + 提示词，外加由插件自己做、不经过 AI 的规则。</summary>
    internal sealed class AiFunction
    {
        internal AiFunction(string name, string prompt, bool removeExtraBlankLines = false)
        {
            Name = name;
            Prompt = prompt;
            RemoveExtraBlankLines = removeExtraBlankLines;
        }

        internal string Name { get; }

        internal string Prompt { get; }

        /// <summary>
        /// 顺带删掉多余的空行：连续的空行只留一行，文本框开头、结尾的空行删掉。
        /// 删空行就是删段落，AI 按约定不能动段落，所以这一步由插件自己做，见 <see cref="BlankLines"/>。
        /// </summary>
        internal bool RemoveExtraBlankLines { get; }
    }

    /// <summary>
    /// 思考强度（变体），Agent 窗口的下拉里直接显示这些名字。取值和请求参数照搬 opencode 配置里 deepseek 模型的 variants：
    /// none 关掉思考（thinking.type = disabled，不传 reasoning_effort），其余打开思考并把名字作为 reasoning_effort 传过去。
    /// </summary>
    internal static class AiEfforts
    {
        internal const string DefaultId = "high";

        internal const string NoneId = "none";

        internal static readonly string[] Ids = { NoneId, "low", "medium", "high", "max" };

        internal static string Normalize(string id)
        {
            return Ids.FirstOrDefault(x => string.Equals(x, id?.Trim(), StringComparison.OrdinalIgnoreCase))
                   ?? DefaultId;
        }

        /// <summary>把这个变体对应的参数写进请求体。</summary>
        internal static void ApplyTo(IDictionary<string, object> body, string id)
        {
            id = Normalize(id);

            if (id == NoneId)
            {
                body["thinking"] = new Dictionary<string, object> { ["type"] = "disabled" };
                return;
            }

            body["thinking"] = new Dictionary<string, object> { ["type"] = "enabled" };
            body["reasoning_effort"] = id;
        }
    }

    /// <summary>
    /// AI 助手的配置：接口地址、Key、模型和各功能的提示词。对应 ai-settings.xml，在「AI 配置」窗口里改，也可以手改。
    /// Models、Functions 保证非空，调用方不用再判空。
    /// </summary>
    internal sealed class AiConfig
    {
        internal AiConfig(string apiUrl, string apiKey, int timeoutSeconds, int maxTokens,
            IReadOnlyList<AiModel> models, IReadOnlyList<AiFunction> functions)
        {
            ApiUrl = apiUrl;
            ApiKey = apiKey;
            TimeoutSeconds = timeoutSeconds;
            MaxTokens = maxTokens;
            Models = models;
            Functions = functions;
        }

        /// <summary>OpenAI 兼容接口的基础地址；请求时会补上 /chat/completions。</summary>
        internal string ApiUrl { get; }

        internal string ApiKey { get; }

        /// <summary>单次请求的超时（秒）。</summary>
        internal int TimeoutSeconds { get; }

        /// <summary>单次请求的 max_tokens，含思考部分。0 表示不传，用接口默认值。</summary>
        internal int MaxTokens { get; }

        internal IReadOnlyList<AiModel> Models { get; }

        internal IReadOnlyList<AiFunction> Functions { get; }

        internal Agent.AgentOptions Agent { get; set; } = new Agent.AgentOptions();

        /// <summary>按 id 找模型，找不到（比如配置里删掉了）就用第一个。</summary>
        internal AiModel FindModel(string id)
        {
            return Models.FirstOrDefault(m => string.Equals(m.Id, id, StringComparison.OrdinalIgnoreCase))
                   ?? Models[0];
        }

        internal AiFunction FindFunction(string name)
        {
            return Functions.FirstOrDefault(f => string.Equals(f.Name, name, StringComparison.OrdinalIgnoreCase))
                   ?? Functions[0];
        }

        internal int IndexOfModel(string id) => Math.Max(0, Models.ToList().IndexOf(FindModel(id)));

        internal int IndexOfFunction(string name) => Math.Max(0, Functions.ToList().IndexOf(FindFunction(name)));

        /// <summary>两份配置的下拉选项是否一样。不一样才需要重建 Agent 窗口的下拉，免得丢掉当前选中项。</summary>
        internal bool HasSameChoices(AiConfig other)
        {
            return other != null
                   && Models.Select(m => m.Id).SequenceEqual(other.Models.Select(m => m.Id))
                   && Functions.Select(f => f.Name).SequenceEqual(other.Functions.Select(f => f.Name));
        }
    }

    /// <summary>
    /// ai-settings.xml 的读写。
    ///
    /// 和 settings.xml 分开放：settings.xml 在每次切换功能区选项时都会被内存里的设置整体重写，
    /// 用户手改的 Key、提示词要是放在那里，OneNote 开着的时候改了也会被覆盖回去。
    /// 这份文件只在两种时候写：不存在时写一次默认值，以及在「AI 配置」窗口里点保存。
    /// 保存是在原文件上就地改（见 <see cref="Apply"/>），手写的注释和窗口里没有的节点都留着。
    ///
    /// 用 LINQ to XML 而不是 XmlSerializer，是为了能在默认文件里写注释，告诉用户每一项怎么填。
    /// </summary>
    internal static class AiConfigStore
    {
        private const string DefaultApiUrl = "https://api.deepseek.com";

        private const int DefaultTimeoutSeconds = 300;

        private const int DefaultMaxTokens = 16384;

        internal const int MinTimeoutSeconds = 10;

        internal const int MaxTimeoutSeconds = 3600;

        internal const int MinMaxTokens = 0;

        internal const int MaxMaxTokens = 1024 * 1024;

        /// <summary>根节点下各项在默认文件里的顺序。保存时缺的节点按这个顺序补在前一项后面。</summary>
        private static readonly string[] RootOrder =
            { "ApiUrl", "ApiKey", "TimeoutSeconds", "MaxTokens", "Agent", "Models", "Functions" };

        internal const string TypoFunctionName = "错别字修复";

        internal const string DefaultModelId = "deepseek-v4-flash";

        private const string RemoveBlankLinesAttribute = "removeExtraBlankLines";

        internal const string SmartFunctionName = "智能校正";

        /// <summary>
        /// Agent 窗口「功能」下拉的第一项：按需求调用格式和改错别字的工具，不是配置里的文字功能。
        /// 配置里要是有同名的 Function，窗口里跳过那一项。
        /// </summary>
        internal const string AgentFunctionName = "自定义排版（Agent）";

        private const string LegacyCombinedFunctionName = "错别字 + 排版";

        /// <summary>「错别字修复」要改的几类错误，「智能校正」也用这一份。</summary>
        private const string TypoRules =
            "- 错别字、同音字和形近字误用（如「在/再」「的/地/得」用错）；\n" +
            "- 漏字、多字、重复的字词；\n" +
            "- 明显的标点错误，以及英文单词的拼写错误。\n";

        /// <summary>「排版优化」的几条规则，「智能校正」也用这一份。</summary>
        private const string LayoutRules =
            "- 中文与英文、中文与数字之间加一个半角空格；数字与单位按惯例处理（如 10 GB、20%）；\n" +
            "- 中文语境使用全角标点，英文句子内部使用半角标点；\n" +
            "- 去掉多余的空格，修正重复或误用的标点；\n" +
            "- 专有名词使用正确的大小写（如 GitHub、iOS、JavaScript）。\n";

        private const string TypoPrompt =
            "你是一名严谨的中文校对编辑，负责修正笔记里的文字错误：\n" +
            TypoRules +
            "只改错误本身：不要改写句子，不要调整语气和用词风格，不要增删内容。没有错误的段落保持原样。";

        private const string LayoutPrompt =
            "你是一名中文排版编辑，请参照《中文文案排版指北》优化每个段落的排版：\n" +
            LayoutRules +
            "只调整排版，不要改动文字内容和意思，不要改写句子。";

        private const string CombinedPrompt =
            "你是一名严谨的中文校对和排版编辑，对每个段落同时做两件事。\n" +
            "一、修正文字错误：\n" +
            TypoRules +
            "二、参照《中文文案排版指北》优化排版：\n" +
            LayoutRules +
            "只改错误和排版：不要改写句子，不要调整语气和用词风格，不要增删内容。没有问题的段落保持原样。";

        internal static string ConfigPath { get; } = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "OneNoteCodeHelper",
            "ai-settings.xml");

        internal static AiConfig Default { get; } = new AiConfig(
            DefaultApiUrl,
            string.Empty,
            DefaultTimeoutSeconds,
            DefaultMaxTokens,
            new[]
            {
                new AiModel(DefaultModelId),
                new AiModel("deepseek-v4-pro")
            },
            new[]
            {
                new AiFunction(SmartFunctionName, CombinedPrompt, removeExtraBlankLines: true),
                new AiFunction(TypoFunctionName, TypoPrompt),
                new AiFunction("排版优化", LayoutPrompt, removeExtraBlankLines: true)
            });

        /// <summary>读配置。文件不存在或写坏了都退回默认值，不抛异常。</summary>
        internal static AiConfig Load()
        {
            TryLoad(out var config, out _);
            return config;
        }

        /// <summary>
        /// 读配置。读不了（文件正被编辑器写着、XML 写坏了）时返回 false，config 给默认值，error 是原因，已记日志。
        /// 文件不存在不算失败，给默认值。
        /// </summary>
        internal static bool TryLoad(out AiConfig config, out Exception error)
        {
            error = null;

            try
            {
                if (!File.Exists(ConfigPath))
                {
                    config = Default;
                    return true;
                }

                var root = XDocument.Load(ConfigPath).Root;
                config = root == null ? Default : Parse(root);
                return true;
            }
            catch (Exception ex)
            {
                AddInLog.Warn("读取 AI 配置失败。文件：" + ConfigPath, ex);
                config = Default;
                error = ex;
                return false;
            }
        }

        /// <summary>从 ai-settings.xml 的根元素读出配置，缺的、写空的项用默认值。</summary>
        internal static AiConfig Parse(XElement root)
        {
            var models = root.Element("Models")?.Elements("Model")
                .Select(e => new AiModel(Trim((string)e.Attribute("id"))))
                .Where(m => m.Id.Length > 0)
                .ToList();

            var functions = root.Element("Functions")?.Elements("Function")
                .Select(ReadFunction)
                .Where(f => f.Name.Length > 0 && f.Prompt.Length > 0)
                .ToList();

            var url = Trim((string)root.Element("ApiUrl"));

            return new AiConfig(
                url.Length > 0 ? url : DefaultApiUrl,
                Trim((string)root.Element("ApiKey")),
                ReadInt(root.Element("TimeoutSeconds"), DefaultTimeoutSeconds, MinTimeoutSeconds, MaxTimeoutSeconds),
                ReadInt(root.Element("MaxTokens"), DefaultMaxTokens, MinMaxTokens, MaxMaxTokens),
                models?.Count > 0 ? models : Default.Models,
                functions?.Count > 0 ? functions : Default.Functions)
            { Agent = Agent.AgentOptions.Parse(root.Element("Agent")) };
        }

        private static AiFunction ReadFunction(XElement element)
        {
            var name = Trim((string)element.Attribute("name"));

            // 没写这个属性时和同名的内置功能一样。旧配置中的「错别字 + 排版」
            // 也继续删空行；写了 false 就关掉。
            var defaultRemoveBlankLines = Default.Functions.FirstOrDefault(f => f.Name == name)?.RemoveExtraBlankLines
                ?? string.Equals(name, LegacyCombinedFunctionName, StringComparison.Ordinal);
            var removeBlankLines = bool.TryParse(Trim((string)element.Attribute(RemoveBlankLinesAttribute)), out var value)
                ? value
                : defaultRemoveBlankLines;

            return new AiFunction(name, Trim((string)element.Element("Prompt")), removeBlankLines);
        }

        /// <summary>配置文件不存在时写出一份带注释的默认配置，返回文件路径。</summary>
        internal static string EnsureFile()
        {
            if (File.Exists(ConfigPath))
            {
                return ConfigPath;
            }

            var directory = Path.GetDirectoryName(ConfigPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            BuildDefaultDocument().Save(ConfigPath);
            AddInLog.Info("已生成默认 AI 配置：" + ConfigPath);
            return ConfigPath;
        }

        internal static XDocument BuildDefaultDocument()
        {
            var config = Default;

            return new XDocument(
                new XDeclaration("1.0", "utf-8", null),
                new XComment(" OneNote 代码高亮 · AI 助手配置。改完保存即可，回到 Agent 窗口点「执行」时生效。 "),
                new XElement(
                    "AiConfig",
                    new XComment(" OpenAI 兼容接口的基础地址；插件会在后面加上 /chat/completions "),
                    new XElement("ApiUrl", config.ApiUrl),
                    new XElement("ApiKey", config.ApiKey),
                    new XComment(" 单次请求的超时（秒）。思考强度开得高、要处理的内容又多时可以调大 "),
                    new XElement("TimeoutSeconds", config.TimeoutSeconds),
                    new XComment(" 单次请求最多输出多少 token（含思考过程）。0 表示用接口的默认值 "),
                    new XElement("MaxTokens", config.MaxTokens),
                    new XComment(" Agent 工具调用：默认启用标题、段间距、图文容器、代码框转换、列表、标记、表格样式，以及删空行、缩进、移动（含跨文本框移动和合并文本框）、插入段落和转表格等结构调整，哪项回存有问题或不想让 Agent 改结构，就把对应的 Enable 开关设为 false。" +
                        "DefaultRequest 是打开 Agent 时预填的需求，可在执行前修改；工具协议由插件附加。" +
                        "若接口不接受思考参数，可把 SendThinking 设为 false。旧配置不写此节点也可使用默认值。 "),
                    new XElement("Agent", new Agent.AgentOptions().ToElements()),
                    new XComment(" Agent 窗口「模型」下拉里的选项：id 是接口的模型名，下拉里直接显示它 "),
                    new XElement(
                        "Models",
                        config.Models.Select(m => new XElement("Model", new XAttribute("id", m.Id)))),
                    new XComment(
                        " Agent 窗口「功能」下拉里的文字功能，可以自己加。Prompt 只需写清楚要做什么；\n" +
                        "       输入输出的 JSON 格式、只返回改动的段落、不要合并拆分段落等约定由插件自动附加，不用写。\n" +
                        "       removeExtraBlankLines=\"true\" 表示顺带删掉多余的空行：连续的空行只留一行，文本框开头、结尾的空行删掉。\n" +
                        "       这一步由插件自己做，不经过 AI。 "),
                    new XElement(
                        "Functions",
                        config.Functions.Select(f => new XElement(
                            "Function",
                            new XAttribute("name", f.Name),
                            f.RemoveExtraBlankLines ? new XAttribute(RemoveBlankLinesAttribute, "true") : null,
                            new XElement("Prompt", f.Prompt))))));
        }

        /// <summary>配置文件的修改时间，用来判断文件是不是在别处被改过。文件不存在时是个固定的很早的时间。</summary>
        internal static DateTime Stamp()
        {
            try { return File.GetLastWriteTimeUtc(ConfigPath); }
            catch (Exception) { return DateTime.MinValue; }
        }

        /// <summary>
        /// 把 config 写进 ai-settings.xml（path 为 null 时是 <see cref="ConfigPath"/>）。在原文件上就地改，见 <see cref="Apply"/>；
        /// 文件不存在时从带注释的默认文件改起。原文件写坏了（XML 解析不了）先复制成 .bak 再整个重写。
        /// 先写同目录的临时文件再替换，Agent 窗口这时去读也不会读到写了一半的文件。读写失败直接抛出，由窗口提示。
        /// </summary>
        internal static void Save(AiConfig config, string path = null)
        {
            path = path ?? ConfigPath;
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            XDocument document = null;
            if (File.Exists(path))
            {
                try
                {
                    document = XDocument.Load(path);
                }
                catch (XmlException ex)
                {
                    File.Copy(path, path + ".bak", true);
                    AddInLog.Warn("AI 配置文件解析不了，已备份为 .bak 后重写：" + path, ex);
                }
            }

            if (document?.Root == null)
            {
                document = BuildDefaultDocument();
            }

            Apply(document, config);

            var temp = path + ".tmp";
            document.Save(temp);
            if (!File.Exists(path))
            {
                File.Move(temp, path);
            }
            else
            {
                try
                {
                    File.Replace(temp, path, null);
                }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is PlatformNotSupportedException)
                {
                    // 重定向到网络位置的 AppData 之类不支持 ReplaceFile 的地方，退回直接覆盖。
                    AddInLog.Warn("替换 AI 配置文件失败，改为直接覆盖。", ex);
                    File.Copy(temp, path, true);
                    File.Delete(temp);
                }
            }

            AddInLog.Info("已保存 AI 配置：" + path);
        }

        /// <summary>
        /// 把 config 写进已有的配置文档，只动窗口里能改的部分：
        /// 根节点下的四个值原地改，缺的按默认文件的顺序补上；Models、Functions 保留节点本身，里面整个按 config 重建，
        /// removeExtraBlankLines 一律明写，不靠「没写时跟同名内置功能走」；Agent 节点已有时改全部已知项，
        /// 没有时只在有一项不是默认值时才加，保持旧文件不写这一节也能用的约定。其余节点和注释原样留着。
        /// </summary>
        internal static void Apply(XDocument document, AiConfig config)
        {
            var root = document.Root;
            Child(root, "ApiUrl", RootOrder).Value = config.ApiUrl;
            Child(root, "ApiKey", RootOrder).Value = config.ApiKey;
            Child(root, "TimeoutSeconds", RootOrder).Value = config.TimeoutSeconds.ToString(CultureInfo.InvariantCulture);
            Child(root, "MaxTokens", RootOrder).Value = config.MaxTokens.ToString(CultureInfo.InvariantCulture);

            if (root.Element("Agent") != null || !config.Agent.IsDefault)
            {
                var agent = Child(root, "Agent", RootOrder);
                var values = config.Agent.ToElements().ToList();
                var order = values.Select(e => e.Name.LocalName).ToArray();
                foreach (var value in values)
                {
                    Child(agent, value.Name.LocalName, order).Value = value.Value;
                }
            }

            Child(root, "Models", RootOrder).ReplaceNodes(
                config.Models.Select(m => new XElement("Model", new XAttribute("id", m.Id))));
            Child(root, "Functions", RootOrder).ReplaceNodes(
                config.Functions.Select(f => new XElement(
                    "Function",
                    new XAttribute("name", f.Name),
                    new XAttribute(RemoveBlankLinesAttribute, f.RemoveExtraBlankLines ? "true" : "false"),
                    new XElement("Prompt", f.Prompt))));
        }

        /// <summary>
        /// parent 下名为 name 的子节点；没有就新建，放在 order 里排在它前面、已经存在的最后一项后面，
        /// 前面一项都没有时放在最前。
        /// </summary>
        private static XElement Child(XElement parent, string name, IReadOnlyList<string> order)
        {
            var existing = parent.Element(name);
            if (existing != null)
            {
                return existing;
            }

            var created = new XElement(name);
            var index = order.ToList().IndexOf(name);
            var previous = order.Take(Math.Max(0, index)).Reverse()
                .Select(n => parent.Elements(n).LastOrDefault())
                .FirstOrDefault(e => e != null);
            if (previous != null)
            {
                previous.AddAfterSelf(created);
            }
            else
            {
                parent.AddFirst(created);
            }

            return created;
        }

        private static string Trim(string value) => value?.Trim() ?? string.Empty;

        private static int ReadInt(XElement element, int fallback, int min, int max)
        {
            if (!int.TryParse(Trim((string)element), NumberStyles.Integer, CultureInfo.InvariantCulture, out var value))
            {
                return fallback;
            }

            return value < min ? min : value > max ? max : value;
        }
    }
}
