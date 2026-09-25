using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace OneNoteCodeHelper.Services
{
    /// <summary>功能区「模型」下拉里的一项。下拉里直接显示模型 id，也就是发给接口的 model 参数。</summary>
    internal sealed class AiModel
    {
        internal AiModel(string id)
        {
            Id = id;
        }

        internal string Id { get; }
    }

    /// <summary>功能区「功能」下拉里的一项：显示名 + 提示词，外加由插件自己做、不经过 AI 的规则。</summary>
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
    /// 思考强度（变体），功能区下拉里直接显示这些名字。取值和请求参数照搬 opencode 配置里 deepseek 模型的 variants：
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
    /// AI 助手的配置：接口地址、Key、模型和各功能的提示词。对应 ai-settings.xml，由用户手改。
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

        /// <summary>两份配置的下拉选项是否一样。不一样才需要刷新功能区，免得每次都闪一下。</summary>
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
    /// 这份文件插件只在它不存在时写一次默认值，之后只读。
    ///
    /// 用 LINQ to XML 而不是 XmlSerializer，是为了能在默认文件里写注释，告诉用户每一项怎么填。
    /// </summary>
    internal static class AiConfigStore
    {
        private const string DefaultApiUrl = "https://api.deepseek.com";

        private const int DefaultTimeoutSeconds = 300;

        private const int DefaultMaxTokens = 16384;

        internal const string TypoFunctionName = "错别字修复";

        internal const string DefaultModelId = "deepseek-v4-flash";

        private const string RemoveBlankLinesAttribute = "removeExtraBlankLines";

        internal const string SmartFunctionName = "智能校正";

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
                ReadInt(root.Element("TimeoutSeconds"), DefaultTimeoutSeconds, 10, 3600),
                ReadInt(root.Element("MaxTokens"), DefaultMaxTokens, 0, 1024 * 1024),
                models?.Count > 0 ? models : Default.Models,
                functions?.Count > 0 ? functions : Default.Functions);
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
                new XComment(" OneNote 代码高亮 · AI 助手配置。改完保存即可，下次点「AI 优化」时生效。 "),
                new XElement(
                    "AiConfig",
                    new XComment(" OpenAI 兼容接口的基础地址；插件会在后面加上 /chat/completions "),
                    new XElement("ApiUrl", config.ApiUrl),
                    new XElement("ApiKey", config.ApiKey),
                    new XComment(" 单次请求的超时（秒）。思考强度开得高、要处理的内容又多时可以调大 "),
                    new XElement("TimeoutSeconds", config.TimeoutSeconds),
                    new XComment(" 单次请求最多输出多少 token（含思考过程）。0 表示用接口的默认值 "),
                    new XElement("MaxTokens", config.MaxTokens),
                    new XComment(" 功能区「模型」下拉里的选项：id 是接口的模型名，下拉里直接显示它 "),
                    new XElement(
                        "Models",
                        config.Models.Select(m => new XElement("Model", new XAttribute("id", m.Id)))),
                    new XComment(
                        " 功能区「功能」下拉里的选项，可以自己加。Prompt 只需写清楚要做什么；\n" +
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
