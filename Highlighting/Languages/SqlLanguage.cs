using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OneNoteCodeHelper.Highlighting.Languages
{
    /// <summary>
    /// SQL 词法着色，兼顾常见方言（SQL Server、MySQL、PostgreSQL、Oracle、SQLite）。整个语言不区分大小写。
    /// 用引号或方括号包起来的标识符（"order"、[Order]、`order`）是名字而不是关键字，整段按普通文字处理。
    /// </summary>
    internal sealed class SqlLanguage : ILanguage
    {
        private static readonly HashSet<string> Keywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "add", "all", "alter", "and", "any", "as", "asc", "begin", "between", "by", "cascade", "case",
            "check", "column", "commit", "constraint", "create", "cross", "database", "declare", "default",
            "delete", "desc", "distinct", "drop", "else", "end", "escape", "except", "exec", "execute", "exists",
            "fetch", "for", "foreign", "from", "full", "function", "grant", "group", "having", "if", "ilike", "in",
            "index", "inner", "insert", "intersect", "into", "is", "join", "key", "left", "like", "limit",
            "merge", "natural", "not", "of", "offset", "on", "or", "order", "outer", "over", "partition",
            "primary", "procedure", "proc", "references", "rename", "replace", "return", "returns", "revoke",
            "right", "rollback", "row", "rows", "schema", "select", "set", "table", "then", "top", "transaction",
            "tran", "trigger", "truncate", "union", "unique", "update", "use", "using", "values", "view", "when",
            "where", "while", "with", "recursive", "returning", "conflict", "do", "nothing", "go", "print",
            "identity", "auto_increment", "autoincrement", "temporary", "temp", "materialized", "language",
            "window", "range", "unbounded", "preceding", "following", "current", "lateral", "pivot", "unpivot",
            "output", "inserted", "deleted", "nocount", "try", "catch", "throw", "raiserror", "cursor", "open",
            "close", "deallocate", "next", "only", "first", "last", "nulls", "collate", "extension", "sequence",
            "type", "enum", "comment", "explain", "analyze", "vacuum", "pragma", "show", "describe", "to", "role",
            "user", "privileges", "admin", "option", "no", "action", "restrict", "enable", "disable", "each",
            "before", "after", "instead", "immediate", "loop", "exit", "continue", "elsif", "raise",
            "notice", "exception", "perform", "call", "within", "filter", "similar", "some", "at", "zone"
        };

        private static readonly HashSet<string> Literals = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "null", "true", "false", "unknown"
        };

        private static readonly HashSet<string> Types = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "bigint", "binary", "bit", "blob", "bool", "boolean", "bytea", "char", "character", "clob", "date",
            "datetime", "datetime2", "datetimeoffset", "dec", "decimal", "double", "float", "int", "integer",
            "interval", "json", "jsonb", "long", "mediumint", "money", "nchar", "ntext", "number", "numeric",
            "nvarchar", "nvarchar2", "real", "serial", "bigserial", "smallint", "smalldatetime", "text", "time",
            "timestamp", "timestamptz", "tinyint", "uniqueidentifier", "uuid", "varbinary", "varchar", "varchar2",
            "xml", "precision", "varying", "unsigned", "image", "rowversion", "smallserial", "citext", "inet"
        };

        /// <summary>内置函数。很多也会被当成列名（count、date、left），所以只在紧跟 ( 时才按函数着色。</summary>
        private static readonly HashSet<string> Functions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "abs", "avg", "cast", "ceiling", "ceil", "char_length", "charindex", "coalesce", "concat",
            "concat_ws", "convert", "count", "current_date", "current_timestamp", "dateadd", "datediff",
            "datepart", "date_trunc", "date_format", "dense_rank", "extract", "first_value", "floor", "format",
            "getdate", "getutcdate", "group_concat", "string_agg", "array_agg", "json_agg", "ifnull", "iif",
            "isnull", "lag", "last_value", "lead", "left", "len", "length", "lower", "ltrim", "max", "min",
            "mod", "now", "ntile", "nullif", "nvl", "nvl2", "power", "rank", "replace", "right", "round",
            "row_number", "rtrim", "substr", "substring", "sum", "sysdate", "to_char", "to_date", "to_number",
            "trim", "upper", "year", "month", "day", "newid", "scope_identity", "object_id", "db_name",
            "random", "rand", "sqrt", "stuff", "reverse", "datalength", "try_cast", "try_convert", "json_value",
            "json_extract", "regexp_replace", "split_part", "generate_series", "unnest", "greatest", "least",
            "decode", "listagg", "percentile_cont", "any_value"
        };

        /// <summary>这些词后面跟的是表名/对象名。</summary>
        private static readonly HashSet<string> TableIntroducers = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "table", "into", "view", "index", "on", "references", "exists", "update", "join", "from", "trigger"
        };

        private const string PunctuationChars = "(),;.";

        public string Id => "sql";

        public string DisplayName => "SQL";

        public IEnumerable<Token> Tokenize(string source)
        {
            var c = new LexerCursor(source);

            while (!c.AtEnd)
            {
                var start = c.Position;
                var ch = c.Current;

                if (ch == '-' && c.Peek() == '-')
                {
                    c.SkipToLineEnd();
                    c.Emit(TokenKind.Comment, start);
                    continue;
                }

                if (ch == '/' && c.Peek() == '*')
                {
                    CommonScanners.ScanBlockComment(c);
                    c.Emit(TokenKind.Comment, start);
                    continue;
                }

                if (ch == '\'')
                {
                    ScanSingleQuoted(c);
                    c.Emit(TokenKind.String, start);
                    continue;
                }

                // "name"、[name]、`name` 是带引号的标识符
                if (ch == '"' || ch == '[' || ch == '`')
                {
                    var end = FindOnSameLine(source, ch == '[' ? ']' : ch, start + 1);
                    if (end > 0)
                    {
                        c.Advance(end + 1 - start);
                        c.Emit(TokenKind.Plain, start);
                        continue;
                    }
                }

                // @var、@@ROWCOUNT（SQL Server）；:name（绑定参数）；$1（PostgreSQL）
                if ((ch == '@' && (IsWordChar(c.Peek()) || c.Peek() == '@'))
                    || (ch == ':' && char.IsLetter(c.Peek()) && c.Peek(-1) != ':')
                    || (ch == '$' && char.IsDigit(c.Peek())))
                {
                    c.Advance();
                    c.SkipWhile(x => IsWordChar(x) || x == '@');
                    c.Emit(TokenKind.Variable, start);
                    continue;
                }

                if (char.IsDigit(ch) || (ch == '.' && char.IsDigit(c.Peek())))
                {
                    CommonScanners.ScanNumber(c, '\0');
                    c.Emit(TokenKind.Number, start);
                    continue;
                }

                // #temp 临时表名
                if (IsWordChar(ch) || (ch == '#' && IsWordChar(c.Peek())))
                {
                    c.Advance();
                    c.SkipWhile(IsWordChar);
                    var word = source.Substring(start, c.Position - start);

                    // N'unicode'、E'escape'、X'hex' 这类带前缀的字符串
                    if (c.Current == '\'' && word.Length == 1 && "NnEeXxBb".IndexOf(word[0]) >= 0)
                    {
                        ScanSingleQuoted(c);
                        c.Emit(TokenKind.String, start);
                        continue;
                    }

                    c.Emit(ClassifyIdentifier(c, word, start), start);
                    continue;
                }

                if (char.IsWhiteSpace(ch))
                {
                    c.SkipWhile(char.IsWhiteSpace);
                    c.Emit(TokenKind.Plain, start);
                    continue;
                }

                c.Advance();
                c.Emit(PunctuationChars.IndexOf(ch) >= 0 ? TokenKind.Punctuation : TokenKind.Operator, start);
            }

            return c.Finish();
        }

        private static TokenKind ClassifyIdentifier(LexerCursor c, string word, int start)
        {
            // 表名.列名 里点号后面的一律是名字，哪怕和关键字重名（t.order、u.user）
            if (start > 0 && c.Source[start - 1] == '.')
            {
                return TokenKind.Plain;
            }

            if (Literals.Contains(word))
            {
                return TokenKind.Literal;
            }

            if (Types.Contains(word))
            {
                return TokenKind.Type;
            }

            // CREATE TABLE users (...)、INSERT INTO t (a, b) 里的是表名，不是函数调用
            var callsFunction = c.NextNonWhitespaceIs('(')
                                && !TableIntroducers.Contains(CommonScanners.PreviousWord(c, start));
            if (callsFunction && Functions.Contains(word))
            {
                return TokenKind.Builtin;
            }

            if (Keywords.Contains(word))
            {
                return TokenKind.Keyword;
            }

            return callsFunction ? TokenKind.Function : TokenKind.Plain;
        }

        /// <summary>
        /// 在同一行里找结束符，找不到返回 -1。碰到换行就停：往后整篇去找的话，
        /// 没配对的引号或方括号一多，每个都要扫到文件末尾。
        /// </summary>
        private static int FindOnSameLine(string source, char close, int from)
        {
            for (var i = from; i < source.Length; i++)
            {
                var ch = source[i];
                if (ch == close)
                {
                    return i;
                }

                if (ch == '\n' || ch == '\r')
                {
                    return -1;
                }
            }

            return -1;
        }

        /// <summary>单引号字符串，'' 表示一个单引号，可以跨行。</summary>
        private static void ScanSingleQuoted(LexerCursor c)
        {
            c.Advance();
            while (!c.AtEnd)
            {
                if (c.Current == '\'')
                {
                    if (c.Peek() == '\'')
                    {
                        c.Advance(2);
                        continue;
                    }

                    c.Advance();
                    return;
                }

                c.Advance();
            }
        }

        private static bool IsWordChar(char c) => char.IsLetterOrDigit(c) || c == '_' || c == '$';

        public int ScoreLikelihood(DetectionSample sample)
        {
            // 在去掉字符串和注释的文本上匹配：别的语言字符串里拼的 SQL 不能算
            var code = sample.Code;
            var score = 0;
            score += 4 * Count(code,
                @"\b(insert\s+into|delete\s+from|create\s+(or\s+replace\s+)?(table|view|index|procedure|function|trigger|database|schema)" +
                @"|alter\s+table|drop\s+(table|view|index)|truncate\s+table)\b");
            score += 3 * Count(code, @"\b(inner|left|right|full|cross)\s+(outer\s+)?join\b");
            score += 3 * Count(code, @"\b(group|order|partition)\s+by\b");
            score += 3 * Count(code, @"\bselect\s+(distinct\s+|top\s*\(?\d+\)?\s+)?(\*|[\w.@\[\]""`]+\s*(,|\bas\b|\bfrom\b|$))");
            score += 2 * Count(code, @"\bwhere\s+[\w.\[\]""`]+\s*(=|<>|!=|<|>|\bin\b|\blike\b|\bis\b|\bbetween\b)");
            score += 3 * Count(code, @"\bupdate\s+[\w.\[\]""`]+\s+set\b");
            score += 3 * Count(code, @"\b(primary|foreign)\s+key\b");
            score += 2 * Count(code, @"\b(varchar|nvarchar|bigint|decimal|datetime)\s*\(");
            score += 2 * Count(code, @"\bvalues\s*\(");

            // 和 Lua 的注释打分对冲：-- 注释两边都有，不该由它决定结果。同样排除 CSS 自定义属性那种行。
            score += 3 * Count(sample.Raw, @"^[^\S\r\n]*--(?![\w-]+\s*:.*;\s*$)");
            return score;
        }

        private static int Count(string source, string pattern)
        {
            return LikelihoodPatterns.Count(source, pattern, RegexOptions.Multiline | RegexOptions.IgnoreCase);
        }
    }
}
