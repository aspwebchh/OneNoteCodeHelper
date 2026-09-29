using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OneNoteCodeHelper.Highlighting.Languages
{
    /// <summary>
    /// PHP 词法着色。源码里有 &lt;?php 或 &lt;?= 时按模板处理：标签外的内容交给 HTML 着色，标签里的按 PHP 着色；
    /// 没有标签的片段整段都是 PHP。双引号字符串里的 $name 按变量着色，{$expr} 里的表达式按 PHP 重新着色；
    /// heredoc 整段按字符串着色。关键字不区分大小写。
    /// 已知限制：&lt;a href="&lt;?= $url ?&gt;"&gt; 这种把 HTML 属性切成两半的写法，前后两段 HTML 各自着色，颜色会有偏差。
    /// </summary>
    internal sealed class PhpLanguage : ILanguage
    {
        private static readonly ILanguage Html = new HtmlLanguage();

        private static readonly HashSet<string> Keywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "abstract", "and", "as", "break", "callable", "case", "catch", "class", "clone", "const", "continue",
            "declare", "default", "do", "echo", "else", "elseif", "empty", "enddeclare", "endfor", "endforeach",
            "endif", "endswitch", "endwhile", "enum", "extends", "final", "finally", "fn", "for", "foreach",
            "function", "global", "goto", "if", "implements", "include", "include_once", "instanceof", "insteadof",
            "interface", "isset", "list", "match", "namespace", "new", "or", "print", "private", "protected",
            "public", "readonly", "require", "require_once", "return", "static", "switch", "throw", "trait", "try",
            "unset", "use", "var", "while", "xor", "yield"
        };

        private static readonly HashSet<string> Literals = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "true", "false", "null"
        };

        /// <summary>类型声明里的内置类型，以及 self::、parent:: 这类类名占位。</summary>
        private static readonly HashSet<string> Types = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "int", "float", "string", "bool", "array", "void", "mixed", "never", "iterable", "object", "self",
            "parent"
        };

        private static readonly HashSet<string> TypeIntroducers = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "class", "interface", "trait", "enum", "extends", "implements", "new", "instanceof", "insteadof",
            "namespace", "use"
        };

        private const string PunctuationChars = "(){}[];,.";

        public string Id => "php";

        public string DisplayName => "PHP";

        public IEnumerable<Token> Tokenize(string source)
        {
            var c = new LexerCursor(source);
            var inPhp = FindOpenTag(c.Source, 0) < 0;

            while (!c.AtEnd)
            {
                if (!inPhp)
                {
                    var open = FindOpenTag(source, c.Position);
                    c.EmitEmbedded(Html, open < 0 ? source.Length : open);
                    if (open < 0)
                    {
                        break;
                    }

                    var tagStart = c.Position;
                    c.Advance(c.MatchesIgnoreCase("<?php") ? 5 : 3);
                    c.Emit(TokenKind.Annotation, tagStart);
                    inPhp = true;
                    continue;
                }

                var start = c.Position;
                var ch = c.Current;

                if (ch == '?' && c.Peek() == '>')
                {
                    c.Advance(2);
                    c.Emit(TokenKind.Annotation, start);
                    inPhp = false;
                    continue;
                }

                // 属性 #[Route('/users')]：#[ 和类名按注解着色，括号里的照常着色
                if (ch == '#' && c.Peek() == '[')
                {
                    c.Advance(2);
                    c.SkipWhile(x => IsIdentifierPart(x) || x == '\\');
                    c.Emit(TokenKind.Annotation, start);
                    continue;
                }

                // 单行注释遇到 ?> 就结束，这是 PHP 本身的规则
                if ((ch == '/' && c.Peek() == '/') || ch == '#')
                {
                    while (!c.AtEnd && c.Current != '\n' && c.Current != '\r' && !(c.Current == '?' && c.Peek() == '>'))
                    {
                        c.Advance();
                    }

                    c.Emit(TokenKind.Comment, start);
                    continue;
                }

                if (ch == '/' && c.Peek() == '*')
                {
                    var isDoc = CommonScanners.IsDocBlockComment(c);
                    CommonScanners.ScanBlockComment(c);
                    c.Emit(isDoc ? TokenKind.DocComment : TokenKind.Comment, start);
                    continue;
                }

                // $name、$$name（可变变量）；$this 按关键字着色
                if (ch == '$')
                {
                    c.SkipWhile(x => x == '$');
                    if (IsIdentifierStart(c.Current))
                    {
                        c.SkipWhile(IsIdentifierPart);
                        var isThis = c.Position - start == 5 && string.CompareOrdinal(source, start, "$this", 0, 5) == 0;
                        c.Emit(isThis ? TokenKind.Keyword : TokenKind.Variable, start);
                    }
                    else
                    {
                        c.Emit(TokenKind.Operator, start);
                    }

                    continue;
                }

                if (ch == '\'' || ch == '`')
                {
                    CommonScanners.ScanQuoted(c, ch, multiLine: true);
                    c.Emit(TokenKind.String, start);
                    continue;
                }

                if (ch == '"')
                {
                    c.Advance();
                    CommonScanners.ScanInterpolatedBody(
                        c, this, start, "\"", '\\', "{$", multiLine: true, dollarVariables: true);
                    continue;
                }

                if (ch == '<' && c.Matches("<<<") && TryScanHeredoc(c))
                {
                    continue;
                }

                if (char.IsDigit(ch) || (ch == '.' && char.IsDigit(c.Peek())))
                {
                    CommonScanners.ScanNumber(c, '_');
                    c.Emit(TokenKind.Number, start);
                    continue;
                }

                if (IsIdentifierStart(ch))
                {
                    c.SkipWhile(IsIdentifierPart);
                    var word = source.Substring(start, c.Position - start);
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
            var source = c.Source;

            // $obj->name、$obj?->name 是成员
            if (start >= 2 && source[start - 1] == '>' && source[start - 2] == '-')
            {
                return c.NextNonWhitespaceIs('(') ? TokenKind.Function : TokenKind.Plain;
            }

            // Foo::bar() 里的 bar 不是关键字，Foo::class 里的 class 是
            var afterStatic = start >= 2 && source[start - 1] == ':' && source[start - 2] == ':';

            if (Keywords.Contains(word) && (!afterStatic || string.Equals(word, "class", StringComparison.OrdinalIgnoreCase)))
            {
                return TokenKind.Keyword;
            }

            if (Literals.Contains(word))
            {
                return TokenKind.Literal;
            }

            if (!afterStatic && Types.Contains(word))
            {
                return TokenKind.Type;
            }

            var previousWord = CommonScanners.PreviousWord(c, start);
            if (string.Equals(previousWord, "function", StringComparison.OrdinalIgnoreCase))
            {
                return TokenKind.Function;
            }

            if (TypeIntroducers.Contains(previousWord))
            {
                return TokenKind.Type;
            }

            if (c.NextNonWhitespaceIs('('))
            {
                return TokenKind.Function;
            }

            if (!char.IsUpper(word[0]))
            {
                return TokenKind.Plain;
            }

            return CommonScanners.IsAllCaps(word) ? TokenKind.Constant : TokenKind.Type;
        }

        /// <summary>
        /// heredoc &lt;&lt;&lt;EOT、&lt;&lt;&lt;"EOT" 与 nowdoc &lt;&lt;&lt;'EOT'，整段按字符串着色。结束标记单独起一行，
        /// PHP 7.3 起可以缩进、后面可以直接跟 ; , )。不是 heredoc 就原地不动返回 false。
        /// </summary>
        private static bool TryScanHeredoc(LexerCursor c)
        {
            var source = c.Source;
            var i = c.Position + 3;
            while (i < source.Length && (source[i] == ' ' || source[i] == '\t'))
            {
                i++;
            }

            var quote = i < source.Length && (source[i] == '\'' || source[i] == '"') ? source[i] : '\0';
            if (quote != '\0')
            {
                i++;
            }

            var idStart = i;
            if (i >= source.Length || !IsIdentifierStart(source[i]))
            {
                return false;
            }

            while (i < source.Length && IsIdentifierPart(source[i]))
            {
                i++;
            }

            var id = source.Substring(idStart, i - idStart);
            if (quote != '\0')
            {
                if (i >= source.Length || source[i] != quote)
                {
                    return false;
                }

                i++;
            }

            if (i < source.Length && source[i] != '\n' && source[i] != '\r')
            {
                return false;
            }

            var start = c.Position;
            c.Advance(FindHeredocEnd(source, i, id) - start);
            c.Emit(TokenKind.String, start);
            return true;
        }

        /// <summary>从开头那行的行尾往后找结束标记所在的行，返回标记之后的位置；找不到返回源码末尾。</summary>
        private static int FindHeredocEnd(string source, int position, string id)
        {
            while (true)
            {
                var newline = source.IndexOf('\n', position);
                if (newline < 0)
                {
                    return source.Length;
                }

                var p = newline + 1;
                while (p < source.Length && (source[p] == ' ' || source[p] == '\t'))
                {
                    p++;
                }

                var end = p + id.Length;
                if (end <= source.Length && string.CompareOrdinal(source, p, id, 0, id.Length) == 0
                    && (end == source.Length || !IsIdentifierPart(source[end])))
                {
                    return end;
                }

                position = p;
            }
        }

        /// <summary>from 之后第一个 &lt;?php 或 &lt;?= 的位置，没有返回 -1。</summary>
        private static int FindOpenTag(string source, int from)
        {
            var php = source.IndexOf("<?php", from, StringComparison.OrdinalIgnoreCase);
            var echo = source.IndexOf("<?=", from, StringComparison.Ordinal);
            if (php < 0 || echo < 0)
            {
                return Math.Max(php, echo);
            }

            return Math.Min(php, echo);
        }

        private static bool IsIdentifierStart(char c) => char.IsLetter(c) || c == '_';

        private static bool IsIdentifierPart(char c) => char.IsLetterOrDigit(c) || c == '_';

        /// <summary>
        /// 不调 <see cref="CFamilyFeatures.Score"/>：PHP 不在 C 系族里，共有的 class、public、分号那部分分要是也记给 PHP，
        /// 只靠共有特征取胜的 Java/C# 片段就会和 PHP 打平、变成无法确定。PHP 靠 $ 变量相关的写法取胜。
        /// </summary>
        public int ScoreLikelihood(DetectionSample sample)
        {
            var code = sample.Code;
            var openTags = Count(sample.Raw, @"<\?(php\b|=)", RegexOptions.IgnoreCase);
            var score = 10 * openTags;
            score += 4 * Count(code, @"\$this->");
            score += 3 * Count(code, @"\$\w+(->|::)\w");
            score += 4 * Count(code, @"\bfunction[^\S\r\n]+&?\w+[^\S\r\n]*\([^)\r\n]*\$\w+");
            score += 3 * Count(code, @"\b(public|private|protected)[^\S\r\n]+(static[^\S\r\n]+)?function\b");
            score += 5 * Count(code, @"^[^\S\r\n]*namespace[^\S\r\n]+\w+(\\\w+)*[^\S\r\n]*;");
            score += 5 * Count(code, @"^[^\S\r\n]*use[^\S\r\n]+\\?\w+(\\\w+)+[^\S\r\n]*(as[^\S\r\n]+\w+[^\S\r\n]*)?;");
            score += 5 * Count(code, @"\bforeach[^\S\r\n]*\([^\S\r\n]*\$\w+[^\S\r\n]+as\b");
            score += 5 * Count(code, @"\$_(GET|POST|SERVER|SESSION|REQUEST|COOKIE|FILES|ENV)\b");
            score += 2 * Count(code, @"['""][^'""\r\n]*['""][^\S\r\n]*=>");
            score += 2 * Count(code,
                @"\.=|(?<![.\w$])(isset|empty|unset|in_array|array_\w+|str_\w+|implode|explode|json_encode|json_decode" +
                @"|strtolower|strtoupper|ucfirst|trim|rtrim|ltrim|substr|strpos|preg_\w+|htmlspecialchars)[^\S\r\n]*\(");
            score += 2 * Count(code, @"\$\w+[^\r\n]*;[^\S\r\n]*$");

            // 含 PHP 标签的 HTML 就是 PHP 模板：HTML 那部分的分也记给 PHP，才比得过 HTML
            if (openTags > 0)
            {
                score += Html.ScoreLikelihood(sample);
            }

            return score;
        }

        private static int Count(string source, string pattern, RegexOptions options = RegexOptions.None)
        {
            return LikelihoodPatterns.Count(source, pattern, RegexOptions.Multiline | options);
        }
    }
}
