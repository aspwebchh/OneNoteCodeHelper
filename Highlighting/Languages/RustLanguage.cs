using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OneNoteCodeHelper.Highlighting.Languages
{
    /// <summary>
    /// Rust 词法着色。难点在单引号：'a' 是字符，'a 是生命周期或循环标签；字符串有 b"" c"" 前缀和
    /// r#"..."# 原始字符串，按 # 的个数配对；块注释可以嵌套。宏调用 name! 按函数着色，#[attr] 按注解着色。
    /// </summary>
    internal sealed class RustLanguage : ILanguage
    {
        private static readonly HashSet<string> Keywords = new HashSet<string>(StringComparer.Ordinal)
        {
            "as", "async", "await", "break", "const", "continue", "crate", "dyn", "else", "enum", "extern", "fn",
            "for", "if", "impl", "in", "let", "loop", "match", "mod", "move", "mut", "pub", "ref", "return",
            "self", "Self", "static", "struct", "super", "trait", "type", "unsafe", "use", "where", "while",
            "yield", "union"
        };

        private static readonly HashSet<string> Literals = new HashSet<string>(StringComparer.Ordinal)
        {
            "true", "false"
        };

        private static readonly HashSet<string> PrimitiveTypes = new HashSet<string>(StringComparer.Ordinal)
        {
            "i8", "i16", "i32", "i64", "i128", "isize", "u8", "u16", "u32", "u64", "u128", "usize", "f32", "f64",
            "bool", "char", "str"
        };

        private static readonly HashSet<string> TypeIntroducers = new HashSet<string>(StringComparer.Ordinal)
        {
            "struct", "enum", "trait", "type", "union"
        };

        private const string PunctuationChars = "(){}[];,.";

        public string Id => "rust";

        public string DisplayName => "Rust";

        public IEnumerable<Token> Tokenize(string source)
        {
            var c = new LexerCursor(source);

            while (!c.AtEnd)
            {
                var start = c.Position;
                var ch = c.Current;

                if (ch == '/' && c.Peek() == '/')
                {
                    // /// 与 //! 是文档注释，//// 这种分隔线不算
                    var isDoc = (c.Peek(2) == '/' && c.Peek(3) != '/') || c.Peek(2) == '!';
                    c.SkipToLineEnd();
                    c.Emit(isDoc ? TokenKind.DocComment : TokenKind.Comment, start);
                    continue;
                }

                if (ch == '/' && c.Peek() == '*')
                {
                    var isDoc = CommonScanners.IsDocBlockComment(c) || c.Peek(2) == '!';
                    CommonScanners.ScanNestedBlockComment(c);
                    c.Emit(isDoc ? TokenKind.DocComment : TokenKind.Comment, start);
                    continue;
                }

                if (ch == '"')
                {
                    CommonScanners.ScanQuoted(c, '"', multiLine: true);
                    c.Emit(TokenKind.String, start);
                    continue;
                }

                if ((ch == 'r' || ch == 'b' || ch == 'c') && TryScanPrefixedLiteral(c))
                {
                    continue;
                }

                if (ch == '\'')
                {
                    ScanQuoteOrLifetime(c);
                    continue;
                }

                // 属性 #[derive(Debug)]、#![allow(dead_code)]：#[ 和路径名按注解着色，括号里的照常着色
                if (ch == '#' && (c.Peek() == '[' || (c.Peek() == '!' && c.Peek(2) == '[')))
                {
                    c.Advance(c.Peek() == '[' ? 2 : 3);
                    c.SkipWhile(x => IsIdentifierPart(x) || x == ':');
                    c.Emit(TokenKind.Annotation, start);
                    continue;
                }

                // 没有 .5 这种写法；t.0 是元组下标
                if (char.IsDigit(ch))
                {
                    CommonScanners.ScanNumber(c, '_');
                    c.SkipWhile(char.IsLetterOrDigit); // 类型后缀 1u32、2.0f64
                    c.Emit(TokenKind.Number, start);
                    continue;
                }

                if (IsIdentifierStart(ch))
                {
                    c.SkipWhile(IsIdentifierPart);

                    // 宏调用 println!、vec![]、macro_rules!；!= 不算
                    if (c.Current == '!' && c.Peek() != '=')
                    {
                        c.Advance();
                        c.Emit(TokenKind.Function, start);
                        continue;
                    }

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
            var afterMember = start > 0 && c.Source[start - 1] == '.';

            if (!afterMember && Keywords.Contains(word))
            {
                return TokenKind.Keyword;
            }

            if (!afterMember && Literals.Contains(word))
            {
                return TokenKind.Literal;
            }

            if (!afterMember && PrimitiveTypes.Contains(word))
            {
                return TokenKind.Type;
            }

            var previousWord = CommonScanners.PreviousWord(c, start);
            if (previousWord == "fn")
            {
                return TokenKind.Function;
            }

            if (TypeIntroducers.Contains(previousWord))
            {
                return TokenKind.Type;
            }

            // Rust 的命名规范很严格：函数和变量一律小写，大写开头的是类型、枚举成员（Some(x) 也算），全大写是常量
            if (char.IsUpper(word[0]))
            {
                return CommonScanners.IsAllCaps(word) ? TokenKind.Constant : TokenKind.Type;
            }

            // 调用，包括带类型参数的 collect::<Vec<_>>()
            return c.NextNonWhitespaceIs('(') || c.Matches("::<") ? TokenKind.Function : TokenKind.Plain;
        }

        /// <summary>
        /// r"..."、r#"..."#、b"..."、b'x'、br#"..."#、c"..."、cr#"..."#。当前位置不是这类字面量就原地不动返回 false，
        /// 按普通标识符处理（包括 r#type 这种原始标识符）。
        /// </summary>
        private static bool TryScanPrefixedLiteral(LexerCursor c)
        {
            var start = c.Position;
            var prefix = 1;
            if ((c.Current == 'b' || c.Current == 'c') && c.Peek() == 'r')
            {
                prefix = 2;
            }

            var raw = c.Peek(prefix - 1) == 'r';
            if (start > 0 && IsIdentifierPart(c.Source[start - 1]))
            {
                return false;
            }

            if (!raw)
            {
                if (c.Peek() == '"')
                {
                    c.Advance();
                    CommonScanners.ScanQuoted(c, '"', multiLine: true);
                    c.Emit(TokenKind.String, start);
                    return true;
                }

                if (c.Current == 'b' && c.Peek() == '\'')
                {
                    c.Advance();
                    CommonScanners.ScanQuoted(c, '\'');
                    c.Emit(TokenKind.Char, start);
                    return true;
                }

                return false;
            }

            var hashes = 0;
            while (c.Peek(prefix + hashes) == '#')
            {
                hashes++;
            }

            // 原始标识符 r#type：拿关键字当名字用，整体是个普通名字
            if (prefix == 1 && hashes == 1 && IsIdentifierStart(c.Peek(2)))
            {
                c.Advance(2);
                c.SkipWhile(IsIdentifierPart);
                c.Emit(TokenKind.Plain, start);
                return true;
            }

            if (c.Peek(prefix + hashes) != '"')
            {
                return false;
            }

            var terminator = "\"" + new string('#', hashes);
            c.Advance(prefix + hashes + 1);
            while (!c.AtEnd && !c.Matches(terminator))
            {
                c.Advance();
            }

            c.Advance(terminator.Length);
            c.Emit(TokenKind.String, start);
            return true;
        }

        /// <summary>'a'、'\n'、'\u{1F600}'、'😀' 是字符；'a、'static、'outer 是生命周期或循环标签。</summary>
        private static void ScanQuoteOrLifetime(LexerCursor c)
        {
            var start = c.Position;
            var next = c.Peek();
            var charLength = char.IsHighSurrogate(next) && char.IsLowSurrogate(c.Peek(2)) ? 2 : 1;

            if (next == '\\' || (next != '\0' && next != '\n' && next != '\r' && c.Peek(1 + charLength) == '\''))
            {
                CommonScanners.ScanQuoted(c, '\'');
                c.Emit(TokenKind.Char, start);
                return;
            }

            c.Advance();
            if (IsIdentifierStart(next))
            {
                c.SkipWhile(IsIdentifierPart);
                c.Emit(TokenKind.Annotation, start);
                return;
            }

            c.Emit(TokenKind.Operator, start);
        }

        private static bool IsIdentifierStart(char c) => char.IsLetter(c) || c == '_';

        private static bool IsIdentifierPart(char c) => char.IsLetterOrDigit(c) || c == '_';

        public int ScoreLikelihood(DetectionSample sample)
        {
            var code = sample.Code;
            var score = 0;
            score += 5 * Count(code, @"^[^\S\r\n]*(pub(\([\w:]+\))?[^\S\r\n]+)?(async[^\S\r\n]+|const[^\S\r\n]+|unsafe[^\S\r\n]+)*fn[^\S\r\n]+\w+[^\S\r\n]*[<(]");
            score += 5 * Count(code, @"\blet[^\S\r\n]+mut\b");
            score += 5 * Count(code, @"^[^\S\r\n]*impl\b[^\r\n]*\{[^\S\r\n]*$");
            score += 5 * Count(code, @"^[^\S\r\n]*(pub[^\S\r\n]+)?use[^\S\r\n]+(std|crate|super|self|core)::");
            score += 5 * Count(code, @"^[^\S\r\n]*#!?\[(derive|cfg|allow|test|tokio::main|inline)\b");
            score += 3 * Count(code, @"\b(println|print|eprintln|format|vec|panic|assert|assert_eq|write|writeln|macro_rules)![^\S\r\n]*[(\[{]");
            score += 3 * Count(code, @"&mut[^\S\r\n]|&self\b");

            // 生命周期要在原文上找：去噪时会把两个单引号之间当成字符串抹掉
            score += 3 * Count(sample.Raw, @"[&<][^\S\r\n]*'[a-z_]\w*\b(?!')");
            score += 3 * Count(code, @"\.unwrap\(\)|\.expect\(|\.iter\(\)|\.collect::<|\.clone\(\)");
            score += 2 * Count(code, @"\b(Some|Ok|Err)\(|\bNone\b|\bOption<|\bResult<|\bVec<|\bBox<");
            score += 2 * Count(code, @"\blet[^\S\r\n]+\w+[^\S\r\n]*(:[^=\r\n]+)?=");
            score += 2 * Count(code, @"\bmatch[^\S\r\n]+[^\r\n{]+\{[^\S\r\n]*$|\bSelf\b");
            score += 2 * Count(code, @"\)[^\S\r\n]*->[^\S\r\n]*[\w<>&'\[\], ]+(where\b|\{)");
            score += 2 * Count(code, @"\b(i32|u32|i64|u64|usize|isize|u8|f64)\b|&str\b");
            return score;
        }

        private static int Count(string source, string pattern)
        {
            return LikelihoodPatterns.Count(source, pattern, RegexOptions.Multiline);
        }
    }
}
