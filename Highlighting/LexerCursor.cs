using System;
using System.Collections.Generic;

namespace OneNoteCodeHelper.Highlighting
{
    /// <summary>
    /// 各语言词法分析器共用的字符游标 + token 收集器。
    ///
    /// 它同时负责保证「token 完整覆盖源码」这条契约：调用方只需在识别出有意义的片段时
    /// 调 <see cref="Emit"/>，中间被跳过的字符由 <see cref="Emit"/> 自动补成 Plain。
    /// </summary>
    internal sealed class LexerCursor
    {
        private readonly List<Token> _tokens = new List<Token>();
        private int _emitted;

        internal LexerCursor(string source)
        {
            Source = source ?? string.Empty;
        }

        internal string Source { get; }

        internal int Position { get; private set; }

        internal bool AtEnd => Position >= Source.Length;

        internal char Current => Position < Source.Length ? Source[Position] : '\0';

        internal char Peek(int offset = 1)
        {
            var index = Position + offset;
            return index >= 0 && index < Source.Length ? Source[index] : '\0';
        }

        internal void Advance(int count = 1)
        {
            Position = Math.Min(Position + count, Source.Length);
        }

        internal bool Matches(string text)
        {
            return string.CompareOrdinal(Source, Position, text, 0, text.Length) == 0
                   && Position + text.Length <= Source.Length;
        }

        /// <summary>不区分大小写的 <see cref="Matches"/>，给 Bat、HTML 这类大小写不敏感的语言用。</summary>
        internal bool MatchesIgnoreCase(string text)
        {
            return Position + text.Length <= Source.Length
                   && string.Compare(Source, Position, text, 0, text.Length, StringComparison.OrdinalIgnoreCase) == 0;
        }

        /// <summary>当前位置之前、同一行内是不是只有空白（或者干脆就在行首）。</summary>
        internal bool IsAtLineStart()
        {
            for (var i = Position - 1; i >= 0; i--)
            {
                var ch = Source[i];
                if (ch == '\n' || ch == '\r')
                {
                    return true;
                }

                if (ch != ' ' && ch != '\t')
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// 把 [Position, end) 交给另一种语言切分，token 平移后原样收下，游标移到 end。
        /// 用于 HTML 里的 &lt;style&gt; 这类内嵌代码；内层语言自己保证完整覆盖，所以契约不受影响。
        /// </summary>
        internal void EmitEmbedded(ILanguage inner, int end)
        {
            end = Math.Max(Position, Math.Min(end, Source.Length));
            if (Position > _emitted)
            {
                _tokens.Add(new Token(TokenKind.Plain, _emitted, Position - _emitted));
            }

            var offset = Position;
            foreach (var token in inner.Tokenize(Source.Substring(offset, end - offset)))
            {
                if (token.Length > 0)
                {
                    _tokens.Add(new Token(token.Kind, offset + token.Start, token.Length));
                }
            }

            Position = end;
            _emitted = end;
        }

        /// <summary>把从 start 到当前位置的区间标成指定类别；中间的空档自动补 Plain。</summary>
        internal void Emit(TokenKind kind, int start)
        {
            if (start > _emitted)
            {
                _tokens.Add(new Token(TokenKind.Plain, _emitted, start - _emitted));
            }

            if (Position > start)
            {
                _tokens.Add(new Token(kind, start, Position - start));
            }

            _emitted = Position;
        }

        /// <summary>吃掉满足条件的连续字符。</summary>
        internal void SkipWhile(Func<char, bool> predicate)
        {
            while (!AtEnd && predicate(Current))
            {
                Position++;
            }
        }

        /// <summary>吃到行尾（不含换行符本身）。</summary>
        internal void SkipToLineEnd()
        {
            SkipWhile(c => c != '\n' && c != '\r');
        }

        /// <summary>
        /// 从当前位置向后看，跳过空白后的第一个非空白字符是不是 c。
        /// 用来判断标识符后面是不是紧跟左括号（即「这是个函数名」）。
        /// </summary>
        internal bool NextNonWhitespaceIs(char c)
        {
            for (var i = Position; i < Source.Length; i++)
            {
                if (!char.IsWhiteSpace(Source[i]))
                {
                    return Source[i] == c;
                }
            }

            return false;
        }

        /// <summary>收尾：补齐末尾残留的 Plain，返回完整 token 序列。</summary>
        internal IReadOnlyList<Token> Finish()
        {
            if (Source.Length > _emitted)
            {
                _tokens.Add(new Token(TokenKind.Plain, _emitted, Source.Length - _emitted));
                _emitted = Source.Length;
            }

            return _tokens;
        }

        internal static bool IsIdentifierStart(char c) => char.IsLetter(c) || c == '_' || c == '$';

        internal static bool IsIdentifierPart(char c) => char.IsLetterOrDigit(c) || c == '_' || c == '$';

        internal static bool IsHexDigit(char c)
        {
            return (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F');
        }
    }
}
