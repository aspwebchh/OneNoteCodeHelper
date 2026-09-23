using System;
using System.Collections.Generic;

namespace OneNoteCodeHelper.Services
{
    internal enum DiffKind
    {
        Keep,
        Delete,
        Insert
    }

    /// <summary>一步编辑：保留或删除旧文本的第 OldIndex 个字符，或者插入 NewChar。</summary>
    internal readonly struct DiffOp
    {
        internal DiffOp(DiffKind kind, int oldIndex, char newChar)
        {
            Kind = kind;
            OldIndex = oldIndex;
            NewChar = newChar;
        }

        internal DiffKind Kind { get; }

        internal int OldIndex { get; }

        internal char NewChar { get; }
    }

    /// <summary>
    /// 逐字符的文本差异。用来把 AI 改过的段落文字合并回原来带格式的段落。
    /// </summary>
    internal static class TextDiff
    {
        /// <summary>
        /// LCS 表格的上限。先去掉公共前后缀，剩下的中间段才建表；错别字这种零星改动中间段很短，
        /// 排版优化改动分散，中间段接近整段，这个上限大约对应两千字的段落（表格约 8MB）。
        /// 超过就整段替换，代价只是那一段的行内格式可能挪位。
        /// </summary>
        private const long MaxTableCells = 4_000_000;

        /// <summary>
        /// 算出把 oldText 变成 newText 的编辑序列。同一处的删除总排在插入前面，
        /// 合并时插入的字符才能沿用被替换那个字符的格式。
        /// </summary>
        internal static List<DiffOp> Compute(string oldText, string newText)
        {
            oldText = oldText ?? string.Empty;
            newText = newText ?? string.Empty;

            var prefix = 0;
            var limit = Math.Min(oldText.Length, newText.Length);
            while (prefix < limit && oldText[prefix] == newText[prefix])
            {
                prefix++;
            }

            var suffix = 0;
            while (suffix < limit - prefix
                   && oldText[oldText.Length - 1 - suffix] == newText[newText.Length - 1 - suffix])
            {
                suffix++;
            }

            var ops = new List<DiffOp>(Math.Max(oldText.Length, newText.Length) + 8);
            for (var i = 0; i < prefix; i++)
            {
                ops.Add(new DiffOp(DiffKind.Keep, i, oldText[i]));
            }

            var oldEnd = oldText.Length - suffix;
            var newEnd = newText.Length - suffix;
            var oldCount = oldEnd - prefix;
            var newCount = newEnd - prefix;

            if ((long)(oldCount + 1) * (newCount + 1) <= MaxTableCells)
            {
                AppendLcsOps(ops, oldText, prefix, oldCount, newText, prefix, newCount);
            }
            else
            {
                for (var i = prefix; i < oldEnd; i++)
                {
                    ops.Add(new DiffOp(DiffKind.Delete, i, oldText[i]));
                }

                for (var j = prefix; j < newEnd; j++)
                {
                    ops.Add(new DiffOp(DiffKind.Insert, -1, newText[j]));
                }
            }

            for (var i = oldEnd; i < oldText.Length; i++)
            {
                ops.Add(new DiffOp(DiffKind.Keep, i, oldText[i]));
            }

            return ops;
        }

        private static void AppendLcsOps(List<DiffOp> ops,
            string a, int aStart, int aCount, string b, int bStart, int bCount)
        {
            // table[i, j] = a[i..] 与 b[j..] 的最长公共子序列长度，从后往前填。
            // 建表的前提是格子数不超过 MaxTableCells，所以较短一方不超过两千，ushort 够用，省一半内存。
            var width = bCount + 1;
            var table = new ushort[(aCount + 1) * width];

            for (var i = aCount - 1; i >= 0; i--)
            {
                for (var j = bCount - 1; j >= 0; j--)
                {
                    table[i * width + j] = a[aStart + i] == b[bStart + j]
                        ? (ushort)(table[(i + 1) * width + j + 1] + 1)
                        : Math.Max(table[(i + 1) * width + j], table[i * width + j + 1]);
                }
            }

            var x = 0;
            var y = 0;
            while (x < aCount && y < bCount)
            {
                if (a[aStart + x] == b[bStart + y])
                {
                    ops.Add(new DiffOp(DiffKind.Keep, aStart + x, a[aStart + x]));
                    x++;
                    y++;
                }
                else if (table[(x + 1) * width + y] >= table[x * width + y + 1])
                {
                    ops.Add(new DiffOp(DiffKind.Delete, aStart + x, a[aStart + x]));
                    x++;
                }
                else
                {
                    ops.Add(new DiffOp(DiffKind.Insert, -1, b[bStart + y]));
                    y++;
                }
            }

            for (; x < aCount; x++)
            {
                ops.Add(new DiffOp(DiffKind.Delete, aStart + x, a[aStart + x]));
            }

            for (; y < bCount; y++)
            {
                ops.Add(new DiffOp(DiffKind.Insert, -1, b[bStart + y]));
            }
        }
    }
}
