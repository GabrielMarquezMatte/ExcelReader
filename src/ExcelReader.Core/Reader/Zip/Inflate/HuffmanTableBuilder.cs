// Ported from libdeflate's decompressor (MIT, Eric Biggers); the notice is in THIRD-PARTY-NOTICES.md.
using System.Numerics;

namespace ExcelReader.Core.Reader.Zip.Inflate
{
    internal static class HuffmanTableBuilder
    {
        internal static bool TryBuild(
            Span<uint> table, ReadOnlySpan<byte> lens, ReadOnlySpan<uint> results, int tableBits, int maxLen,
            Span<ushort> sorted, bool shrink, out int tableBitsUsed)
        {
            Span<int> lenCounts = stackalloc int[InflateFormat.MaxCodewordLength + 2];
            lenCounts.Clear();
            foreach (byte len in lens)
            {
                lenCounts[len]++;
            }
            while (maxLen > 1 && lenCounts[maxLen] == 0)
            {
                maxLen--;
            }
            if (shrink)
            {
                tableBits = Math.Min(tableBits, maxLen);
            }
            tableBitsUsed = tableBits;

            uint codespaceUsed = SortSymbols(lens, lenCounts, maxLen, sorted);
            ReadOnlySpan<ushort> used = sorted[lenCounts[0]..lens.Length];
            uint codespace = 1U << maxLen;
            if (codespaceUsed > codespace)
            {
                return false;
            }
            if (codespaceUsed < codespace)
            {
                return TryFillIncomplete(table[..(1 << tableBits)], results, used, codespaceUsed, codespace, lenCounts[1]);
            }
            FillComplete(table, results, used, lenCounts, tableBits);
            return true;
        }

        private static uint SortSymbols(ReadOnlySpan<byte> lens, ReadOnlySpan<int> lenCounts, int maxLen, Span<ushort> sorted)
        {
            Span<int> offsets = stackalloc int[InflateFormat.MaxCodewordLength + 2];
            offsets[0] = 0;
            offsets[1] = lenCounts[0];
            uint codespaceUsed = 0;
            for (int len = 1; len < maxLen; len++)
            {
                offsets[len + 1] = offsets[len] + lenCounts[len];
                codespaceUsed = (codespaceUsed << 1) + (uint)lenCounts[len];
            }
            codespaceUsed = (codespaceUsed << 1) + (uint)lenCounts[maxLen];
            for (int symbol = 0; symbol < lens.Length; symbol++)
            {
                sorted[offsets[lens[symbol]]++] = (ushort)symbol;
            }
            return codespaceUsed;
        }

        // Deflate allows two incomplete codes: no codewords at all, and a single one-bit codeword.
        private static bool TryFillIncomplete(
            Span<uint> mainTable, ReadOnlySpan<uint> results, ReadOnlySpan<ushort> used, uint codespaceUsed, uint codespace, int oneBitCodewords)
        {
            int symbol = 0;
            if (codespaceUsed != 0)
            {
                if (codespaceUsed != codespace >> 1 || oneBitCodewords != 1)
                {
                    return false;
                }
                symbol = used[0];
            }
            mainTable.Fill(Entry(results[symbol], 1));
            return true;
        }

        private static void FillComplete(Span<uint> table, ReadOnlySpan<uint> results, ReadOnlySpan<ushort> used, ReadOnlySpan<int> lenCounts, int tableBits)
        {
            int next = 0;
            uint codeword = 0;
            int len = 1;
            while (lenCounts[len] == 0)
            {
                len++;
            }
            int count = lenCounts[len];
            int tableEnd = 1 << len;
            while (len <= tableBits)
            {
                do
                {
                    table[(int)codeword] = Entry(results[used[next++]], len);
                    if (codeword == (uint)tableEnd - 1)
                    {
                        Replicate(table, tableEnd, 1 << tableBits);
                        return;
                    }
                    codeword = NextCodeword(codeword, (uint)tableEnd - 1);
                }
                while (--count != 0);

                do
                {
                    len++;
                    if (len <= tableBits)
                    {
                        Replicate(table, tableEnd, tableEnd << 1);
                        tableEnd <<= 1;
                    }
                    count = lenCounts[len];
                }
                while (count == 0);
            }
            FillSubtables(table, results, used[next..], lenCounts, tableBits, codeword, len, count);
        }

        private static void FillSubtables(
            Span<uint> table, ReadOnlySpan<uint> results, ReadOnlySpan<ushort> used, ReadOnlySpan<int> lenCounts,
            int tableBits, uint codeword, int len, int count)
        {
            uint tableEnd = 1U << tableBits;
            uint mainMask = tableEnd - 1;
            uint prefix = uint.MaxValue;
            uint start = 0;
            int next = 0;
            while (true)
            {
                if ((codeword & mainMask) != prefix)
                {
                    prefix = codeword & mainMask;
                    start = tableEnd;
                    int subtableBits = SubtableBits(lenCounts, tableBits, len, count);
                    tableEnd = start + (1U << subtableBits);
                    table[(int)prefix] = (start << 16) | InflateFormat.Exceptional | InflateFormat.SubtablePointer
                        | ((uint)subtableBits << 8) | (uint)tableBits;
                }

                int subLen = len - tableBits;
                uint entry = Entry(results[used[next++]], subLen);
                for (uint i = start + (codeword >> tableBits); i < tableEnd; i += 1U << subLen)
                {
                    table[(int)i] = entry;
                }

                uint allOnes = (1U << len) - 1;
                if (codeword == allOnes)
                {
                    return;
                }
                codeword = NextCodeword(codeword, allOnes);
                count--;
                while (count == 0)
                {
                    count = lenCounts[++len];
                }
            }
        }

        private static int SubtableBits(ReadOnlySpan<int> lenCounts, int tableBits, int len, int count)
        {
            int bits = len - tableBits;
            uint codespaceUsed = (uint)count;
            while (codespaceUsed < 1U << bits)
            {
                bits++;
                codespaceUsed = (codespaceUsed << 1) + (uint)lenCounts[tableBits + bits];
            }
            return bits;
        }

        private static void Replicate(Span<uint> table, int filled, int target)
        {
            while (filled < target)
            {
                table[..filled].CopyTo(table[filled..]);
                filled <<= 1;
            }
        }

        // Codewords are stored bit-reversed, so the next one in canonical order is a reversed increment.
        private static uint NextCodeword(uint codeword, uint allOnes)
        {
            uint bit = 1U << BitOperations.Log2(codeword ^ allOnes);
            return (codeword & (bit - 1)) | bit;
        }

        private static uint Entry(uint result, int len)
        {
            return result + ((uint)len << 8) + (uint)len;
        }
    }
}
