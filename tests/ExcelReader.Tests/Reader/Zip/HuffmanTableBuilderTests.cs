using ExcelReader.Core.Reader.Zip.Inflate;

namespace ExcelReader.Tests.Reader.Zip
{
    public class HuffmanTableBuilderTests
    {
        private static readonly uint[] _symbolResults = [.. Enumerable.Range(0, InflateFormat.MaxLitlenSymbols).Select(static i => (uint)i << 16)];

        public static TheoryData<byte[], int> CompleteCodes()
        {
            byte[] fixedLitlen = new byte[288];
            fixedLitlen.AsSpan(0, 144).Fill(8);
            fixedLitlen.AsSpan(144, 112).Fill(9);
            fixedLitlen.AsSpan(256, 24).Fill(7);
            fixedLitlen.AsSpan(280, 8).Fill(8);
            return new TheoryData<byte[], int>
            {
                { [3, 3, 3, 3, 3, 2, 4, 4], 11 },
                { [3, 3, 3, 3, 3, 2, 4, 4], 3 },
                { fixedLitlen, 11 },
                { fixedLitlen, 8 },
                { [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 15], 11 },
                { [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 15], 8 },
                { [0, 2, 0, 2, 2, 0, 0, 2], 8 },
            };
        }

        [Theory]
        [MemberData(nameof(CompleteCodes))]
        public void EveryCanonicalCodewordDecodesToItsSymbol(byte[] lens, int tableBits)
        {
            uint[] table = new uint[InflateFormat.LitlenTableSize];
            ushort[] sorted = new ushort[lens.Length];

            Assert.True(HuffmanTableBuilder.TryBuild(table, lens, _symbolResults, tableBits, 15, sorted, shrink: false, out int used));
            Assert.Equal(tableBits, used);

            uint[] codewords = CanonicalCodewords(lens);
            int checkedSymbols = 0;
            for (int symbol = 0; symbol < lens.Length; symbol++)
            {
                if (lens[symbol] == 0)
                {
                    continue;
                }
                uint reversed = Reverse(codewords[symbol], lens[symbol]);
                foreach (uint padding in (uint[])[0, uint.MaxValue])
                {
                    uint bits = reversed | (padding << lens[symbol]);
                    (int decoded, int length) = Lookup(table, bits, tableBits);
                    Assert.Equal(symbol, decoded);
                    Assert.Equal(lens[symbol], length);
                }
                checkedSymbols++;
            }
            Assert.Equal(lens.Count(static len => len != 0), checkedSymbols);
        }

        [Fact]
        public void ShrinkNarrowsTheTableToTheLongestCodeword()
        {
            uint[] table = new uint[InflateFormat.LitlenTableSize];
            ushort[] sorted = new ushort[8];

            Assert.True(HuffmanTableBuilder.TryBuild(table, [3, 3, 3, 3, 3, 2, 4, 4], _symbolResults, 11, 15, sorted, shrink: true, out int used));

            Assert.Equal(4, used);
        }

        [Theory]
        [InlineData(new byte[] { 1, 1, 1 })]
        [InlineData(new byte[] { 2, 2, 2, 2, 2 })]
        [InlineData(new byte[] { 1, 2, 2, 15 })]
        public void OverSubscribedCodeIsRejected(byte[] lens)
        {
            Assert.False(HuffmanTableBuilder.TryBuild(
                new uint[InflateFormat.LitlenTableSize], lens, _symbolResults, 11, 15, new ushort[lens.Length], shrink: false, out _));
        }

        [Theory]
        [InlineData(new byte[] { 2, 2, 2 })]
        [InlineData(new byte[] { 2 })]
        [InlineData(new byte[] { 1, 0, 2 })]
        [InlineData(new byte[] { 1, 15 })]
        public void IncompleteCodeIsRejected(byte[] lens)
        {
            Assert.False(HuffmanTableBuilder.TryBuild(
                new uint[InflateFormat.LitlenTableSize], lens, _symbolResults, 11, 15, new ushort[lens.Length], shrink: false, out _));
        }

        [Fact]
        public void SingleOneBitCodewordFillsTheTableWithItsSymbol()
        {
            uint[] table = new uint[256];

            Assert.True(HuffmanTableBuilder.TryBuild(table, [0, 0, 1, 0], _symbolResults, 8, 15, new ushort[4], shrink: false, out _));

            Assert.All(table, static entry => Assert.Equal((2U << 16) | (1U << 8) | 1U, entry));
        }

        [Fact]
        public void EmptyCodeFillsTheTableWithSymbolZero()
        {
            uint[] table = new uint[256];
            Array.Fill(table, uint.MaxValue);

            Assert.True(HuffmanTableBuilder.TryBuild(table, [0, 0, 0, 0], _symbolResults, 8, 15, new ushort[4], shrink: false, out _));

            Assert.All(table, static entry => Assert.Equal((1U << 8) | 1U, entry));
        }

        private static (int Symbol, int Length) Lookup(uint[] table, uint bits, int tableBits)
        {
            uint entry = table[bits & ((1U << tableBits) - 1)];
            if ((entry & InflateFormat.SubtablePointer) == 0)
            {
                return ((int)(entry >> 16), (int)(entry & 0xFF));
            }
            int subtableBits = (int)((entry >> 8) & 0x3F);
            uint subEntry = table[(entry >> 16) + ((bits >> tableBits) & ((1U << subtableBits) - 1))];
            return ((int)(subEntry >> 16), tableBits + (int)(subEntry & 0xFF));
        }

        // RFC 1951 section 3.2.2.
        private static uint[] CanonicalCodewords(byte[] lens)
        {
            int[] counts = new int[16];
            foreach (byte len in lens)
            {
                counts[len]++;
            }
            counts[0] = 0;
            uint[] next = new uint[16];
            uint code = 0;
            for (int bits = 1; bits <= 15; bits++)
            {
                code = (code + (uint)counts[bits - 1]) << 1;
                next[bits] = code;
            }
            uint[] codewords = new uint[lens.Length];
            for (int symbol = 0; symbol < lens.Length; symbol++)
            {
                if (lens[symbol] != 0)
                {
                    codewords[symbol] = next[lens[symbol]]++;
                }
            }
            return codewords;
        }

        private static uint Reverse(uint codeword, int length)
        {
            uint reversed = 0;
            for (int i = 0; i < length; i++)
            {
                reversed = (reversed << 1) | ((codeword >> i) & 1);
            }
            return reversed;
        }
    }
}
