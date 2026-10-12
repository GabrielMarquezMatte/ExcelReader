namespace ExcelReader.Core.Reader.Zip.Inflate
{
    internal static class InflateFormat
    {
        internal const int LitlenTableBits = 11;
        internal const int OffsetTableBits = 8;
        internal const int PrecodeTableBits = 7;

        // Worst-case table sizes for these table widths, from zlib's `enough` tool as used by libdeflate.
        internal const int LitlenTableSize = 2342;
        internal const int OffsetTableSize = 402;
        internal const int PrecodeTableSize = 128;

        internal const int MaxLitlenSymbols = 288;
        internal const int MaxOffsetSymbols = 32;
        internal const int PrecodeSymbols = 19;
        internal const int MaxCodewordLength = 15;
        internal const int MaxPrecodeLength = 7;
        internal const int MaxRepeat = 138;

        internal const uint Literal = 0x8000_0000;
        internal const uint Exceptional = 0x8000;
        internal const uint SubtablePointer = 0x4000;
        internal const uint EndOfBlock = 0x2000;

        internal static readonly uint[] LitlenResults = BuildLitlenResults();
        internal static readonly uint[] OffsetResults = BuildOffsetResults();
        internal static readonly uint[] PrecodeResults = BuildPrecodeResults();

        internal static ReadOnlySpan<byte> PrecodeOrder => [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

        private static uint[] BuildLitlenResults()
        {
            ReadOnlySpan<ushort> bases = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258, 258, 258];
            ReadOnlySpan<byte> extraBits = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0, 0, 0];
            uint[] results = new uint[MaxLitlenSymbols];
            for (uint literal = 0; literal < 256; literal++)
            {
                results[literal] = Literal | (literal << 16);
            }
            results[256] = Exceptional | EndOfBlock;
            for (int i = 0; i < bases.Length; i++)
            {
                results[257 + i] = ((uint)bases[i] << 16) | extraBits[i];
            }
            return results;
        }

        private static uint[] BuildOffsetResults()
        {
            ReadOnlySpan<ushort> bases = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577, 24577, 24577];
            ReadOnlySpan<byte> extraBits = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13, 13, 13];
            uint[] results = new uint[MaxOffsetSymbols];
            for (int i = 0; i < results.Length; i++)
            {
                results[i] = ((uint)bases[i] << 16) | extraBits[i];
            }
            return results;
        }

        private static uint[] BuildPrecodeResults()
        {
            uint[] results = new uint[PrecodeSymbols];
            for (uint i = 0; i < results.Length; i++)
            {
                results[i] = i << 16;
            }
            return results;
        }
    }
}
