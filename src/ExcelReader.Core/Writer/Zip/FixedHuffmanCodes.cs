namespace ExcelReader.Core.Writer.Zip
{
    /// <summary>
    /// The fixed Huffman codes of RFC 1951 section 3.2.6, bit-reversed for LSB-first output. Each entry
    /// packs the code, with any extra bits already appended above it, over a 5-bit total length.
    /// </summary>
    internal static class FixedHuffmanCodes
    {
        internal const int LengthShift = 5;
        internal const uint LengthMask = 31;

        private static readonly int[] LengthBases = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
        private static readonly int[] LengthExtraBits = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
        private static readonly int[] DistanceBases = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
        private static readonly int[] DistanceExtraBits = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];

        /// <summary>Indexed by byte value.</summary>
        internal static readonly uint[] Literals = BuildLiterals();

        /// <summary>Indexed by match length, 3 to 258.</summary>
        internal static readonly uint[] Lengths = BuildLengths();

        /// <summary>Indexed by match distance, 1 to 32768.</summary>
        internal static readonly uint[] Distances = BuildDistances();

        private static uint[] BuildLiterals()
        {
            var table = new uint[256];
            for (int literal = 0; literal < table.Length; literal++)
            {
                table[literal] = Symbol(literal);
            }
            return table;
        }

        private static uint[] BuildLengths()
        {
            var table = new uint[259];
            for (int index = 0; index < LengthBases.Length; index++)
            {
                uint symbol = Symbol(257 + index);
                int count = Math.Min(1 << LengthExtraBits[index], table.Length - LengthBases[index]);
                for (int extra = 0; extra < count; extra++)
                {
                    table[LengthBases[index] + extra] = AppendExtra(symbol, extra, LengthExtraBits[index]);
                }
            }
            return table;
        }

        private static uint[] BuildDistances()
        {
            var table = new uint[32_769];
            for (int index = 0; index < DistanceBases.Length; index++)
            {
                uint code = Pack(Reverse((uint)index, 5), 5);
                for (int extra = 0; extra < 1 << DistanceExtraBits[index]; extra++)
                {
                    table[DistanceBases[index] + extra] = AppendExtra(code, extra, DistanceExtraBits[index]);
                }
            }
            return table;
        }

        private static uint Symbol(int symbol)
        {
            if (symbol < 144)
            {
                return Pack(Reverse((uint)(0x30 + symbol), 8), 8);
            }
            if (symbol < 256)
            {
                return Pack(Reverse((uint)(0x190 + symbol - 144), 9), 9);
            }
            if (symbol < 280)
            {
                return Pack(Reverse((uint)(symbol - 256), 7), 7);
            }
            return Pack(Reverse((uint)(0xC0 + symbol - 280), 8), 8);
        }

        private static uint AppendExtra(uint packed, int extra, int extraBits)
        {
            int length = (int)(packed & LengthMask);
            uint code = (packed >> LengthShift) | ((uint)extra << length);
            return Pack(code, length + extraBits);
        }

        private static uint Pack(uint code, int length)
        {
            return (code << LengthShift) | (uint)length;
        }

        private static uint Reverse(uint code, int length)
        {
            uint reversed = 0;
            for (int bit = 0; bit < length; bit++)
            {
                reversed = (reversed << 1) | ((code >> bit) & 1);
            }
            return reversed;
        }
    }
}
