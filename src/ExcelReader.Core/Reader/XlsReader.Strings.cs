using System.Buffers;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelReader.Core.Reader
{
    public sealed partial class XlsReader
    {
        private const string _cp1252 = "\u20AC\u0081\u201A\u0192\u201E\u2026\u2020\u2021" +
                                        "\u02C6\u2030\u0160\u2039\u0152\u008D\u017D\u008F" +
                                        "\u0090\u2018\u2019\u201C\u201D\u2022\u2013\u2014" +
                                        "\u02DC\u2122\u0161\u203A\u0153\u009D\u017E\u0178";

        private static string DecodeCompressedString(ReadOnlySpan<byte> bytes, int charCount)
        {
            ReadOnlySpan<byte> compressed = bytes[..charCount];
            // CP1252 and ASCII agree byte-for-byte below 0x80 (only 0x80-0x9F get remapped by
            // DecodeCp1252); almost all real-world text is ASCII, so this Ascii.IsValid guard (a
            // vectorized scan) skips the byte->char widening loop and rented buffer entirely for the
            // common case. Any byte >= 0x80 fails IsValid and falls through to the general path below.
            if (Ascii.IsValid(compressed))
            {
                return Encoding.ASCII.GetString(compressed);
            }
            char[] rented = ArrayPool<char>.Shared.Rent(charCount);
            try
            {
                for (int i = 0; i < charCount; i++)
                {
                    rented[i] = DecodeCp1252(compressed[i]);
                }
                return new string(rented, 0, charCount);
            }
            finally
            {
                ArrayPool<char>.Shared.Return(rented);
            }
        }

        // BIFF8 string: bit 0 of `flags` selects compressed (1 byte/char, CP1252) vs wide (UTF-16LE).
        private static int DecodeStringToUtf8(ReadOnlySpan<byte> src, int charCount, byte flags, Span<byte> dest)
        {
            if ((flags & 1) != 0)
            {
                // UTF-16LE code units read directly as chars; Encoding.UTF8.GetBytes combines surrogate
                // pairs into 4-byte sequences and replaces any lone surrogate with U+FFFD — matters for
                // astral chars (e.g. emoji) split across a CONTINUE boundary.
                return Encoding.UTF8.GetBytes(MemoryMarshal.Cast<byte, char>(src[..(charCount * 2)]), dest);
            }
            // Pure ASCII (ASCII and CP1252 agree byte-for-byte below 0x80): the source bytes already
            // are the UTF-8 encoding, so copy directly instead of widening to char and transcoding.
            ReadOnlySpan<byte> compressed = src[..charCount];
            if (Ascii.IsValid(compressed))
            {
                compressed.CopyTo(dest);
                return charCount;
            }
            // Widen CP1252 bytes to chars once, then one bulk UTF-8 transcode — same shape as
            // DecodeSharedStrings, instead of a Rune-per-byte encode.
            char[] rented = ArrayPool<char>.Shared.Rent(charCount);
            try
            {
                for (int i = 0; i < charCount; i++)
                {
                    rented[i] = DecodeCp1252(compressed[i]);
                }
                return Encoding.UTF8.GetBytes(rented.AsSpan(0, charCount), dest);
            }
            finally
            {
                ArrayPool<char>.Shared.Return(rented);
            }
        }

        private static char DecodeCp1252(byte value)
        {
            return value is >= 0x80 and <= 0x9F ? _cp1252[value - 0x80] : (char)value;
        }
    }
}
