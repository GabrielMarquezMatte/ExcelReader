using System.Buffers;

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
            char[] rented = ArrayPool<char>.Shared.Rent(charCount);
            try
            {
                for (int i = 0; i < charCount; i++)
                {
                    rented[i] = DecodeCp1252(bytes[i]);
                }
                return new string(rented, 0, charCount);
            }
            finally
            {
                ArrayPool<char>.Shared.Return(rented);
            }
        }

        private static int DecodeStringToUtf8(ReadOnlySpan<byte> src, int charCount, byte flags, Span<byte> dest)
        {
            int written = 0;
            if ((flags & 1) == 0)
            {
                foreach (ref readonly var ch in src[..charCount])
                {
                    written += WriteUtf8(DecodeCp1252(ch), dest[written..]);
                }
                return written;
            }

            for (int i = 0; i < charCount; i++)
            {
                int cp = src[i * 2] | (src[(i * 2) + 1] << 8);
                // Combine a high+low surrogate pair into one scalar so it emits 4-byte UTF-8 rather than
                // two 3-byte sequences (CESU-8, invalid UTF-8) — matters for astral chars like emoji.
                if (cp is not >= 0xD800 or not <= 0xDBFF || i + 1 >= charCount)
                {
                    written += WriteUtf8(cp, dest[written..]);
                    continue;
                }
                int lo = src[(i + 1) * 2] | (src[((i + 1) * 2) + 1] << 8);
                if (lo is >= 0xDC00 and <= 0xDFFF)
                {
                    cp = 0x10000 + ((cp - 0xD800) << 10) + (lo - 0xDC00);
                    i++;
                }
                written += WriteUtf8(cp, dest[written..]);
            }
            return written;
        }

        private static char DecodeCp1252(byte value)
        {
            return value is >= 0x80 and <= 0x9F ? _cp1252[value - 0x80] : (char)value;
        }

        // Encodes one Unicode scalar as UTF-8 (1–4 bytes). A lone surrogate becomes U+FFFD so the
        // output is always valid UTF-8. CP1252 decoding only ever yields BMP scalars.
        private static int WriteUtf8(int cp, Span<byte> dest)
        {
            if (cp <= 0x7F)
            {
                dest[0] = (byte)cp;
                return 1;
            }
            if (cp <= 0x7FF)
            {
                dest[0] = (byte)(0xC0 | (cp >> 6));
                dest[1] = (byte)(0x80 | (cp & 0x3F));
                return 2;
            }
            if (cp <= 0xFFFF)
            {
                if (cp is >= 0xD800 and <= 0xDFFF)
                {
                    cp = 0xFFFD;
                }
                dest[0] = (byte)(0xE0 | (cp >> 12));
                dest[1] = (byte)(0x80 | ((cp >> 6) & 0x3F));
                dest[2] = (byte)(0x80 | (cp & 0x3F));
                return 3;
            }
            dest[0] = (byte)(0xF0 | (cp >> 18));
            dest[1] = (byte)(0x80 | ((cp >> 12) & 0x3F));
            dest[2] = (byte)(0x80 | ((cp >> 6) & 0x3F));
            dest[3] = (byte)(0x80 | (cp & 0x3F));
            return 4;
        }
    }
}
