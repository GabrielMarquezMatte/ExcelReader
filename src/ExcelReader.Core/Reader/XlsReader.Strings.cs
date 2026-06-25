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
                char ch = (char)(src[i * 2] | (src[(i * 2) + 1] << 8));
                written += WriteUtf8(ch, dest[written..]);
            }
            return written;
        }

        private static char DecodeCp1252(byte value)
        {
            return value is >= 0x80 and <= 0x9F ? _cp1252[value - 0x80] : (char)value;
        }

        private static int WriteUtf8(char ch, Span<byte> dest)
        {
            if (ch <= 0x7F)
            {
                dest[0] = (byte)ch;
                return 1;
            }
            if (ch <= 0x7FF)
            {
                dest[0] = (byte)(0xC0 | (ch >> 6));
                dest[1] = (byte)(0x80 | (ch & 0x3F));
                return 2;
            }
            dest[0] = (byte)(0xE0 | (ch >> 12));
            dest[1] = (byte)(0x80 | ((ch >> 6) & 0x3F));
            dest[2] = (byte)(0x80 | (ch & 0x3F));
            return 3;
        }
    }
}
