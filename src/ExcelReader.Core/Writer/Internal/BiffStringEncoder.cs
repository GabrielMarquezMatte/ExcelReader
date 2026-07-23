using System.Buffers;
using System.Text;

namespace ExcelReader.Core.Writer.Internal
{
    // Encodes .NET strings as BIFF8 unicode strings. Bit 0 of the flags byte selects compressed
    // (1 byte/char) vs UTF-16. Compressed is used only when every char survives the reader's
    // round trip: chars in [0x00..0x7F] ∪ [0xA0..0xFF] map byte==char. 0x80..0x9F decode as CP1252
    // specials on read, so those (and anything > 0xFF) force UTF-16 to stay lossless.
    internal static class BiffStringEncoder
    {
        // u8 char count + flags + chars. Used by BoundSheet sheet names.
        internal static void WriteShort(BiffBuffer buffer, ReadOnlySpan<char> value)
        {
            bool compressed = CanCompress(value);
            buffer.WriteByte((byte)value.Length);
            WriteFlagsAndChars(buffer, value, compressed);
        }

        private static void WriteFlagsAndChars(BiffBuffer buffer, ReadOnlySpan<char> value, bool compressed)
        {
            buffer.WriteByte((byte)(compressed ? 0 : 1));
            WriteChars(buffer, value, compressed);
        }

        internal static void WriteChars(BiffBuffer buffer, ReadOnlySpan<char> value, bool compressed)
        {
            if (!compressed)
            {
                buffer.WriteUtf16(value);
                return;
            }
            // CanCompress already guarantees every char is <= 0xFF and outside 0x80-0x9F, so a
            // narrowing cast per char is exactly what Encoding.Latin1 does — one bulk pass instead
            // of a per-char WriteByte.
            Span<byte> dest = buffer.GetSpan(value.Length);
            Encoding.Latin1.GetBytes(value, dest);
            buffer.Advance(value.Length);
        }

        // [0x00..0x7F] ∪ [0xA0..0xFF] — the chars a compressed (1 byte/char) BIFF8 string can hold
        // without loss (see the class comment for why 0x80-0x9F is excluded).
        private static readonly SearchValues<char> CompressibleChars = BuildCompressibleChars();

        private static SearchValues<char> BuildCompressibleChars()
        {
            Span<char> chars = stackalloc char[0x7F - 0x00 + 1 + 0xFF - 0xA0 + 1];
            int i = 0;
            for (int c = 0x00; c <= 0x7F; c++)
            {
                chars[i++] = (char)c;
            }
            for (int c = 0xA0; c <= 0xFF; c++)
            {
                chars[i++] = (char)c;
            }
            return SearchValues.Create(chars);
        }

        internal static bool CanCompress(ReadOnlySpan<char> value)
        {
            return value.IndexOfAnyExcept(CompressibleChars) < 0;
        }
    }
}
