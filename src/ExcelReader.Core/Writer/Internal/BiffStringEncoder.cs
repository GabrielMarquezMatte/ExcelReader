using System.Diagnostics.CodeAnalysis;

namespace ExcelReader.Core.Writer.Internal
{
    // Encodes .NET strings as BIFF8 unicode strings. Bit 0 of the flags byte selects compressed
    // (1 byte/char) vs UTF-16. Compressed is used only when every char survives the reader's
    // round trip: chars in [0x00..0x7F] ∪ [0xA0..0xFF] map byte==char. 0x80..0x9F decode as CP1252
    // specials on read, so those (and anything > 0xFF) force UTF-16 to stay lossless.
    internal static class BiffStringEncoder
    {
        // u16 char count + flags + chars. Used by Label and Format records.
        internal static void WriteLong(BiffBuffer buffer, ReadOnlySpan<char> value)
        {
            bool compressed = CanCompress(value);
            buffer.WriteU16(value.Length);
            WriteFlagsAndChars(buffer, value, compressed);
        }

        // u8 char count + flags + chars. Used by BoundSheet sheet names.
        internal static void WriteShort(BiffBuffer buffer, ReadOnlySpan<char> value)
        {
            bool compressed = CanCompress(value);
            buffer.WriteByte((byte)value.Length);
            WriteFlagsAndChars(buffer, value, compressed);
        }

        [SuppressMessage("Performance", "HLQ004:The enumerator returns a reference to the item",
            Justification = "Iterating char by value; 'ref readonly char' gains nothing for a 2-byte primitive.")]
        private static void WriteFlagsAndChars(BiffBuffer buffer, ReadOnlySpan<char> value, bool compressed)
        {
            buffer.WriteByte((byte)(compressed ? 0 : 1));
            if (compressed)
            {
                foreach (char c in value)
                {
                    buffer.WriteByte((byte)c);
                }
            }
            else
            {
                foreach (char c in value)
                {
                    buffer.WriteU16(c);
                }
            }
        }

        [SuppressMessage("Performance", "HLQ004:The enumerator returns a reference to the item",
            Justification = "Iterating char by value; 'ref readonly char' gains nothing for a 2-byte primitive.")]
        private static bool CanCompress(ReadOnlySpan<char> value)
        {
            foreach (char c in value)
            {
                if (c > 0xFF || c is >= (char)0x80 and <= (char)0x9F)
                {
                    return false;
                }
            }
            return true;
        }
    }
}
