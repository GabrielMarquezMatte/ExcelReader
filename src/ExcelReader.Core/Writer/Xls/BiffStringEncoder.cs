using System.Buffers;
using System.Text;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Xls
{
    internal static class BiffStringEncoder
    {
        internal static void WriteShort(BiffBuffer buffer, ReadOnlySpan<char> value)
        {
            bool compressed = CanCompress(value);
            buffer.WriteByte((byte)value.Length);
            WriteFlagsAndChars(buffer, value, compressed);
        }

        internal static void WriteLong(BiffBuffer buffer, ReadOnlySpan<char> value)
        {
            bool compressed = CanCompress(value);
            buffer.WriteU16(value.Length);
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
            Span<byte> dest = buffer.GetSpan(value.Length);
            Encoding.Latin1.GetBytes(value, dest);
            buffer.Advance(value.Length);
        }

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
