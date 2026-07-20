using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;
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

        [SuppressMessage("Performance", "HLQ004:The enumerator returns a reference to the item",
            Justification = "Iterating char by value; 'ref readonly char' gains nothing for a 2-byte primitive.")]
        internal static bool CanCompress(ReadOnlySpan<char> value)
        {
            ref char ptr = ref MemoryMarshal.GetReference(value);
            int length = value.Length;
            int i = 0;
            if (Vector256.IsHardwareAccelerated && length >= Vector256<ushort>.Count)
            {
                var maxValid = Vector256.Create((ushort)0x00FF);
                var shift = Vector256.Create((ushort)0x0080);
                var rangeLength = Vector256.Create((ushort)0x001F);
                int vectorLoopLimit = length - Vector256<ushort>.Count;
                for (; i <= vectorLoopLimit; i += Vector256<ushort>.Count)
                {
                    ref var offsetPtr = ref Unsafe.Add(ref ptr, i);
                    var v = Vector256.LoadUnsafe(ref Unsafe.As<char, ushort>(ref offsetPtr)).AsUInt16();
                    var overFF = Vector256.GreaterThan(v, maxValid);
                    var shifted = Vector256.Subtract(v, shift);
                    var inRange = Vector256.LessThanOrEqual(shifted, rangeLength);
                    var invalidChars = Vector256.BitwiseOr(overFF, inRange);
                    if (invalidChars != Vector256<ushort>.Zero)
                    {
                        return false;
                    }
                }
            }
            if (Vector128.IsHardwareAccelerated && length >= Vector128<ushort>.Count)
            {
                var maxValid = Vector128.Create((ushort)0x00FF);
                var shift = Vector128.Create((ushort)0x0080);
                var rangeLength = Vector128.Create((ushort)0x001F);
                int vectorLoopLimit = length - Vector128<ushort>.Count;
                for (; i <= vectorLoopLimit; i += Vector128<ushort>.Count)
                {
                    ref var offsetPtr = ref Unsafe.Add(ref ptr, i);
                    var v = Vector128.LoadUnsafe(ref Unsafe.As<char, ushort>(ref offsetPtr)).AsUInt16();
                    var overFF = Vector128.GreaterThan(v, maxValid);
                    var shifted = Vector128.Subtract(v, shift);
                    var inRange = Vector128.LessThanOrEqual(shifted, rangeLength);
                    var invalidChars = Vector128.BitwiseOr(overFF, inRange);
                    if (invalidChars != Vector128<ushort>.Zero)
                    {
                        return false;
                    }
                }
            }
            for (; i < length; i++)
            {
                char c = Unsafe.Add(ref ptr, i);
                if (c is > (char)0xFF or (>= (char)0x80 and <= (char)0x9F))
                {
                    return false;
                }
            }
            return true;
        }
    }
}
