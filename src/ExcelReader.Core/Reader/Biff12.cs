using System.Buffers.Binary;
using System.Runtime.InteropServices;

namespace ExcelReader.Core.Reader
{
    // Little-endian field decoders for BIFF12 (.xlsb) record payloads. The record framing lives in
    // Biff12RecordReader; these read the fields inside a payload span.
    internal static class Biff12
    {
        internal static ushort ReadU16(ReadOnlySpan<byte> src, int offset)
        {
            return BinaryPrimitives.ReadUInt16LittleEndian(src.Slice(offset, 2));
        }

        internal static uint ReadU32(ReadOnlySpan<byte> src, int offset)
        {
            return BinaryPrimitives.ReadUInt32LittleEndian(src.Slice(offset, 4));
        }

        internal static double ReadF64(ReadOnlySpan<byte> src, int offset)
        {
            return BinaryPrimitives.ReadDoubleLittleEndian(src.Slice(offset, 8));
        }

        // RkNumber (4 bytes), identical to BIFF8: bit0 = ÷100, bit1 = the upper 30 bits are a signed
        // integer; otherwise they are the high 30 bits of an IEEE-754 double (low 34 bits zero).
        internal static double Rk(uint rk)
        {
            double value;
            if ((rk & 0x02) != 0)
            {
                value = (int)rk >> 2; // arithmetic shift keeps the sign
            }
            else
            {
                ulong bits = (ulong)(rk & 0xFFFFFFFC) << 32;
                value = BitConverter.Int64BitsToDouble((long)bits);
            }
            if ((rk & 0x01) != 0)
            {
                value /= 100.0;
            }
            return value;
        }

        // XLWideString / XLNullableWideString at `offset`: cch (u32) + UTF-16LE chars. The chars span
        // aliases `src` (no copy). cch == 0xFFFFFFFF marks a null nullable string (empty, 4 bytes).
        // Returns false if the declared length runs past `src`.
        internal static bool TryReadWideString(ReadOnlySpan<byte> src, int offset, out ReadOnlySpan<char> chars, out int bytesConsumed)
        {
            chars = default;
            bytesConsumed = 0;
            if (offset + 4 > src.Length)
            {
                return false;
            }
            uint cch = ReadU32(src, offset);
            if (cch == 0xFFFFFFFF)
            {
                bytesConsumed = 4;
                return true;
            }
            long byteLength = (long)cch * 2;
            if (offset + 4 + byteLength > src.Length)
            {
                return false;
            }
            chars = MemoryMarshal.Cast<byte, char>(src.Slice(offset + 4, (int)byteLength));
            bytesConsumed = 4 + (int)byteLength;
            return true;
        }
    }
}
