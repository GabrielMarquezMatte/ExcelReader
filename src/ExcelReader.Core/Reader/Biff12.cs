using System.Buffers.Binary;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ExcelReader.Core.Reader
{
    internal static class Biff12
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static ushort ReadU16(ReadOnlySpan<byte> src, int offset)
        {
            return BinaryPrimitives.ReadUInt16LittleEndian(src.Slice(offset, 2));
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static uint ReadU32(ReadOnlySpan<byte> src, int offset)
        {
            return BinaryPrimitives.ReadUInt32LittleEndian(src.Slice(offset, 4));
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static int ReadI32(ReadOnlySpan<byte> src, int offset)
        {
            return BinaryPrimitives.ReadInt32LittleEndian(src.Slice(offset, 4));
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static double ReadF64(ReadOnlySpan<byte> src, int offset)
        {
            return BinaryPrimitives.ReadDoubleLittleEndian(src.Slice(offset, 8));
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static double Rk(uint rk)
        {
            double value = (rk & 0x02) != 0
                ? (int)rk >> 2
                : BitConverter.Int64BitsToDouble((long)((ulong)(rk & 0xFFFFFFFC) << 32));
            return (rk & 0x01) != 0 ? value / 100.0 : value;
        }

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
