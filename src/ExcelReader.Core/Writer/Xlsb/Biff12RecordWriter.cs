using System.Buffers.Binary;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Xlsb
{
    internal static class Biff12RecordWriter
    {
        internal static void WriteRecord(BiffBuffer dest, int id, ReadOnlySpan<byte> payload = default)
        {
            WriteId(dest, id);
            WriteVarint(dest, payload.Length);
            dest.Write(payload);
        }

        // Cell records: one-byte id and length, so the header is two plain bytes once inlined.
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static Span<byte> WriteCellRecord(BiffBuffer dest, int id, int payloadLen)
        {
            Debug.Assert((uint)id < 0x80 && (uint)payloadLen < 0x80);
            Span<byte> span = dest.GetSpan(2 + payloadLen);
            span[0] = (byte)id;
            span[1] = (byte)payloadLen;
            dest.Advance(2 + payloadLen);
            return span.Slice(2, payloadLen);
        }

        internal static void WriteFixedRecord(BiffBuffer dest, int id, int payloadLen, out Span<byte> payload)
        {
            int idLen = (uint)id < 0x80 ? 1 : 2;
            int varintLen = VarintLength(payloadLen);
            Span<byte> span = dest.GetSpan(idLen + varintLen + payloadLen);

            int i = 0;
            if (idLen == 1)
            {
                span[i++] = (byte)id;
            }
            else
            {
                span[i++] = (byte)((id & 0x7F) | 0x80);
                span[i++] = (byte)(id >> 7);
            }

            int v = payloadLen;
            do
            {
                byte b = (byte)(v & 0x7F);
                v >>= 7;
                if (v != 0)
                {
                    b |= 0x80;
                }
                span[i++] = b;
            }
            while (v != 0);

            dest.Advance(i + payloadLen);
            payload = span.Slice(i, payloadLen);
        }

        private static int VarintLength(int value)
        {
            int len = 1;
            while (value >= 0x80)
            {
                value >>= 7;
                len++;
            }
            return len;
        }

        internal static void WriteWideString(BiffBuffer dest, ReadOnlySpan<char> value)
        {
            dest.WriteU32((uint)value.Length);
            if (!value.IsEmpty)
            {
                dest.WriteUtf16(value);
            }
        }

        internal static void WriteCellHeader(Span<byte> dest, int column, int style)
        {
            BinaryPrimitives.WriteUInt32LittleEndian(dest, (uint)column);
            BinaryPrimitives.WriteUInt32LittleEndian(dest[4..], (uint)style);
        }

        private static void WriteId(BiffBuffer dest, int id)
        {
            if ((uint)id < 0x80)
            {
                dest.WriteByte((byte)id);
                return;
            }
            dest.WriteByte((byte)((id & 0x7F) | 0x80));
            dest.WriteByte((byte)(id >> 7));
        }

        private static void WriteVarint(BiffBuffer dest, int value)
        {
            do
            {
                byte b = (byte)(value & 0x7F);
                value >>= 7;
                if (value != 0)
                {
                    b |= 0x80;
                }
                dest.WriteByte(b);
            }
            while (value != 0);
        }
    }
}
