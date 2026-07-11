using System.Buffers.Binary;

namespace ExcelReader.Core.Writer.Internal
{
    internal static class Biff12RecordWriter
    {
        internal static void WriteRecord(BiffBuffer dest, int id, ReadOnlySpan<byte> payload = default)
        {
            WriteId(dest, id);
            WriteVarint(dest, payload.Length);
            dest.Write(payload);
        }

        // For records whose payload length the caller already knows (fixed-layout cells, row headers):
        // writes just the id + length header, so the caller can then write the payload's fields straight
        // into `dest` (e.g. via WriteCellHeader/WriteU32/WriteDouble) instead of building it in a temp
        // buffer and copying it in — one fewer memcpy per cell.
        internal static void WriteRecordHeader(BiffBuffer dest, int id, int length)
        {
            WriteId(dest, id);
            WriteVarint(dest, length);
        }

        // One reservation (one Ensure/bounds check) for the whole record — id + length header + payload
        // — instead of a separate Ensure per field (WriteRecordHeader's two WriteByte calls, then one
        // per WriteU32/WriteDouble the caller would otherwise make). The returned span aliases dest's
        // buffer at the payload's offset; the caller fills it directly (e.g. via BinaryPrimitives) and
        // must fill every byte, since dest.Length has already been advanced past it.
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

        internal static void WriteCellHeader(BiffBuffer dest, int column, int style)
        {
            if (column < 0)
            {
                throw new InvalidOperationException("Column index cannot be negative.");
            }
            dest.WriteU32((uint)column);
            dest.WriteU32((uint)style);
        }

        // Span counterpart for callers already holding a WriteFixedRecord payload span.
        internal static void WriteCellHeader(Span<byte> dest, int column, int style)
        {
            if (column < 0)
            {
                throw new InvalidOperationException("Column index cannot be negative.");
            }
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
