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
