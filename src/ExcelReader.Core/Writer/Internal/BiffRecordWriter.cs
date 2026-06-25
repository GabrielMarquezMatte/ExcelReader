namespace ExcelReader.Core.Writer.Internal
{
    // Emits BIFF8 records into a BiffBuffer. Layouts follow [MS-XLS]; field offsets match what
    // XlsReader parses (e.g. XF format index at byte 2, Label string at byte 6).
    internal static class BiffRecordWriter
    {
        internal static void WriteBof(BiffBuffer buffer, int substreamType)
        {
            int len = buffer.BeginRecord(BiffRecord.Bof);
            buffer.WriteU16(BiffRecord.Biff8Version);
            buffer.WriteU16(substreamType);
            for (int i = 0; i < 12; i++)
            {
                buffer.WriteByte(0);
            }
            buffer.EndRecord(len);
        }

        internal static void WriteEof(BiffBuffer buffer)
        {
            int len = buffer.BeginRecord(BiffRecord.Eof);
            buffer.EndRecord(len);
        }

        internal static void WriteCodePage(BiffBuffer buffer, int codePage)
        {
            int len = buffer.BeginRecord(BiffRecord.CodePage);
            buffer.WriteU16(codePage);
            buffer.EndRecord(len);
        }

        internal static void WriteDate1904(BiffBuffer buffer, bool date1904)
        {
            int len = buffer.BeginRecord(BiffRecord.Date1904);
            buffer.WriteU16(date1904 ? 1 : 0);
            buffer.EndRecord(len);
        }

        // Minimal 20-byte XF: only the format index (byte 2) matters to the reader's date detection.
        internal static void WriteXf(BiffBuffer buffer, int formatIndex)
        {
            int len = buffer.BeginRecord(BiffRecord.Xf);
            buffer.WriteU16(0);            // ifnt (font index)
            buffer.WriteU16(formatIndex);  // ifmt
            for (int i = 0; i < 16; i++)
            {
                buffer.WriteByte(0);
            }
            buffer.EndRecord(len);
        }

        internal static void WriteBoundSheet(BiffBuffer buffer, int sheetOffset, ReadOnlySpan<char> name)
        {
            int len = buffer.BeginRecord(BiffRecord.BoundSheet);
            buffer.WriteI32(sheetOffset);  // lbPlyPos
            buffer.WriteU16(0);            // grbit (visible worksheet)
            BiffStringEncoder.WriteShort(buffer, name);
            buffer.EndRecord(len);
        }

        internal static void WriteDimension(BiffBuffer buffer, int rowCount, int colCount)
        {
            int len = buffer.BeginRecord(BiffRecord.Dimension);
            buffer.WriteI32(0);            // rwMic
            buffer.WriteI32(rowCount);     // rwMac (last row + 1)
            buffer.WriteU16(0);            // colMic
            buffer.WriteU16(colCount);     // colMac (last col + 1)
            buffer.WriteU16(0);            // reserved
            buffer.EndRecord(len);
        }

        internal static void WriteNumber(BiffBuffer buffer, int row, int col, int xf, double value)
        {
            int len = buffer.BeginRecord(BiffRecord.Number);
            WriteCellHeader(buffer, row, col, xf);
            buffer.WriteDouble(value);
            buffer.EndRecord(len);
        }

        internal static void WriteLabel(BiffBuffer buffer, int row, int col, int xf, ReadOnlySpan<char> value)
        {
            int len = buffer.BeginRecord(BiffRecord.Label);
            WriteCellHeader(buffer, row, col, xf);
            BiffStringEncoder.WriteLong(buffer, value);
            buffer.EndRecord(len);
        }

        internal static void WriteBool(BiffBuffer buffer, int row, int col, int xf, bool value)
        {
            int len = buffer.BeginRecord(BiffRecord.BoolErr);
            WriteCellHeader(buffer, row, col, xf);
            buffer.WriteByte((byte)(value ? 1 : 0));
            buffer.WriteByte(0); // fError = 0 -> boolean
            buffer.EndRecord(len);
        }

        internal static void WriteBlank(BiffBuffer buffer, int row, int col, int xf)
        {
            int len = buffer.BeginRecord(BiffRecord.Blank);
            WriteCellHeader(buffer, row, col, xf);
            buffer.EndRecord(len);
        }

        private static void WriteCellHeader(BiffBuffer buffer, int row, int col, int xf)
        {
            buffer.WriteU16(row);
            buffer.WriteU16(col);
            buffer.WriteU16(xf);
        }
    }
}
