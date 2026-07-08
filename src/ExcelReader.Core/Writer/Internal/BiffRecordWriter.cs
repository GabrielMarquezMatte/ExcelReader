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

        internal static void WriteInterfaceHdr(BiffBuffer buffer, int codePage)
        {
            int len = buffer.BeginRecord(BiffRecord.InterfaceHdr);
            buffer.WriteU16(codePage);
            buffer.EndRecord(len);
        }

        internal static void WriteMms(BiffBuffer buffer)
        {
            int len = buffer.BeginRecord(BiffRecord.Mms);
            buffer.WriteU16(0);
            buffer.EndRecord(len);
        }

        internal static void WriteInterfaceEnd(BiffBuffer buffer)
        {
            int len = buffer.BeginRecord(BiffRecord.InterfaceEnd);
            buffer.EndRecord(len);
        }

        internal static void WriteWriteAccess(BiffBuffer buffer)
        {
            int len = buffer.BeginRecord(BiffRecord.WriteAccess);
            ReadOnlySpan<byte> name = "ExcelReader"u8;
            buffer.Write(name);
            for (int i = name.Length; i < 112; i++)
            {
                buffer.WriteByte((byte)' ');
            }
            buffer.EndRecord(len);
        }

        internal static void WriteDsf(BiffBuffer buffer)
        {
            WriteU16Record(buffer, BiffRecord.Dsf, 0);
        }

        internal static void WriteTabId(BiffBuffer buffer, int sheetCount)
        {
            int len = buffer.BeginRecord(BiffRecord.TabId);
            for (int i = 1; i <= sheetCount; i++)
            {
                buffer.WriteU16(i);
            }
            buffer.EndRecord(len);
        }

        internal static void WriteFnGroupCount(BiffBuffer buffer)
        {
            WriteU16Record(buffer, BiffRecord.FnGroupCount, 17);
        }
        private static ReadOnlySpan<byte> WriteWindow1Payload => [
            0xFF, 0x7F, 0xFF, 0x7F, 0x80, 0x70, 0x30, 0x39,
            0x38, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0xC0, 0x20,
        ];
        internal static void WriteWindow1(BiffBuffer buffer)
        {
            int len = buffer.BeginRecord(BiffRecord.Window1);
            buffer.Write(WriteWindow1Payload);
            buffer.EndRecord(len);
        }

        internal static void WriteBackup(BiffBuffer buffer)
        {
            WriteU16Record(buffer, BiffRecord.Backup, 0);
        }

        internal static void WriteHideObj(BiffBuffer buffer)
        {
            WriteU16Record(buffer, BiffRecord.HideObj, 0);
        }

        internal static void WriteDate1904(BiffBuffer buffer, bool date1904)
        {
            int len = buffer.BeginRecord(BiffRecord.Date1904);
            buffer.WriteU16(date1904 ? 1 : 0);
            buffer.EndRecord(len);
        }

        internal static void WritePrecision(BiffBuffer buffer)
        {
            WriteU16Record(buffer, BiffRecord.Precision, 1);
        }

        internal static void WriteRefreshAll(BiffBuffer buffer)
        {
            WriteU16Record(buffer, BiffRecord.RefreshAll, 0);
        }

        internal static void WriteBookBool(BiffBuffer buffer)
        {
            WriteU16Record(buffer, BiffRecord.BookBool, 0);
        }

        internal static void WriteFont(BiffBuffer buffer)
        {
            int len = buffer.BeginRecord(BiffRecord.Font);
            buffer.WriteU16(220);
            buffer.WriteU16(0);
            buffer.WriteU16(0x7FFF);
            buffer.WriteU16(400);
            buffer.WriteU16(0);
            buffer.WriteByte(0);
            buffer.WriteByte(0);
            buffer.WriteByte(0);
            buffer.WriteByte(0);
            BiffStringEncoder.WriteShort(buffer, "Arial");
            buffer.EndRecord(len);
        }
        private static ReadOnlySpan<byte> FirstPayload => [0xF5, 0xFF, 0x20, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x20];
        private static ReadOnlySpan<byte> SecondPayload => [0x01, 0x00, 0x20, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0xC0, 0x20];
        internal static void WriteStyleXf(BiffBuffer buffer, int formatIndex)
        {
            int len = buffer.BeginRecord(BiffRecord.Xf);
            buffer.WriteU16(0);            // ifnt (font index)
            buffer.WriteU16(formatIndex);  // ifmt
            buffer.Write(FirstPayload);
            buffer.EndRecord(len);
        }

        internal static void WriteCellXf(BiffBuffer buffer, int formatIndex)
        {
            int len = buffer.BeginRecord(BiffRecord.Xf);
            buffer.WriteU16(0);
            buffer.WriteU16(formatIndex);
            buffer.Write(SecondPayload);
            buffer.EndRecord(len);
        }

        internal static void WriteStyle(BiffBuffer buffer)
        {
            int len = buffer.BeginRecord(BiffRecord.Style);
            buffer.WriteU16(0x8000);
            buffer.WriteByte(0);
            buffer.WriteByte(0xFF);
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
        private static ReadOnlySpan<byte> ThirdPayload => [0xB6, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x40, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00];
        internal static void WriteWindow2(BiffBuffer buffer)
        {
            int len = buffer.BeginRecord(BiffRecord.Window2);
            buffer.Write(ThirdPayload);
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
            bool compressed = BiffStringEncoder.CanCompress(value);
            int charBytes = compressed ? value.Length : value.Length * 2;
            if (9 + charBytes <= BiffRecord.MaxPayload)
            {
                int len = buffer.BeginRecord(BiffRecord.Label);
                WriteCellHeader(buffer, row, col, xf);
                buffer.WriteU16(value.Length);
                buffer.WriteByte((byte)(compressed ? 0 : 1));
                BiffStringEncoder.WriteChars(buffer, value, compressed);
                buffer.EndRecord(len);
                return;
            }
            int firstChars = compressed ? 8215 : 4107;
            int lenSplit = buffer.BeginRecord(BiffRecord.Label);
            WriteCellHeader(buffer, row, col, xf);
            buffer.WriteU16(value.Length);
            buffer.WriteByte((byte)(compressed ? 0 : 1));
            BiffStringEncoder.WriteChars(buffer, value[..firstChars], compressed);
            buffer.EndRecord(lenSplit);
            int offset = firstChars;
            int maxContChars = compressed ? 8223 : 4111;
            while (offset < value.Length)
            {
                int take = Math.Min(maxContChars, value.Length - offset);
                int contLen = buffer.BeginRecord(0x003C); // CONTINUE record
                buffer.WriteByte((byte)(compressed ? 0 : 1));
                BiffStringEncoder.WriteChars(buffer, value.Slice(offset, take), compressed);
                buffer.EndRecord(contLen);
                offset += take;
            }
        }

        internal static void WriteBool(BiffBuffer buffer, int row, int col, int xf, bool value)
        {
            int len = buffer.BeginRecord(BiffRecord.BoolErr);
            WriteCellHeader(buffer, row, col, xf);
            buffer.WriteByte((byte)(value ? 1 : 0));
            buffer.WriteByte(0); // fError = 0 -> boolean
            buffer.EndRecord(len);
        }

        private static void WriteCellHeader(BiffBuffer buffer, int row, int col, int xf)
        {
            buffer.WriteU16(row);
            buffer.WriteU16(col);
            buffer.WriteU16(xf);
        }

        private static void WriteU16Record(BiffBuffer buffer, int id, int value)
        {
            int len = buffer.BeginRecord(id);
            buffer.WriteU16(value);
            buffer.EndRecord(len);
        }
    }
}
