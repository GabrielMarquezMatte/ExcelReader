namespace ExcelReader.Core.Writer.Internal
{
    // Emits the workbook globals substream: BOF, CODEPAGE, optional DATEMODE, the two XFs
    // (general + builtin date format 14), one BoundSheet per sheet, EOF. BoundSheet offsets are
    // written as placeholders; the returned positions let the caller back-patch them once each
    // sheet's stream offset is known.
    internal static class XlsGlobals
    {
        internal const int GeneralXf = 0;
        internal const int DateXf = 1;
        private const int BuiltinDateFormat = 14;

        // Returns the buffer position of each BoundSheet's lbPlyPos field, in sheet order.
        internal static int[] Write(BiffBuffer buffer, ReadOnlySpan<string> sheetNames, bool date1904)
        {
            BiffRecordWriter.WriteBof(buffer, BiffRecord.SubstreamGlobals);
            BiffRecordWriter.WriteCodePage(buffer, 1200); // UTF-16
            if (date1904)
            {
                BiffRecordWriter.WriteDate1904(buffer, date1904: true);
            }
            BiffRecordWriter.WriteXf(buffer, formatIndex: 0);
            BiffRecordWriter.WriteXf(buffer, formatIndex: BuiltinDateFormat);

            int[] offsetPositions = new int[sheetNames.Length];
            for (int i = 0; i < sheetNames.Length; i++)
            {
                // lbPlyPos is the first payload field, i.e. 4 bytes past the record header.
                offsetPositions[i] = buffer.Length + 4;
                BiffRecordWriter.WriteBoundSheet(buffer, sheetOffset: 0, sheetNames[i]);
            }

            BiffRecordWriter.WriteEof(buffer);
            return offsetPositions;
        }
    }
}
