namespace ExcelReader.Core.Writer.Internal
{
    // Emits the workbook globals substream: BOF, CODEPAGE, optional DATEMODE, the two XFs
    // (general + builtin date format 14), one BoundSheet per sheet, EOF. BoundSheet offsets are
    // written as placeholders; the returned positions let the caller back-patch them once each
    // sheet's stream offset is known.
    internal static class XlsGlobals
    {
        internal const int GeneralXf = 16;
        internal const int DateXf = 17;
        private const int BuiltinDateFormat = 14;

        // A custom style's native XF index: the two builtin XFs (general, date) occupy 16 and 17,
        // so StyleTable's abstract index 2 (its first custom entry) lands at 18, 3 at 19, and so on.
        internal static int CustomXf(int abstractStyleIndex)
        {
            return GeneralXf + abstractStyleIndex;
        }

        // Returns the buffer position of each BoundSheet's lbPlyPos field, in sheet order.
        internal static int[] Write(BiffBuffer buffer, ReadOnlySpan<string> sheetNames, bool date1904, StyleTable styles)
        {
            BiffRecordWriter.WriteBof(buffer, BiffRecord.SubstreamGlobals);
            BiffRecordWriter.WriteInterfaceHdr(buffer, 1200);
            BiffRecordWriter.WriteMms(buffer);
            BiffRecordWriter.WriteInterfaceEnd(buffer);
            BiffRecordWriter.WriteWriteAccess(buffer);
            BiffRecordWriter.WriteCodePage(buffer, 1200); // UTF-16
            BiffRecordWriter.WriteDsf(buffer);
            BiffRecordWriter.WriteTabId(buffer, sheetNames.Length);
            BiffRecordWriter.WriteFnGroupCount(buffer);
            BiffRecordWriter.WriteWindow1(buffer);
            BiffRecordWriter.WriteBackup(buffer);
            BiffRecordWriter.WriteHideObj(buffer);
            if (date1904)
            {
                BiffRecordWriter.WriteDate1904(buffer, date1904: true);
            }
            BiffRecordWriter.WritePrecision(buffer);
            BiffRecordWriter.WriteRefreshAll(buffer);
            BiffRecordWriter.WriteBookBool(buffer);
            BiffRecordWriter.WriteFont(buffer);
            Dictionary<string, int> numFmtIds = styles.AssignCustomNumberFormatIds();
            foreach (var (format, id) in numFmtIds)
            {
                BiffRecordWriter.WriteFormat(buffer, id, format);
            }
            for (int i = 0; i < GeneralXf; i++)
            {
                BiffRecordWriter.WriteStyleXf(buffer, formatIndex: 0);
            }
            BiffRecordWriter.WriteCellXf(buffer, formatIndex: 0);
            BiffRecordWriter.WriteCellXf(buffer, formatIndex: BuiltinDateFormat);
            // Bold/Italic are not represented here (see XlsbWorkbookWriter.WriteStylesAsync for why
            // binary formats keep every custom style on the default font); only NumberFormat varies.
            IReadOnlyList<CellStyle> styleList = styles.Styles;
            for (int i = 2; i < styleList.Count; i++)
            {
                int formatIndex = styleList[i].NumberFormat is string format ? numFmtIds[format] : 0;
                BiffRecordWriter.WriteCellXf(buffer, formatIndex);
            }
            BiffRecordWriter.WriteStyle(buffer);

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
