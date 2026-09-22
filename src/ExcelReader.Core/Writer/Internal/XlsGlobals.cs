using ExcelReader.Core.Enums;

namespace ExcelReader.Core.Writer.Internal
{
    internal static class XlsGlobals
    {
        internal const int GeneralXf = 16;
        internal const int DateXf = 17;
        private const int BuiltinDateFormat = 14;

        internal static int CustomXf(int abstractStyleIndex)
        {
            return GeneralXf + abstractStyleIndex;
        }

        internal static int[] Write(BiffBuffer buffer, ReadOnlySpan<string> sheetNames,
            ReadOnlySpan<ExcelSheetVisibility> sheetVisibilities, bool date1904, StyleTable styles)
        {
            BiffRecordWriter.WriteBof(buffer, BiffRecord.SubstreamGlobals);
            BiffRecordWriter.WriteInterfaceHdr(buffer, 1200);
            BiffRecordWriter.WriteMms(buffer);
            BiffRecordWriter.WriteInterfaceEnd(buffer);
            BiffRecordWriter.WriteWriteAccess(buffer);
            BiffRecordWriter.WriteCodePage(buffer, 1200);
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
                offsetPositions[i] = buffer.Length + 4;
                BiffRecordWriter.WriteBoundSheet(buffer, sheetOffset: 0, sheetNames[i], sheetVisibilities[i]);
            }

            BiffRecordWriter.WriteEof(buffer);
            return offsetPositions;
        }
    }
}
