using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Xls
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
            ReadOnlySpan<ExcelSheetVisibility> sheetVisibilities, int activeSheet, bool date1904, StyleTable styles,
            XlsSharedStrings sharedStrings)
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
            BiffRecordWriter.WriteU16Records(buffer,
                [BiffRecord.WinProtect, BiffRecord.Protect, BiffRecord.Password, BiffRecord.Prot4Rev, BiffRecord.Prot4RevPass], 0);
            BiffRecordWriter.WriteWindow1(buffer, activeSheet);
            BiffRecordWriter.WriteBackup(buffer);
            BiffRecordWriter.WriteHideObj(buffer);
            BiffRecordWriter.WriteDate1904(buffer, date1904);
            BiffRecordWriter.WritePrecision(buffer);
            BiffRecordWriter.WriteRefreshAll(buffer);
            BiffRecordWriter.WriteBookBool(buffer);
            StyleComponents components = styles.AssignComponents();
            IReadOnlyList<ExcelColor> palette = styles.PaletteColors;
            // Excel writes at least four FONT records, all the default font here.
            for (int i = 0; i < 4; i++)
            {
                BiffRecordWriter.WriteFont(buffer, default, palette);
            }
            for (int i = 1; i < components.Fonts.Count; i++)
            {
                BiffRecordWriter.WriteFont(buffer, components.Fonts[i], palette);
            }
            Dictionary<string, int> numFmtIds = styles.AssignCustomNumberFormatIds();
            foreach (var (format, id) in numFmtIds)
            {
                BiffRecordWriter.WriteFormat(buffer, id, format);
            }
            for (int i = 0; i < GeneralXf; i++)
            {
                BiffRecordWriter.WriteStyleXf(buffer, formatIndex: 0);
            }
            IReadOnlyList<CellStyle> styleList = styles.Styles;
            for (int i = 0; i < styleList.Count; i++)
            {
                BiffRecordWriter.WriteCellXf(buffer, FontIndex(components.FontId(i)), FormatIndex(i, styleList[i], numFmtIds), styleList[i], palette);
            }
            BiffRecordWriter.WriteStyle(buffer);
            if (palette.Count > 0)
            {
                BiffRecordWriter.WritePalette(buffer, palette);
            }

            int[] offsetPositions = new int[sheetNames.Length];
            for (int i = 0; i < sheetNames.Length; i++)
            {
                offsetPositions[i] = buffer.Length + 4;
                BiffRecordWriter.WriteBoundSheet(buffer, sheetOffset: 0, sheetNames[i], sheetVisibilities[i]);
            }

            sharedStrings.Write(buffer);
            BiffRecordWriter.WriteEof(buffer);
            return offsetPositions;
        }

        // BIFF8 has no FONT index 4, so font k of the component table is index k + 4.
        private static int FontIndex(int fontId)
        {
            if (fontId == 0)
            {
                return 0;
            }
            return fontId + 4;
        }

        private static int FormatIndex(int styleIndex, in CellStyle style, Dictionary<string, int> numFmtIds)
        {
            if (styleIndex == 1)
            {
                return BuiltinDateFormat;
            }
            return style.NumberFormat is string format ? numFmtIds[format] : 0;
        }
    }
}
