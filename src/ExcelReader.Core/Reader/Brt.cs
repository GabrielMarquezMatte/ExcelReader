namespace ExcelReader.Core.Reader
{
    // BIFF12 (.xlsb) record type ids. Verified against [MS-XLSB] and pyxlsb.
    internal static class Brt
    {
        internal const int RowHdr = 0;
        internal const int CellBlank = 1;
        internal const int CellRk = 2;
        internal const int CellError = 3;
        internal const int CellBool = 4;
        internal const int CellReal = 5;
        internal const int CellSt = 6;     // inline string
        internal const int CellIsst = 7;   // shared-string index
        internal const int SSTItem = 19;
        internal const int Fmt = 44;
        internal const int Xf = 47;
        internal const int WbProp = 153;
        internal const int BundleSh = 156;
        internal const int BeginCellXFs = 617;
        internal const int EndCellXFs = 618;
        internal const int EndSheetData = 92;
    }
}
