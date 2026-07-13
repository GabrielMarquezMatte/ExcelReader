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
        internal const int FmlaString = 8; // formula cell, cached string result
        internal const int FmlaNum = 9;    // formula cell, cached numeric result
        internal const int FmlaBool = 10;  // formula cell, cached bool result
        internal const int FmlaError = 11; // formula cell, cached error result
        internal const int CellRString = 62; // inline rich string
        internal const int SSTItem = 19;
        internal const int Fmt = 44;
        internal const int Xf = 47;
        internal const int BeginSheet = 129;
        internal const int EndSheet = 130;
        internal const int BeginBook = 131;
        internal const int EndBook = 132;
        internal const int BeginWsViews = 133;
        internal const int EndWsViews = 134;
        internal const int BeginWsView = 137;
        internal const int EndWsView = 138;
        internal const int BeginSheetData = 145;
        internal const int EndSheetData = 146;
        internal const int Pane = 151;
        internal const int WbProp = 153;
        internal const int BundleSh = 156;
        internal const int BeginSst = 159;
        internal const int EndSst = 160;
        internal const int BeginCellMetadata = 161;
        internal const int EndCellMetadata = 162;
        internal const int BeginStyleSheet = 278;
        internal const int EndStyleSheet = 279;
        internal const int BeginColInfos = 390;
        internal const int EndColInfos = 391;
        internal const int BeginFills = 603;
        internal const int EndFills = 604;
        internal const int BeginFonts = 611;
        internal const int EndFonts = 612;
        internal const int BeginBorders = 613;
        internal const int EndBorders = 614;
        internal const int BeginFmts = 615;
        internal const int EndFmts = 616;
        internal const int BeginCellXFs = 617;
        internal const int EndCellXFs = 618;
        internal const int BeginCellStyleXFs = 626;
        internal const int EndCellStyleXFs = 627;
        internal const int BeginTableStyles = 648;
        internal const int TableStyleClient = 649;
        internal const int EndTableStyles = 650;
        internal const int LegacyEndSheetData = 92;
        internal const int BeginBundleShs = 143;
        internal const int EndBundleShs = 144;
    }
}
