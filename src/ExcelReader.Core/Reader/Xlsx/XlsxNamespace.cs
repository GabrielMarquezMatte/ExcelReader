namespace ExcelReader.Core.Reader.Xlsx
{
    internal sealed class NsTokens
    {
        internal readonly byte[] RowOpen;
        internal readonly byte[] RowEnd;
        internal readonly byte[] SheetDataEnd;
        internal readonly byte[] WorksheetEnd;
        internal readonly byte[] CellOpen;
        internal readonly byte[] VOpen;
        internal readonly byte[] VClose;
        internal readonly byte[] CClose;
        internal readonly byte[] TOpen;
        internal readonly byte[] TClose;
        internal readonly byte[] RPhOpen;
        internal readonly byte[] RPhClose;
        internal readonly int HeadEnsure;

        internal NsTokens(ReadOnlySpan<byte> prefix)
        {
            RowOpen = XlsxXml.Token("<"u8, prefix, "row"u8);
            RowEnd = XlsxXml.Token("</"u8, prefix, "row"u8);
            SheetDataEnd = XlsxXml.Token("</"u8, prefix, "sheetData"u8);
            WorksheetEnd = XlsxXml.Token("</"u8, prefix, "worksheet"u8);
            CellOpen = XlsxXml.Token("<"u8, prefix, "c"u8);
            VOpen = XlsxXml.Token("<"u8, prefix, "v>"u8);
            VClose = XlsxXml.Token("</"u8, prefix, "v>"u8);
            CClose = XlsxXml.Token("</"u8, prefix, "c>"u8);
            TOpen = XlsxXml.Token("<"u8, prefix, "t"u8);
            TClose = XlsxXml.Token("</"u8, prefix, "t>"u8);
            RPhOpen = XlsxXml.Token("<"u8, prefix, "rPh"u8);
            RPhClose = XlsxXml.Token("</"u8, prefix, "rPh>"u8);
            HeadEnsure = WorksheetEnd.Length + 1;
        }
    }
}
