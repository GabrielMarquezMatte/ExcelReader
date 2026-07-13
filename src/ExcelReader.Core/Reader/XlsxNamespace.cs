namespace ExcelReader.Core.Reader
{
    // Prefixed forms of every SpreadsheetML element token the worksheet scanner and inline-string
    // decoder match against, built once when a worksheet's root element carries a namespace prefix
    // (e.g. <x:worksheet>). When the document is unprefixed — the default-namespace case Excel and
    // almost every producer emit — no NsTokens is created and the scanner keeps its compile-time
    // literal byte-match fast paths untouched. See XlsxReader.Enumerator's `_ns` field.
    internal sealed class NsTokens
    {
        internal readonly byte[] RowOpen;      // "<x:row"
        internal readonly byte[] RowEnd;       // "</x:row"
        internal readonly byte[] SheetDataEnd; // "</x:sheetData"
        internal readonly byte[] WorksheetEnd; // "</x:worksheet"
        internal readonly byte[] CellOpen;     // "<x:c"
        internal readonly byte[] VOpen;        // "<x:v>"
        internal readonly byte[] VClose;       // "</x:v>"
        internal readonly byte[] CClose;       // "</x:c>"
        internal readonly byte[] TOpen;        // "<x:t"
        internal readonly byte[] TClose;       // "</x:t>"
        internal readonly byte[] RPhOpen;      // "<x:rPh"
        internal readonly byte[] RPhClose;     // "</x:rPh>"
        // Bytes the top-level scan must buffer before ClassifyHead can match the longest head token
        // ("</x:worksheet") plus one trailing byte for the boundary check.
        internal readonly int HeadEnsure;

        // `prefix` includes the trailing ':' (e.g. "x:"), exactly as DetectElementPrefix returns it.
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
