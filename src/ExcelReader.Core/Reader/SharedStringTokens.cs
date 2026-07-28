namespace ExcelReader.Core.Reader
{
    // Streaming twin of NsTokens for xl/sharedStrings.xml: every element token ParseShared matches,
    // built once whether or not the part uses a namespace prefix. Unlike NsTokens (which only exists
    // for the prefixed case, so the unprefixed scanner keeps compile-time literal fast paths), this
    // always materializes byte[] copies — the shared-strings scan runs once per workbook, not once per
    // cell, so the extra array beats holding a ReadOnlySpan<byte> literal across the Fill/FillAsync
    // calls the streaming parser needs (a span cannot survive an await, and the sync growing-search
    // helpers are shared with the async ones to avoid a second copy of that logic).
    internal sealed class SharedStringTokens
    {
        internal readonly byte[] SstTag;
        internal readonly byte[] SiTag;
        internal readonly byte[] SiClose;
        internal readonly byte[] TOpen;
        internal readonly byte[] TClose;
        internal readonly byte[] RPhOpen;
        internal readonly byte[] RPhClose;

        // `prefix` includes the trailing ':' (e.g. "x:"), exactly as DetectElementPrefix returns it
        // empty for the default-namespace case almost every producer emits.
        internal SharedStringTokens(ReadOnlySpan<byte> prefix)
        {
            if (prefix.IsEmpty)
            {
                SstTag = "<sst"u8.ToArray();
                SiTag = "<si"u8.ToArray();
                SiClose = "</si>"u8.ToArray();
                TOpen = "<t"u8.ToArray();
                TClose = "</t>"u8.ToArray();
                RPhOpen = "<rPh"u8.ToArray();
                RPhClose = "</rPh>"u8.ToArray();
                return;
            }
            SstTag = XlsxXml.Token("<"u8, prefix, "sst"u8);
            SiTag = XlsxXml.Token("<"u8, prefix, "si"u8);
            SiClose = XlsxXml.Token("</"u8, prefix, "si>"u8);
            TOpen = XlsxXml.Token("<"u8, prefix, "t"u8);
            TClose = XlsxXml.Token("</"u8, prefix, "t>"u8);
            RPhOpen = XlsxXml.Token("<"u8, prefix, "rPh"u8);
            RPhClose = XlsxXml.Token("</"u8, prefix, "rPh>"u8);
        }
    }
}
