namespace ExcelReader.Core.Reader
{
    internal sealed class SharedStringTokens
    {
        internal readonly byte[] SstTag;
        internal readonly byte[] SiTag;
        internal readonly byte[] SiClose;
        internal readonly byte[] TOpen;
        internal readonly byte[] TClose;
        internal readonly byte[] RPhOpen;
        internal readonly byte[] RPhClose;

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
