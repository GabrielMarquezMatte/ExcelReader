using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Core.Reader.Xlsx
{
    public sealed partial class XlsxReader
    {
        private static bool[] ParseStyleDateFlags(ReadOnlySpan<byte> src)
        {
            if (src.IsEmpty)
            {
                return [];
            }

            ReadOnlySpan<byte> prefix = XlsxXml.DetectElementPrefix(src);
            ReadOnlySpan<byte> numFmtTag = "<numFmt "u8, cellXfsTag = "<cellXfs"u8;
            ReadOnlySpan<byte> cellXfsClose = "</cellXfs>"u8, xfTag = "<xf "u8;
            if (!prefix.IsEmpty)
            {
                numFmtTag = XlsxXml.Token("<"u8, prefix, "numFmt "u8);
                cellXfsTag = XlsxXml.Token("<"u8, prefix, "cellXfs"u8);
                cellXfsClose = XlsxXml.Token("</"u8, prefix, "cellXfs>"u8);
                xfTag = XlsxXml.Token("<"u8, prefix, "xf "u8);
            }

            Dictionary<int, bool> custom = new(capacity: 16);
            foreach (var tag in Tags(src, numFmtTag))
            {
                int id = XlsxXml.ParseIntOr(XlsxXml.Attr(tag, " numFmtId="u8), -1);
                if (id >= 0)
                {
                    custom[id] = NumberFormat.LooksLikeDate(XlsxXml.DecodeToString(XlsxXml.Attr(tag, " formatCode="u8)));
                }
            }

            int region = IdxOf(src, 0, cellXfsTag);
            if (region < 0)
            {
                return [];
            }
            int open = IdxOf(src, region, (byte)'>');
            if (open < 0)
            {
                return [];
            }
            int end = IdxOf(src, open, cellXfsClose);
            if (end < 0)
            {
                return [];
            }
            List<bool> flags = new(capacity: 16);
            foreach (var xf in Tags(src.Slice(open + 1, end - open - 1), xfTag))
            {
                int numFmtId = XlsxXml.ParseIntOr(XlsxXml.Attr(xf, " numFmtId="u8), 0);
                flags.Add(WorkbookLookups.ResolveDateFlag(custom, numFmtId));
            }
            return [.. flags];
        }

    }
}
