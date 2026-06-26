using System.Buffers.Text;

namespace ExcelReader.Core.Reader
{
    public sealed partial class XlsxReader
    {
        // Builds the cellXfs-index -> isDate table from xl/styles.xml. A style is a date when its
        // numFmtId is a builtin date/time format or a custom <numFmt> whose code reads as a date.
        private static bool[] ParseStyleDateFlags(ReadOnlySpan<byte> src)
        {
            if (src.IsEmpty)
            {
                return [];
            }

            // Custom formats: numFmtId -> isDate(formatCode). Builtin ids (<164) handled by IsBuiltinDate.
            Dictionary<int, bool> custom = new(capacity: 16);
            foreach (var tag in Tags(src, "<numFmt "u8))
            {
                int id = ParseIntOr(XlsxXml.Attr(tag, " numFmtId=\""u8), -1);
                if (id >= 0)
                {
                    custom[id] = NumberFormat.LooksLikeDate(Decode(XlsxXml.Attr(tag, " formatCode=\""u8)));
                }
            }

            // Only the <xf> entries inside <cellXfs> are cell styles; <cellStyleXfs> is the master table.
            int region = IdxOf(src, 0, "<cellXfs"u8);
            if (region < 0)
            {
                return [];
            }
            int open = IdxOf(src, region, (byte)'>');
            int end = IdxOf(src, open, "</cellXfs>"u8);
            if (open < 0 || end < 0)
            {
                return [];
            }
            List<bool> flags = new(capacity: 16);
            foreach (var xf in Tags(src.Slice(open + 1, end - open - 1), "<xf "u8))
            {
                int numFmtId = ParseIntOr(XlsxXml.Attr(xf, " numFmtId=\""u8), 0);
                flags.Add(custom.TryGetValue(numFmtId, out bool d) ? d : NumberFormat.IsBuiltinDate(numFmtId));
            }
            return [.. flags];
        }

        private static int ParseIntOr(ReadOnlySpan<byte> src, int fallback)
        {
            return Utf8Parser.TryParse(src, out int v, out _) ? v : fallback;
        }
    }
}
