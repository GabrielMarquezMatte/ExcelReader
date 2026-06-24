using System.Buffers;
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
                    custom[id] = LooksLikeDateFormat(Decode(XlsxXml.Attr(tag, " formatCode=\""u8)));
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
                flags.Add(custom.TryGetValue(numFmtId, out bool d) ? d : IsBuiltinDateFormat(numFmtId));
            }
            return [.. flags];
        }

        // Builtin SpreadsheetML date/time numFmtIds (ECMA-376 §18.8.30, incl. locale variants).
        private static bool IsBuiltinDateFormat(int id)
        {
            return id is (>= 14 and <= 22) or (>= 27 and <= 36) or (>= 45 and <= 47) or (>= 50 and <= 58) or (>= 71 and <= 81);
        }

        // True if a format code contains a date/time token (y/m/d/h/s) outside quoted text, [bracketed]
        // sections, and \-escapes. ponytail: heuristic, not a full format parser — upgrade if a format
        // with date letters only inside literals is misclassified.
        private static readonly SearchValues<char> _dateLetters = SearchValues.Create("yYmMdDhHsS");

        private static bool LooksLikeDateFormat(ReadOnlySpan<char> code)
        {
            // Fast SIMD exit: if none of the date letters appear, skip the full parse.
            if (code.IndexOfAny(_dateLetters) < 0)
            {
                return false;
            }
            int i = 0;
            while (i < code.Length)
            {
                switch (code[i])
                {
                    case '"':
                        {
                            int q = code[(i + 1)..].IndexOf('"');
                            if (q < 0)
                            {
                                return false;
                            }
                            i = i + 2 + q;
                            break;
                        }
                    case '[':
                        {
                            int q = code[(i + 1)..].IndexOf(']');
                            if (q < 0)
                            {
                                return false;
                            }
                            i = i + 2 + q;
                            break;
                        }
                    case '\\':
                        i += 2; // skip the escaped char
                        break;
                    case 'y' or 'Y' or 'm' or 'M' or 'd' or 'D' or 'h' or 'H' or 's' or 'S':
                        return true;
                    default:
                        i++;
                        break;
                }
            }
            return false;
        }

        private static int ParseIntOr(ReadOnlySpan<byte> src, int fallback)
        {
            return Utf8Parser.TryParse(src, out int v, out _) ? v : fallback;
        }
    }
}
