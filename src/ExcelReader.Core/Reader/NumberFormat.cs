using System.Buffers;

namespace ExcelReader.Core.Reader
{
    // Date/time number-format detection shared by every reader (xls, xlsx, xlsb): a numeric cell
    // renders as a date when its number-format id is a builtin date format, or a custom format
    // code that reads as a date.
    internal static class NumberFormat
    {
        // Builtin SpreadsheetML date/time numFmtIds (ECMA-376 §18.8.30, incl. locale variants).
        internal static bool IsBuiltinDate(int id)
        {
            return id is (>= 14 and <= 22) or (>= 27 and <= 36) or (>= 45 and <= 47) or (>= 50 and <= 58) or (>= 71 and <= 81);
        }

        private static readonly SearchValues<char> _dateLetters = SearchValues.Create("yYmMdDhHsS");

        // True if a format code contains a date/time token (y/m/d/h/s) outside quoted text,
        // [bracketed] sections, and \-escapes. Heuristic, not a full format parser — upgrade if a
        // format with date letters only inside literals is misclassified.
        internal static bool LooksLikeDate(ReadOnlySpan<char> code)
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
    }
}
