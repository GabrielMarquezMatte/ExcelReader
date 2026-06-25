using System.Buffers;

namespace ExcelReader.Core.Reader
{
    public sealed partial class XlsReader
    {
        private static bool IsBuiltinDateFormat(int id)
        {
            return id is (>= 14 and <= 22) or (>= 27 and <= 36) or (>= 45 and <= 47) or (>= 50 and <= 58) or (>= 71 and <= 81);
        }

        private static readonly SearchValues<char> _dateLetters = SearchValues.Create("yYmMdDhHsS");

        private static bool LooksLikeDateFormat(ReadOnlySpan<char> code)
        {
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
                        i += 2;
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
