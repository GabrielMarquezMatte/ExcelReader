using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Enums;

namespace ExcelReader.Core.ValueObjects
{
    [StructLayout(LayoutKind.Auto)]
    public readonly ref struct Cell
    {
        public CellType Type { get; }
        // UTF-8 bytes of the cell's text: resolved shared-string text, or the raw <v> for numbers/bools/etc.
        public ReadOnlySpan<byte> Value { get; }
        // The `s` style index from the worksheet (0 when absent). Escape hatch for date detection,
        // which is deferred: dates arrive as Number, and the caller maps StyleIndex -> format itself.
        public int StyleIndex { get; }

        public Cell(CellType type, ReadOnlySpan<byte> value, int styleIndex = 0)
        {
            Type = type;
            Value = value;
            StyleIndex = styleIndex;
        }

        public bool TryParse<T>(IFormatProvider? provider, [MaybeNullWhen(false)] out T result) where T : IUtf8SpanParsable<T>
        {
            return T.TryParse(Value, provider, out result);
        }

        // Interprets the cell's numeric value as an Excel serial date (1900 date system).
        // Works on any cell whose value parses as a number — Type == Date signals the source style was a
        // date/time format. For Mac-authored workbooks with XlsxReader.IsDate1904 == true, use the
        // overload that accepts isDate1904 so the 1462-day epoch offset is applied.
        public bool TryGetDateTime(out DateTime result)
        {
            return TryGetDateTime(isDate1904: false, out result);
        }

        // Interprets the cell's numeric value as an Excel serial date.
        // Pass isDate1904: true (from XlsxReader.IsDate1904) to shift the 1904 epoch to
        // DateTime.FromOADate's 1900 epoch (+1462 days: Jan 1 1904 = OADate 1462).
        public bool TryGetDateTime(bool isDate1904, out DateTime result)
        {
            if (!double.TryParse(Value, CultureInfo.InvariantCulture, out double serial))
            {
                result = default;
                return false;
            }
            double oadate = isDate1904 ? serial + 1462.0 : serial;
            // FromOADate throws outside this range; guard first.
            if (oadate is > -657435.0 and < 2958466.0)
            {
                result = DateTime.FromOADate(oadate);
                return true;
            }
            result = default;
            return false;
        }

        // Allocates — only call when you actually need a string.
        public string GetString()
        {
            return Encoding.UTF8.GetString(Value);
        }


        public override string ToString()
        {
            return GetString();
        }

    }
}
