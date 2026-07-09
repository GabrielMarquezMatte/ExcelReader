using System.Buffers.Text;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using ExcelReader.Core.Enums;

namespace ExcelReader.Core.ValueObjects
{
    public readonly ref struct Cell
    {
        private readonly double _number;
        private readonly bool _hasNumber;

        public CellType Type { get; }
        // UTF-8 text bytes: shared-string text, or the raw <v> for bools/errors/XLSX numbers.
        // EMPTY for binary-numeric cells (XLS Number/RK/Date/Formula), which carry the raw double
        // instead — read those via TryGetDouble/TryParse, or TryFormat/GetString for text.
        public ReadOnlySpan<byte> Value { get; }
        // The `s` style index from the worksheet (0 when absent). Escape hatch for date detection,
        // which is deferred: dates arrive as Number, and the caller maps StyleIndex -> format itself.
        public int StyleIndex { get; }

        public Cell(CellType type, ReadOnlySpan<byte> value, int styleIndex = 0)
            : this(type, value, 0, hasNumber: false, styleIndex)
        {
        }

        internal Cell(CellType type, ReadOnlySpan<byte> value, double number, bool hasNumber, int styleIndex)
        {
            Type = type;
            Value = value;
            _number = number;
            _hasNumber = hasNumber;
            StyleIndex = styleIndex;
        }

        // Numeric cells from binary formats (XLS) carry the raw double, so this avoids the
        // format-then-parse round trip. Text-backed cells (XLSX, strings) parse Value as a fallback.
        public bool TryGetDouble(out double value)
        {
            if (_hasNumber)
            {
                value = _number;
                return true;
            }
            return double.TryParse(Value, CultureInfo.InvariantCulture, out value);
        }

        public bool TryParse<T>(IFormatProvider? provider, [MaybeNullWhen(false)] out T result) where T : IUtf8SpanParsable<T>
        {
            // Fast path for binary doubles: hand back the stored value without round-tripping
            // through text. Guards are JIT constants, so non-matching T compiles them away.
            if (!_hasNumber)
            {
                return T.TryParse(Value, provider, out result);
            }
            if (typeof(T) == typeof(double))
            {
                double d = _number;
                result = Unsafe.As<double, T>(ref d);
                return true;
            }
            if (typeof(T) == typeof(float))
            {
                float f = (float)_number;
                result = Unsafe.As<float, T>(ref f);
                return true;
            }
            if (typeof(T) == typeof(decimal))
            {
                decimal m = (decimal)_number;
                result = Unsafe.As<decimal, T>(ref m);
                return true;
            }
            // Integral targets: cast directly when the stored double is a whole number that fits the
            // target's range — skips the format+parse round trip that the general path below needs.
            // Non-integral values (e.g. 12.5) and out-of-range values fall through, matching
            // int.TryParse("12.5") semantics.
            bool isIntegral = _number == Math.Truncate(_number);
            if (typeof(T) == typeof(int))
            {
                if (isIntegral && _number is >= int.MinValue and <= int.MaxValue)
                {
                    int v = (int)_number;
                    result = Unsafe.As<int, T>(ref v);
                    return true;
                }
            }
            else if (typeof(T) == typeof(long))
            {
                if (isIntegral && _number is >= -9223372036854775808.0 and < 9223372036854775808.0)
                {
                    long v = (long)_number;
                    result = Unsafe.As<long, T>(ref v);
                    return true;
                }
            }
            else if (typeof(T) == typeof(short))
            {
                if (isIntegral && _number is >= short.MinValue and <= short.MaxValue)
                {
                    short v = (short)_number;
                    result = Unsafe.As<short, T>(ref v);
                    return true;
                }
            }
            else if (typeof(T) == typeof(sbyte))
            {
                if (isIntegral && _number is >= sbyte.MinValue and <= sbyte.MaxValue)
                {
                    sbyte v = (sbyte)_number;
                    result = Unsafe.As<sbyte, T>(ref v);
                    return true;
                }
            }
            else if (typeof(T) == typeof(uint))
            {
                if (isIntegral && _number is >= uint.MinValue and <= uint.MaxValue)
                {
                    uint v = (uint)_number;
                    result = Unsafe.As<uint, T>(ref v);
                    return true;
                }
            }
            else if (typeof(T) == typeof(ulong))
            {
                if (isIntegral && _number is >= 0.0 and < 18446744073709551616.0)
                {
                    ulong v = (ulong)_number;
                    result = Unsafe.As<ulong, T>(ref v);
                    return true;
                }
            }
            else if (typeof(T) == typeof(ushort))
            {
                if (isIntegral && _number is >= ushort.MinValue and <= ushort.MaxValue)
                {
                    ushort v = (ushort)_number;
                    result = Unsafe.As<ushort, T>(ref v);
                    return true;
                }
            }
            else if (typeof(T) == typeof(byte) && isIntegral && _number is >= byte.MinValue and <= byte.MaxValue)
            {
                byte v = (byte)_number;
                result = Unsafe.As<byte, T>(ref v);
                return true;
            }
            // Other numeric targets (decimal, ...), plus out-of-range/non-integral cases above:
            // format once and parse, which exactly matches "parse the formatted text" —
            // e.g. int.TryParse fails on "12.5".
            Span<byte> buffer = stackalloc byte[32];
            return Utf8Formatter.TryFormat(_number, buffer, out int written)
                ? T.TryParse(buffer[..written], provider, out result)
                : T.TryParse(Value, provider, out result);
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
            if (!TryGetDouble(out double serial))
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

        // Writes the cell's text into destination as UTF-8; false if it doesn't fit.
        // Zero-allocation way to get the text of a binary-numeric cell.
        public bool TryFormat(Span<byte> destination, out int bytesWritten)
        {
            if (_hasNumber)
            {
                return Utf8Formatter.TryFormat(_number, destination, out bytesWritten);
            }
            if (Value.TryCopyTo(destination))
            {
                bytesWritten = Value.Length;
                return true;
            }
            bytesWritten = 0;
            return false;
        }

        // Allocates — only call when you actually need a string.
        public string GetString()
        {
            if (_hasNumber)
            {
                Span<byte> buffer = stackalloc byte[32];
                return Utf8Formatter.TryFormat(_number, buffer, out int written)
                    ? Encoding.UTF8.GetString(buffer[..written])
                    : string.Empty;
            }
            return Encoding.UTF8.GetString(Value);
        }


        public override string ToString()
        {
            return GetString();
        }

    }
}
