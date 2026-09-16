using System.Buffers.Text;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using ExcelReader.Core.Enums;

namespace ExcelReader.Core.ValueObjects
{
    /// <summary>
    /// A single worksheet cell's value, exposed as a zero-allocation view over the reader's underlying
    /// buffers. Only valid for the lifetime of the row it was read from — do not store it past that point.
    /// </summary>
    public readonly ref struct Cell
    {
        private readonly double _number;
        private readonly bool _hasNumber;
        private readonly int _sharedIndex;
        private readonly string?[]? _sharedCache;
        private readonly Utf8StringCache? _contentCache;

        /// <summary>The kind of value this cell holds.</summary>
        public CellType Type { get; }
        /// <summary>
        /// The cell's raw UTF-8 text bytes. Empty for binary-numeric cells (XLS Number/RK/Date/Formula);
        /// use <see cref="TryGetDouble"/>, <see cref="TryParse{T}"/>, <see cref="TryFormat"/>, or
        /// <see cref="GetString"/> to read those instead.
        /// </summary>
        public ReadOnlySpan<byte> Value { get; }
        /// <summary>
        /// The worksheet's <c>s</c> style index for this cell (0 when absent). Callers can map this to a
        /// number format to detect dates themselves, since date cells arrive with <see cref="Type"/> of
        /// <see cref="CellType.Number"/>.
        /// </summary>
        public int StyleIndex { get; }

        /// <summary>Creates a cell with the given type, text bytes, and optional style index.</summary>
        public Cell(CellType type, ReadOnlySpan<byte> value, int styleIndex = 0)
            : this(type, value, 0, hasNumber: false, styleIndex)
        {
        }

        internal Cell(CellType type, ReadOnlySpan<byte> value, double number, bool hasNumber, int styleIndex)
            : this(type, value, number, hasNumber, styleIndex, sharedIndex: -1, sharedCache: null, contentCache: null)
        {
        }

        internal Cell(CellType type, ReadOnlySpan<byte> value, double number, bool hasNumber, int styleIndex,
            int sharedIndex, string?[]? sharedCache)
            : this(type, value, number, hasNumber, styleIndex, sharedIndex, sharedCache, contentCache: null)
        {
        }

        internal Cell(CellType type, ReadOnlySpan<byte> value, double number, bool hasNumber, int styleIndex,
            int sharedIndex, string?[]? sharedCache, Utf8StringCache? contentCache)
        {
            Type = type;
            Value = value;
            _number = number;
            _hasNumber = hasNumber;
            StyleIndex = styleIndex;
            _sharedIndex = sharedIndex;
            _sharedCache = sharedCache;
            _contentCache = contentCache;
        }

        /// <summary>Reads the cell's value as a <see cref="double"/>; returns false if it isn't numeric.</summary>
        /// <remarks>
        /// Numeric cells from binary formats (XLS) carry the raw double, so this avoids the
        /// format-then-parse round trip. Text-backed cells (XLSX, strings) parse <see cref="Value"/> as a fallback.
        /// </remarks>
        public bool TryGetDouble(out double value)
        {
            if (_hasNumber)
            {
                value = _number;
                return true;
            }
            return FastDouble.TryParse(Value, out value)
                || double.TryParse(Value, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }

        /// <summary>
        /// Parses the cell's value as <typeparamref name="T"/>, using the stored binary double directly
        /// when available instead of round-tripping through text.
        /// </summary>
        /// <param name="provider">
        /// The format provider used for text parsing and for deciding whether '.' is the decimal separator.
        /// </param>
        /// <param name="result">The parsed value, when this method returns true.</param>
        [SkipLocalsInit]
        public bool TryParse<T>(IFormatProvider? provider, [MaybeNullWhen(false)] out T result) where T : IUtf8SpanParsable<T>
        {
            if (!_hasNumber)
            {
                if (typeof(T) == typeof(double) && UsesDotDecimalSeparator(provider) && FastDouble.TryParse(Value, out double fast))
                {
                    result = Unsafe.As<double, T>(ref fast);
                    return true;
                }
                if (TryParseAsciiDigits(Value, provider, out result))
                {
                    return true;
                }
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
                if (double.IsNaN(_number) || double.IsInfinity(_number)
                    || _number < (double)decimal.MinValue || _number > (double)decimal.MaxValue)
                {
                    result = default;
                    return false;
                }
                decimal m = (decimal)_number;
                result = Unsafe.As<decimal, T>(ref m);
                return true;
            }
            if (TryParseIntegral(out result))
            {
                return true;
            }
            Span<byte> buffer = stackalloc byte[32];
            return Utf8Formatter.TryFormat(_number, buffer, out int written)
                ? T.TryParse(buffer[..written], provider, out result)
                : T.TryParse(Value, provider, out result);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static bool TryParseAsciiDigits<T>(ReadOnlySpan<byte> utf8, IFormatProvider? provider, [MaybeNullWhen(false)] out T result)
            where T : IUtf8SpanParsable<T>
        {
            result = default;
            if (typeof(T) != typeof(int) && typeof(T) != typeof(long))
            {
                return false;
            }
            bool negative = !utf8.IsEmpty && utf8[0] == (byte)'-';
            if (negative)
            {
                if (!ReferenceEquals(provider, CultureInfo.InvariantCulture))
                {
                    return false;
                }
                utf8 = utf8[1..];
            }
            int maxDigits = typeof(T) == typeof(int) ? 9 : 18;
            if (utf8.IsEmpty || utf8.Length > maxDigits)
            {
                return false;
            }
            long accumulated = 0;
            foreach (ref readonly byte b in utf8)
            {
                uint digit = (uint)(b - (byte)'0');
                if (digit > 9)
                {
                    return false;
                }
                accumulated = (accumulated * 10) + digit;
            }
            if (negative)
            {
                accumulated = -accumulated;
            }
            if (typeof(T) == typeof(int))
            {
                int i = (int)accumulated;
                result = Unsafe.As<int, T>(ref i);
                return true;
            }
            result = Unsafe.As<long, T>(ref accumulated);
            return true;
        }

        [SkipLocalsInit]
        private bool TryParseIntegral<T>([MaybeNullWhen(false)] out T result) where T : IUtf8SpanParsable<T>
        {
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
            result = default;
            return false;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static bool UsesDotDecimalSeparator(IFormatProvider? provider)
        {
            return provider is null
                || ReferenceEquals(provider, CultureInfo.InvariantCulture)
                || string.Equals(NumberFormatInfo.GetInstance(provider).NumberDecimalSeparator, ".", StringComparison.Ordinal);
        }

        /// <summary>
        /// Interprets the cell's numeric value as an Excel serial date under the 1900 date system, or,
        /// when the value is not numeric, parses it as an ISO-8601 date or date-time. Works on any cell,
        /// not only cells whose <see cref="Type"/> is <see cref="CellType.Date"/> — text formats such as
        /// CSV carry dates as text rather than serials. Use the
        /// <see cref="TryGetDateTime(bool, out DateTime)"/> overload for workbooks using the 1904 date system.
        /// </summary>
        public bool TryGetDateTime(out DateTime result)
        {
            return TryGetDateTime(isDate1904: false, out result);
        }

        /// <summary>
        /// Interprets the cell's numeric value as an Excel serial date, falling back to an ISO-8601
        /// text date (<c>yyyy-MM-dd</c>, optionally followed by <c>T</c> or a space, a time, and up to
        /// seven fractional-second digits) when the value is not numeric. A trailing zone designator is
        /// not accepted.
        /// </summary>
        /// <param name="isDate1904">
        /// Pass true for workbooks using the 1904 date system (e.g. when the reader's IsDate1904 is true)
        /// so the epoch offset is applied correctly. Ignored for a text date, which carries its own calendar.
        /// </param>
        /// <param name="result">The parsed date, when this method returns true.</param>
        public bool TryGetDateTime(bool isDate1904, out DateTime result)
        {
            // Text dates go first: FastDate rejects a non-date on its second byte, while reaching it
            // through TryGetDouble would mean a full failing double parse on every date cell.
            if (!_hasNumber && FastDate.TryParse(Value, out result))
            {
                return true;
            }
            if (!TryGetDouble(out double serial))
            {
                result = default;
                return false;
            }
            double oadate = ExcelEpoch.SerialToOADate(serial, isDate1904);
            if (oadate is > -657435.0 and < 2958466.0)
            {
                result = DateTime.FromOADate(oadate);
                return true;
            }
            result = default;
            return false;
        }

        /// <summary>
        /// Writes the cell's text into <paramref name="destination"/> as UTF-8, without allocating;
        /// returns false if the buffer is too small.
        /// </summary>
        public bool TryFormat(Span<byte> destination, out int bytesWritten)
        {
            if (_hasNumber && Value.IsEmpty)
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

        /// <summary>
        /// Returns the cell's value as a string, allocating a new instance unless it is a repeated
        /// shared string served from the reader's dedup cache. Only call this when a string is required.
        /// </summary>
        /// <remarks>
        /// A repeated value — the common case for categorical columns — returns the same cached instance
        /// instead of decoding UTF-8 and allocating again. See the constructor for how the dedup cache is supplied.
        /// </remarks>
        [SkipLocalsInit]
        public string GetString()
        {
            if (_hasNumber && Value.IsEmpty)
            {
                Span<byte> buffer = stackalloc byte[32];
                return Utf8Formatter.TryFormat(_number, buffer, out int written)
                    ? Encoding.UTF8.GetString(buffer[..written])
                    : string.Empty;
            }
            if (_sharedCache is not null && !Value.IsEmpty && (uint)_sharedIndex < (uint)_sharedCache.Length)
            {
                ref string? cached = ref _sharedCache[_sharedIndex];
                return cached ??= Encoding.UTF8.GetString(Value);
            }
            if (_contentCache is not null && !Value.IsEmpty)
            {
                return _contentCache.GetOrAdd(Value);
            }
            return Encoding.UTF8.GetString(Value);
        }


        /// <summary>Returns the cell's value as a string. Equivalent to <see cref="GetString"/>.</summary>
        public override string ToString()
        {
            return GetString();
        }

    }
}
