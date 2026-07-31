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
        // Set only for a shared-string cell whose reader was constructed with a dedup cache (currently
        // XLSX/XLSB/XLS): _sharedIndex is the string's index into that reader's shared-string table
        // (see CellDesc.ToCell), and _sharedCache is the reader-owned index -> materialized-string array,
        // sized to the table's string count. -1 for non-shared cells and for an out-of-range/corrupt
        // index (WorkbookLookups.SharedAt), which the bounds check in GetString() below excludes safely.
        private readonly int _sharedIndex;
        private readonly string?[]? _sharedCache;

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
            : this(type, value, number, hasNumber, styleIndex, sharedIndex: -1, sharedCache: null)
        {
        }

        internal Cell(CellType type, ReadOnlySpan<byte> value, double number, bool hasNumber, int styleIndex,
            int sharedIndex, string?[]? sharedCache)
        {
            Type = type;
            Value = value;
            _number = number;
            _hasNumber = hasNumber;
            StyleIndex = styleIndex;
            _sharedIndex = sharedIndex;
            _sharedCache = sharedCache;
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
            // COR-2: NumberStyles.Float (no AllowThousands) — the default style lets a comma act as a
            // thousands separator, so pt-BR comma-decimal text like "1,5" silently parsed as 15.0
            // instead of failing. This method takes no IFormatProvider, so it can only ever be
            // correct for genuine invariant-formatted text; rejecting ambiguous input is strictly
            // better than silently returning a wrong number 10x off.
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
            // Fast path for binary doubles: hand back the stored value without round-tripping
            // through text. Guards are JIT constants, so non-matching T compiles them away.
            if (!_hasNumber)
            {
                // Text-backed double (e.g. CSV, or an XLSX cell FastDouble.TryParse declined to parse
                // eagerly): try the same exact-representability fast parse before the general parser.
                // FastDouble always treats '.' as the decimal separator, so it's only valid when the
                // caller's culture agrees — otherwise "1.234" under e.g. pt-BR (comma decimal) would
                // silently parse as 1.234 instead of the correct 1234.
                if (typeof(T) == typeof(double) && UsesDotDecimalSeparator(provider) && FastDouble.TryParse(Value, out double fast))
                {
                    result = Unsafe.As<double, T>(ref fast);
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
            // Other numeric targets (decimal, ...), plus out-of-range/non-integral cases above:
            // format once and parse, which exactly matches "parse the formatted text" —
            // e.g. int.TryParse fails on "12.5".
            Span<byte> buffer = stackalloc byte[32];
            return Utf8Formatter.TryFormat(_number, buffer, out int written)
                ? T.TryParse(buffer[..written], provider, out result)
                : T.TryParse(Value, provider, out result);
        }

        // Integral targets: cast directly when the stored double is a whole number that fits the
        // target's range — skips the format+parse round trip that the general path below needs.
        // Non-integral values (e.g. 12.5) and out-of-range values return false so the caller can
        // preserve the general parser's exact semantics.
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
        /// Interprets the cell's numeric value as an Excel serial date under the 1900 date system.
        /// Works on any cell whose value parses as a number, not only cells whose <see cref="Type"/> is
        /// <see cref="CellType.Date"/>. Use the <see cref="TryGetDateTime(bool, out DateTime)"/> overload
        /// for workbooks using the 1904 date system.
        /// </summary>
        public bool TryGetDateTime(out DateTime result)
        {
            return TryGetDateTime(isDate1904: false, out result);
        }

        /// <summary>Interprets the cell's numeric value as an Excel serial date.</summary>
        /// <param name="isDate1904">
        /// Pass true for workbooks using the 1904 date system (e.g. when the reader's IsDate1904 is true)
        /// so the epoch offset is applied correctly.
        /// </param>
        /// <param name="result">The parsed date, when this method returns true.</param>
        public bool TryGetDateTime(bool isDate1904, out DateTime result)
        {
            if (!TryGetDouble(out double serial))
            {
                result = default;
                return false;
            }
            double oadate = ExcelEpoch.SerialToOADate(serial, isDate1904);
            // FromOADate throws outside this range; guard first.
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
            return Encoding.UTF8.GetString(Value);
        }


        /// <summary>Returns the cell's value as a string. Equivalent to <see cref="GetString"/>.</summary>
        public override string ToString()
        {
            return GetString();
        }

    }
}
