using System.Buffers;
using System.Buffers.Text;
using System.Collections.Frozen;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using ExcelReader.Core.Enums;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser.Internal
{
    internal static class ColumnParserFactory
    {
        [SuppressMessage("Blocker Code Smell", "S3011:Reflection should not be used to increase accessibility of classes, methods, or fields",
            Justification = "Private method accessed within same class for generic dispatch; intentional and type-safe.")]
        private static readonly MethodInfo _buildParsableMethod =
            typeof(ColumnParserFactory).GetMethod(
                nameof(BuildParsableCore),
                BindingFlags.NonPublic | BindingFlags.Static)!;

        [SuppressMessage("Blocker Code Smell", "S3011:Reflection should not be used to increase accessibility of classes, methods, or fields",
            Justification = "Private method accessed within same class for generic dispatch; intentional and type-safe.")]
        private static readonly MethodInfo _buildNullableParsableMethod =
            typeof(ColumnParserFactory).GetMethod(
                nameof(BuildNullableParsableCore),
                BindingFlags.NonPublic | BindingFlags.Static)!;

        [SuppressMessage("Blocker Code Smell", "S3011:Reflection should not be used to increase accessibility of classes, methods, or fields",
            Justification = "Private method accessed within same class for generic dispatch; intentional and type-safe.")]
        private static readonly MethodInfo _buildEnumMethod =
            typeof(ColumnParserFactory).GetMethod(
                nameof(BuildEnumCore),
                BindingFlags.NonPublic | BindingFlags.Static)!;

        [SuppressMessage("Blocker Code Smell", "S3011:Reflection should not be used to increase accessibility of classes, methods, or fields",
            Justification = "Private method accessed within same class for generic dispatch; intentional and type-safe.")]
        private static readonly MethodInfo _buildNullableEnumMethod =
            typeof(ColumnParserFactory).GetMethod(
                nameof(BuildNullableEnumCore),
                BindingFlags.NonPublic | BindingFlags.Static)!;

        [SuppressMessage("Blocker Code Smell", "S3011:Reflection should not be used to increase accessibility of classes, methods, or fields",
            Justification = "Private method accessed within same class for generic dispatch; intentional and type-safe.")]
        private static readonly MethodInfo _buildConverterMethod =
            typeof(ColumnParserFactory).GetMethod(
                nameof(BuildConverterCore),
                BindingFlags.NonPublic | BindingFlags.Static)!;

        private static readonly HashSet<Type> _parsableTypes =
        [
            typeof(int), typeof(long), typeof(double), typeof(float), typeof(decimal),
            typeof(short), typeof(byte), typeof(uint), typeof(ulong), typeof(ushort),
            typeof(Guid),
        ];

        // When csvTextDates is true, DateTime and DateOnly parse the cell text (ISO or culture format)
        // rather than an Excel serial number. Only the CSV parser opts in, because CSV has no serial
        // date form. Every other reader leaves csvTextDates false and keeps the serial semantics.
        internal static ColumnParser<T>? Build<T>(PropertyInfo prop, bool csvTextDates = false)
        {
            Type propType = prop.PropertyType;
            Type? innerNullable = Nullable.GetUnderlyingType(propType);
            if (csvTextDates)
            {
                Type effective = innerNullable ?? propType;
                if (effective == typeof(DateTime))
                {
                    return innerNullable is null ? BuildTextDateTimeParser<T>(prop) : BuildTextNullableDateTimeParser<T>(prop);
                }
                if (effective == typeof(DateOnly))
                {
                    return innerNullable is null ? BuildTextDateOnlyParser<T>(prop) : BuildTextNullableDateOnlyParser<T>(prop);
                }
                if (effective == typeof(TimeOnly))
                {
                    return innerNullable is null ? BuildTextTimeOnlyParser<T>(prop) : BuildTextNullableTimeOnlyParser<T>(prop);
                }
            }
            if (innerNullable is not null)
            {
                return BuildNullableParser<T>(prop, innerNullable);
            }
            return BuildConcreteParser<T>(prop, propType);
        }

        // Builds a parser from a user-supplied IExcelCellConverter<TProperty>. The converter type must
        // implement the interface for the property's exact type and have a public parameterless ctor
        // a single shared instance is created here and reused for every row.
        internal static ColumnParser<T> BuildConverter<T>(PropertyInfo prop, Type converterType)
        {
            Type propType = prop.PropertyType;
            Type ifaceType = typeof(IExcelCellConverter<>).MakeGenericType(propType);
            if (!ifaceType.IsAssignableFrom(converterType))
            {
                throw new InvalidOperationException(
                    $"Converter '{converterType}' must implement IExcelCellConverter<{propType}> to convert property '{prop.DeclaringType?.Name}.{prop.Name}'.");
            }
            object converter = Activator.CreateInstance(converterType)
                ?? throw new InvalidOperationException($"Converter '{converterType}' could not be instantiated.");
            return (ColumnParser<T>)_buildConverterMethod
                .MakeGenericMethod(typeof(T), propType, converterType)
                .Invoke(null, [prop, converter])!;
        }

        private static ColumnParser<T>? BuildConcreteParser<T>(PropertyInfo prop, Type propType)
        {
            if (propType == typeof(string))
            {
                return BuildStringParser<T>(prop);
            }
            if (propType == typeof(bool))
            {
                return BuildBoolParser<T>(prop);
            }
            if (propType == typeof(DateTime))
            {
                return BuildDateTimeParser<T>(prop);
            }
            if (propType == typeof(DateOnly))
            {
                return BuildDateOnlyParser<T>(prop);
            }
            if (propType == typeof(TimeOnly))
            {
                return BuildTimeOnlyParser<T>(prop);
            }
#if NET8_0
            if (propType == typeof(Guid))
            {
                return BuildGuidParser<T>(prop);
            }
#endif
            if (propType.IsEnum)
            {
                return (ColumnParser<T>?)_buildEnumMethod
                    .MakeGenericMethod(typeof(T), propType)
                    .Invoke(null, [prop]);
            }
            if (!_parsableTypes.Contains(propType))
            {
                return null;
            }
            return (ColumnParser<T>?)_buildParsableMethod
                .MakeGenericMethod(typeof(T), propType)
                .Invoke(null, [prop]);
        }

        private static ColumnParser<T>? BuildNullableParser<T>(PropertyInfo prop, Type innerType)
        {
            if (innerType == typeof(bool))
            {
                return BuildNullableBoolParser<T>(prop);
            }
            if (innerType == typeof(DateTime))
            {
                return BuildNullableDateTimeParser<T>(prop);
            }
#if NET8_0
            if (innerType == typeof(Guid))
            {
                return BuildNullableGuidParser<T>(prop);
            }
#endif
            if (innerType == typeof(DateOnly))
            {
                return BuildNullableDateOnlyParser<T>(prop);
            }
            if (innerType == typeof(TimeOnly))
            {
                return BuildNullableTimeOnlyParser<T>(prop);
            }
            if (innerType.IsEnum)
            {
                return (ColumnParser<T>?)_buildNullableEnumMethod
                    .MakeGenericMethod(typeof(T), innerType)
                    .Invoke(null, [prop]);
            }
            if (!_parsableTypes.Contains(innerType))
            {
                return null;
            }
            return (ColumnParser<T>?)_buildNullableParsableMethod
                .MakeGenericMethod(typeof(T), innerType)
                .Invoke(null, [prop]);
        }

        private static ColumnParser<T> BuildStringParser<T>(PropertyInfo prop)
        {
            RefAction<T, string> setter = CompileSetter<T, string>(prop);
            return (ref model, in cell, _, _) =>
            {
                setter(ref model, cell.GetString());
                return true;
            };
        }

        private static ColumnParser<T> BuildBoolParser<T>(PropertyInfo prop)
        {
            RefAction<T, bool> setter = CompileSetter<T, bool>(prop);
            return (ref model, in cell, _, _) =>
            {
                if (!TryParseBool(in cell, out bool value))
                {
                    return false;
                }
                setter(ref model, value);
                return true;
            };
        }

        private static ColumnParser<T> BuildDateTimeParser<T>(PropertyInfo prop)
        {
            RefAction<T, DateTime> setter = CompileSetter<T, DateTime>(prop);
            return (ref model, in cell, isDate1904, _) =>
            {
                if (!cell.TryGetDateTime(isDate1904, out DateTime dt))
                {
                    return false;
                }
                setter(ref model, dt);
                return true;
            };
        }

        private static ColumnParser<T> BuildDateOnlyParser<T>(PropertyInfo prop)
        {
            RefAction<T, DateOnly> setter = CompileSetter<T, DateOnly>(prop);
            return (ref model, in cell, isDate1904, _) =>
            {
                if (!cell.TryGetDateTime(isDate1904, out DateTime dt))
                {
                    return false;
                }
                setter(ref model, DateOnly.FromDateTime(dt));
                return true;
            };
        }

        private static ColumnParser<T> BuildTimeOnlyParser<T>(PropertyInfo prop)
        {
            RefAction<T, TimeOnly> setter = CompileSetter<T, TimeOnly>(prop);
            return (ref model, in cell, _, _) =>
            {
                // TryGetDouble reads the binary double (XLS/XLSB) or parses the text invariantly (XLSX),
                // matching how the serial is written; a culture-aware parse would misread "0.5" cells.
                if (!cell.TryGetDouble(out double serial))
                {
                    return false;
                }
                setter(ref model, TimeOnlyFromSerial(serial));
                return true;
            };
        }

        private static ColumnParser<T> BuildNullableTimeOnlyParser<T>(PropertyInfo prop)
        {
            RefAction<T, TimeOnly?> setter = CompileSetter<T, TimeOnly?>(prop);
            return (ref model, in cell, _, _) =>
            {
                // TryGetDouble reads the binary double (XLS/XLSB) or parses the text invariantly (XLSX),
                // matching how the serial is written; a culture-aware parse would misread "0.5" cells.
                if (!cell.TryGetDouble(out double serial))
                {
                    return false;
                }
                setter(ref model, TimeOnlyFromSerial(serial));
                return true;
            };
        }

        // Excel time serial -> TimeOnly: the fractional part of the day, rounded to the nearest tick to
        // undo the double round-trip. A value that rounds up to a whole day wraps back to midnight.
        private static TimeOnly TimeOnlyFromSerial(double serial)
        {
            double fraction = serial - Math.Floor(serial);
            long ticks = (long)Math.Round(fraction * TimeSpan.TicksPerDay, MidpointRounding.AwayFromZero);
            return new TimeOnly(ticks == TimeSpan.TicksPerDay ? 0 : ticks);
        }

        // CSV text-date parsers: the cell holds a date string (e.g. "2026-07-02" or ISO "O" form).
        // DateTime/DateOnly implement ISpanParsable (char) but not IUtf8SpanParsable, so decode the
        // short field to a stack char buffer and parse culture-aware — no heap allocation.
        private static ColumnParser<T> BuildTextDateTimeParser<T>(PropertyInfo prop)
        {
            RefAction<T, DateTime> setter = CompileSetter<T, DateTime>(prop);
            return (ref model, in cell, _, provider) =>
            {
                if (!TryParseDateTimeText(in cell, provider, out DateTime dt))
                {
                    return false;
                }
                setter(ref model, dt);
                return true;
            };
        }

        private static ColumnParser<T> BuildTextNullableDateTimeParser<T>(PropertyInfo prop)
        {
            RefAction<T, DateTime?> setter = CompileSetter<T, DateTime?>(prop);
            return (ref model, in cell, _, provider) =>
            {
                if (!TryParseDateTimeText(in cell, provider, out DateTime dt))
                {
                    return false;
                }
                setter(ref model, dt);
                return true;
            };
        }

        private static ColumnParser<T> BuildTextDateOnlyParser<T>(PropertyInfo prop)
        {
            RefAction<T, DateOnly> setter = CompileSetter<T, DateOnly>(prop);
            return (ref model, in cell, _, provider) =>
            {
                if (!TryParseDateOnlyText(in cell, provider, out DateOnly d))
                {
                    return false;
                }
                setter(ref model, d);
                return true;
            };
        }

        private static ColumnParser<T> BuildTextNullableDateOnlyParser<T>(PropertyInfo prop)
        {
            RefAction<T, DateOnly?> setter = CompileSetter<T, DateOnly?>(prop);
            return (ref model, in cell, _, provider) =>
            {
                if (!TryParseDateOnlyText(in cell, provider, out DateOnly d))
                {
                    return false;
                }
                setter(ref model, d);
                return true;
            };
        }

        private static ColumnParser<T> BuildTextTimeOnlyParser<T>(PropertyInfo prop)
        {
            RefAction<T, TimeOnly> setter = CompileSetter<T, TimeOnly>(prop);
            return (ref model, in cell, _, provider) =>
            {
                if (!TryParseTimeOnlyText(in cell, provider, out TimeOnly t))
                {
                    return false;
                }
                setter(ref model, t);
                return true;
            };
        }

        private static ColumnParser<T> BuildTextNullableTimeOnlyParser<T>(PropertyInfo prop)
        {
            RefAction<T, TimeOnly?> setter = CompileSetter<T, TimeOnly?>(prop);
            return (ref model, in cell, _, provider) =>
            {
                if (!TryParseTimeOnlyText(in cell, provider, out TimeOnly t))
                {
                    return false;
                }
                setter(ref model, t);
                return true;
            };
        }

        // DateTime/DateOnly implement ISpanParsable (char) and IUtf8SpanFormattable, but NOT
        // IUtf8SpanParsable (no parse-from-UTF-8) on either net8 or net10. So decode the short date
        // field to a stack char buffer — allocation-free — and parse culture-aware (honors Culture,
        // e.g. pt-BR "02/07/2026"). Falls back to a string for pathologically long fields.
        private const int MaxStackDateChars = 128;

        private static bool TryParseDateTimeText(in Cell cell, IFormatProvider provider, out DateTime value)
        {
            ReadOnlySpan<byte> utf8 = cell.Value;
            // Round-trip ISO 8601 ("O", no offset — 27 bytes exactly) parses straight from UTF-8,
            // skipping the transcode and the general format-probing parser. Offset/Z forms fall
            // through so their DateTimeKind/local-adjustment semantics stay identical to TryParse.
            if (utf8.Length == 27 && utf8[10] == (byte)'T'
                && Utf8Parser.TryParse(utf8, out value, out int consumed, 'O') && consumed == 27)
            {
                return true;
            }
            if (utf8.Length <= MaxStackDateChars)
            {
                Span<char> chars = stackalloc char[MaxStackDateChars];
                int n = Encoding.UTF8.GetChars(utf8, chars);
                return DateTime.TryParse(chars[..n], provider, DateTimeStyles.None, out value);
            }
            return DateTime.TryParse(cell.GetString(), provider, DateTimeStyles.None, out value);
        }

        private static bool TryParseDateOnlyText(in Cell cell, IFormatProvider provider, out DateOnly value)
        {
            ReadOnlySpan<byte> utf8 = cell.Value;
            if (utf8.Length <= MaxStackDateChars)
            {
                Span<char> chars = stackalloc char[MaxStackDateChars];
                int n = Encoding.UTF8.GetChars(utf8, chars);
                return DateOnly.TryParse(chars[..n], provider, DateTimeStyles.None, out value);
            }
            return DateOnly.TryParse(cell.GetString(), provider, DateTimeStyles.None, out value);
        }

        private static bool TryParseTimeOnlyText(in Cell cell, IFormatProvider provider, out TimeOnly value)
        {
            ReadOnlySpan<byte> utf8 = cell.Value;
            if (utf8.Length <= MaxStackDateChars)
            {
                Span<char> chars = stackalloc char[MaxStackDateChars];
                int n = Encoding.UTF8.GetChars(utf8, chars);
                return TimeOnly.TryParse(chars[..n], provider, DateTimeStyles.None, out value);
            }
            return TimeOnly.TryParse(cell.GetString(), provider, DateTimeStyles.None, out value);
        }

        private static ColumnParser<T> BuildParsableCore<T, TProp>(PropertyInfo prop)
            where TProp : IUtf8SpanParsable<TProp>
        {
            RefAction<T, TProp> setter = CompileSetter<T, TProp>(prop);
            return (ref model, in cell, _, provider) =>
            {
                if (!cell.TryParse<TProp>(provider, out var value))
                {
                    return false;
                }
                setter(ref model, value);
                return true;
            };
        }

        private static ColumnParser<T> BuildNullableBoolParser<T>(PropertyInfo prop)
        {
            RefAction<T, bool?> setter = CompileSetter<T, bool?>(prop);
            return (ref model, in cell, _, _) =>
            {
                if (!TryParseBool(in cell, out bool value))
                {
                    return false;
                }
                setter(ref model, value);
                return true;
            };
        }

        private static ColumnParser<T> BuildNullableDateTimeParser<T>(PropertyInfo prop)
        {
            RefAction<T, DateTime?> setter = CompileSetter<T, DateTime?>(prop);
            return (ref model, in cell, isDate1904, _) =>
            {
                if (!cell.TryGetDateTime(isDate1904, out DateTime dt))
                {
                    return false;
                }
                setter(ref model, dt);
                return true;
            };
        }

        private static ColumnParser<T> BuildNullableDateOnlyParser<T>(PropertyInfo prop)
        {
            RefAction<T, DateOnly?> setter = CompileSetter<T, DateOnly?>(prop);
            return (ref model, in cell, isDate1904, _) =>
            {
                if (!cell.TryGetDateTime(isDate1904, out DateTime dt))
                {
                    return false;
                }
                setter(ref model, DateOnly.FromDateTime(dt));
                return true;
            };
        }

        [SuppressMessage("Blocker Code Smell", "S3011:Reflection should not be used to increase accessibility of classes, methods, or fields",
            Justification = "Called via MakeGenericMethod dispatch; private access is intentional and type-safe.")]
        private static ColumnParser<T> BuildNullableParsableCore<T, TProp>(PropertyInfo prop)
            where TProp : struct, IUtf8SpanParsable<TProp>
        {
            RefAction<T, TProp?> setter = CompileSetter<T, TProp?>(prop);
            return (ref model, in cell, _, provider) =>
            {
                if (!cell.TryParse(provider, out TProp parsed))
                {
                    return false;
                }
                TProp? value = parsed;
                setter(ref model, value);
                return true;
            };
        }

#if NET8_0
        private static bool TryParseGuid(in Cell cell, out Guid value)
        {
            ReadOnlySpan<byte> utf8 = cell.Value;
            if (utf8.Length <= 128)
            {
                Span<char> chars = stackalloc char[128];
                int charCount;
                if (utf8.IsEmpty && cell.TryGetDouble(out double d))
                {
                    Span<byte> byteBuf = stackalloc byte[32];
                    if (Utf8Formatter.TryFormat(d, byteBuf, out int byteWritten))
                    {
                        charCount = Encoding.UTF8.GetChars(byteBuf[..byteWritten], chars);
                    }
                    else
                    {
                        value = Guid.Empty;
                        return false;
                    }
                }
                else
                {
                    charCount = Encoding.UTF8.GetChars(utf8, chars);
                }
                return Guid.TryParse(chars[..charCount], out value);
            }
            else
            {
                char[] chars = ArrayPool<char>.Shared.Rent(utf8.Length);
                try
                {
                    int charCount = Encoding.UTF8.GetChars(utf8, chars);
                    return Guid.TryParse(chars.AsSpan(0, charCount), out value);
                }
                finally
                {
                    ArrayPool<char>.Shared.Return(chars);
                }
            }
        }

        // Guid does not implement IUtf8SpanParsable<Guid> on all targets, so parse from the string
        // form rather than the UTF-8 generic dispatch. Culture is irrelevant for Guid.
        private static ColumnParser<T> BuildGuidParser<T>(PropertyInfo prop)
        {
            RefAction<T, Guid> setter = CompileSetter<T, Guid>(prop);
            return (ref model, in cell, _, _) =>
            {
                if (!TryParseGuid(in cell, out Guid value))
                {
                    return false;
                }
                setter(ref model, value);
                return true;
            };
        }

        private static ColumnParser<T> BuildNullableGuidParser<T>(PropertyInfo prop)
        {
            RefAction<T, Guid?> setter = CompileSetter<T, Guid?>(prop);
            return (ref model, in cell, _, _) =>
            {
                if (!TryParseGuid(in cell, out Guid value))
                {
                    return false;
                }
                setter(ref model, value);
                return true;
            };
        }
#endif

        private static class EnumCache<TEnum>
            where TEnum : struct, Enum
        {
#if !NET8_0
            private static readonly FrozenDictionary<string, TEnum> _nameMap = BuildNameMap();
            private static readonly FrozenDictionary<string, TEnum>.AlternateLookup<ReadOnlySpan<char>> _alternateLookup = _nameMap.GetAlternateLookup<ReadOnlySpan<char>>();
#endif
            private static readonly FrozenDictionary<long, TEnum> _valueMap = BuildValueMap();
#if NET8_0
            private static readonly (string Name, TEnum Value)[] _sortedNames = BuildSortedNames();

            private static (string Name, TEnum Value)[] BuildSortedNames()
            {
                var list = new List<(string Name, TEnum Value)>();
                foreach (TEnum value in Enum.GetValues<TEnum>())
                {
                    string name = value.ToString();
                    list.Add((name, value));
                    long numericValue = Convert.ToInt64(value, CultureInfo.InvariantCulture);
                    string intStr = numericValue.ToString(CultureInfo.InvariantCulture);
                    if (!string.Equals(name, intStr, StringComparison.OrdinalIgnoreCase))
                    {
                        list.Add((intStr, value));
                    }
                }
                var arr = list.ToArray();
                Array.Sort(arr, static (a, b) => string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase));
                return arr;
            }

            private static bool TryLookupSpan(ReadOnlySpan<char> span, out TEnum value)
            {
                int low = 0;
                int high = _sortedNames.Length - 1;
                while (low <= high)
                {
                    int mid = (low + high) >>> 1;
                    var (Name, Value) = _sortedNames[mid];
                    int cmp = span.CompareTo(Name.AsSpan(), StringComparison.OrdinalIgnoreCase);
                    if (cmp == 0)
                    {
                        value = Value;
                        return true;
                    }
                    if (cmp < 0)
                    {
                        high = mid - 1;
                        continue;
                    }
                    low = mid + 1;
                }
                value = default;
                return false;
            }
#endif

#if !NET8_0
            private static FrozenDictionary<string, TEnum> BuildNameMap()
            {
                Dictionary<string, TEnum> map = new(StringComparer.OrdinalIgnoreCase);
                foreach (TEnum value in Enum.GetValues<TEnum>())
                {
                    string name = value.ToString();
                    long numericValue = Convert.ToInt64(value, CultureInfo.InvariantCulture);
                    map[name] = value;
                    map[numericValue.ToString(CultureInfo.InvariantCulture)] = value;
                }
                return map.ToFrozenDictionary(StringComparer.OrdinalIgnoreCase);
            }
#endif
            private static FrozenDictionary<long, TEnum> BuildValueMap()
            {
                Dictionary<long, TEnum> map = [];
                foreach (TEnum value in Enum.GetValues<TEnum>())
                {
                    long numericValue = Convert.ToInt64(value, CultureInfo.InvariantCulture);
                    map[numericValue] = value;
                }
                return map.ToFrozenDictionary();
            }
            public static bool TryParse(in Cell cell, out TEnum value)
            {
                if (cell.Type == CellType.Number && cell.TryGetDouble(out double d))
                {
                    if (d != Math.Truncate(d) || d < long.MinValue || d > long.MaxValue)
                    {
                        value = default;
                        return false;
                    }
                    return _valueMap.TryGetValue((long)d, out value);
                }
#if NET8_0
                ReadOnlySpan<byte> utf8 = cell.Value;
                if (utf8.Length <= 128)
                {
                    Span<char> spanChars = stackalloc char[128];
                    int n = Encoding.UTF8.GetChars(utf8, spanChars);
                    return TryLookupSpan(spanChars[..n], out value);
                }
                char[] chars = ArrayPool<char>.Shared.Rent(utf8.Length);
                try
                {
                    int n = Encoding.UTF8.GetChars(utf8, chars);
                    return TryLookupSpan(chars.AsSpan(0, n), out value);
                }
                finally
                {
                    ArrayPool<char>.Shared.Return(chars);
                }
#else
                ReadOnlySpan<byte> utf8 = cell.Value;
                if (utf8.Length <= 128)
                {
                    Span<char> spanChars = stackalloc char[128];
                    int n = Encoding.UTF8.GetChars(utf8, spanChars);
                    return _alternateLookup.TryGetValue(spanChars[..n], out value);
                }
                char[] chars = ArrayPool<char>.Shared.Rent(utf8.Length);
                try
                {
                    int n = Encoding.UTF8.GetChars(utf8, chars);
                    return _alternateLookup.TryGetValue(chars.AsSpan(0, n), out value);
                }
                finally
                {
                    ArrayPool<char>.Shared.Return(chars);
                }
#endif
            }
        }

        private static ColumnParser<T> BuildEnumCore<T, TEnum>(PropertyInfo prop)
            where TEnum : struct, Enum
        {
            RefAction<T, TEnum> setter = CompileSetter<T, TEnum>(prop);
            return (ref model, in cell, _, _) =>
            {
                if (!EnumCache<TEnum>.TryParse(in cell, out TEnum value))
                {
                    return false;
                }
                setter(ref model, value);
                return true;
            };
        }

        private static ColumnParser<T> BuildNullableEnumCore<T, TEnum>(PropertyInfo prop)
            where TEnum : struct, Enum
        {
            RefAction<T, TEnum?> setter = CompileSetter<T, TEnum?>(prop);
            return (ref model, in cell, _, _) =>
            {
                if (!EnumCache<TEnum>.TryParse(in cell, out TEnum parsed))
                {
                    return false;
                }
                setter(ref model, parsed);
                return true;
            };
        }

        // Callers (RowProjector/CsvRowProjector) skip invoking any ColumnParser for an empty cell, so
        // the converter only ever sees a populated one.
        // TConv is the concrete converter type (not just the interface). This only devirtualizes
        // typed.TryConvert for a value-type converter: CoreCLR shares one compiled body across all
        // reference-type instantiations of a generic method (canonical __Canon sharing), so for a class
        // converter — the common case — the constrained call still resolves through the interface at
        // runtime, same as calling through IExcelCellConverter<TProp> directly.

        private static ColumnParser<T> BuildConverterCore<T, TProp, TConv>(PropertyInfo prop, object converter)
            where TConv : IExcelCellConverter<TProp>
        {
            var typed = (TConv)converter;
            RefAction<T, TProp> setter = CompileSetter<T, TProp>(prop);
            return (ref model, in cell, isDate1904, provider) =>
            {
                if (!typed.TryConvert(in cell, isDate1904, provider, out TProp value))
                {
                    return false;
                }
                setter(ref model, value);
                return true;
            };
        }

        private static RefAction<T, TProp> CompileSetter<T, TProp>(PropertyInfo prop)
        {
            ParameterExpression modelParam = Expression.Parameter(typeof(T).MakeByRefType(), "model");
            ParameterExpression valueParam = Expression.Parameter(typeof(TProp), "value");
            MemberExpression propAccess = Expression.Property(modelParam, prop);
            BinaryExpression assign = Expression.Assign(propAccess, valueParam);
            Type delegateType = typeof(RefAction<,>).MakeGenericType(typeof(T), typeof(TProp));
            LambdaExpression lambda = Expression.Lambda(delegateType, assign, modelParam, valueParam);
            return (RefAction<T, TProp>)lambda.Compile();
        }

        // Matches "1"/"0" and "true"/"false" case-insensitively (so .NET's own bool.ToString() form
        // "True"/"False" round-trips) and reports failure for anything else, so a nullable bool?
        // column with garbage text stays null instead of silently becoming false.
        private static bool TryParseBool(in Cell cell, out bool value)
        {
            ReadOnlySpan<byte> v = cell.Value;
            if (v.Length == 1)
            {
                if (v[0] == (byte)'1') { value = true; return true; }
                if (v[0] == (byte)'0') { value = false; return true; }
            }
            if (Ascii.EqualsIgnoreCase(v, "true"u8)) { value = true; return true; }
            if (Ascii.EqualsIgnoreCase(v, "false"u8)) { value = false; return true; }
            value = false;
            return false;
        }
    }
}
