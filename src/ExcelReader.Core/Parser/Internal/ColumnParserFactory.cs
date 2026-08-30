using System.Buffers.Text;
using System.Collections.Frozen;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Reflection;
using System.Runtime.CompilerServices;
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

        // Excel has no serial-number concept for TimeSpan/DateTimeOffset the way it does for
        // DateTime/DateOnly/TimeOnly, so — unlike those three — text parsing via this generic path is
        // the only sensible interpretation for them.
        private static readonly HashSet<Type> _parsableTypes =
        [
            typeof(int), typeof(long), typeof(double), typeof(float), typeof(decimal),
            typeof(short), typeof(byte), typeof(uint), typeof(ulong), typeof(ushort),
            typeof(sbyte), typeof(char), typeof(Half), typeof(Int128), typeof(UInt128),
            typeof(TimeSpan), typeof(DateTimeOffset),
            // Guid is only reached here on net9+, where it implements IUtf8SpanParsable. On net8 the
            // dedicated Guid build paths (guarded by #if NET8_0 below) intercept it before this set.
            typeof(Guid),
        ];

        // When csvTextDates is true, DateTime and DateOnly parse the cell text (ISO or culture format)
        // rather than an Excel serial number. Only the CSV parser opts in, because CSV has no serial
        // date form. Every other reader leaves csvTextDates false and keeps the serial semantics.
        [RequiresUnreferencedCode("Building a column parser reflects over the property's type and setter, which trimming may remove.")]
        [RequiresDynamicCode("Building a column parser dispatches through MakeGenericMethod for the property's concrete type.")]
        internal static ColumnParser<T>? Build<T>(PropertyInfo prop, bool csvTextDates = false)
#if NET9_0_OR_GREATER
            where T : allows ref struct
#endif
        {
            Type propType = prop.PropertyType;
            Type? innerNullable = Nullable.GetUnderlyingType(propType);
            if (innerNullable is not null)
            {
                return BuildNullableParser<T>(prop, innerNullable, csvTextDates);
            }
            return BuildConcreteParser<T>(prop, propType, csvTextDates);
        }

        // Builds a parser from a user-supplied IExcelCellConverter<TProperty>; a single shared instance
        // is created here and reused for every row.
        [RequiresUnreferencedCode("Building a converter-backed parser instantiates converterType and dispatches through MakeGenericMethod, which trimming may remove.")]
        [RequiresDynamicCode("Building a converter-backed parser calls MakeGenericType/MakeGenericMethod for the converter's concrete type.")]
        internal static ColumnParser<T> BuildConverter<T>(PropertyInfo prop, Type converterType)
#if NET9_0_OR_GREATER
            where T : allows ref struct
#endif
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

        [RequiresUnreferencedCode("Building a column parser reflects over the property's type and setter, which trimming may remove.")]
        [RequiresDynamicCode("Building a column parser dispatches through MakeGenericMethod for the property's concrete type.")]
        private static ColumnParser<T>? BuildConcreteParser<T>(PropertyInfo prop, Type propType, bool textDates)
#if NET9_0_OR_GREATER
            where T : allows ref struct
#endif
        {
            if (propType == typeof(string))
            {
                return BuildStringParser<T>(prop);
            }
#if NET9_0_OR_GREATER
            if (propType == typeof(ReadOnlySpan<byte>))
            {
                return BuildSpanParser<T>(prop);
            }
#endif
            if (propType == typeof(bool))
            {
                return BuildValue<T, bool>(prop, ReadBool);
            }
            if (propType == typeof(DateTime))
            {
                return BuildValue<T, DateTime>(prop, DateTimeReader(textDates));
            }
            if (propType == typeof(DateOnly))
            {
                return BuildValue<T, DateOnly>(prop, DateOnlyReader(textDates));
            }
            if (propType == typeof(TimeOnly))
            {
                // Always serial, regardless of textDates: CSV has no distinct textual form for TimeOnly.
                return BuildValue<T, TimeOnly>(prop, ReadTimeOnly);
            }
#if NET8_0
            if (propType == typeof(Guid))
            {
                return BuildValue<T, Guid>(prop, ReadGuid);
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

        [RequiresUnreferencedCode("Building a column parser reflects over the property's type and setter, which trimming may remove.")]
        [RequiresDynamicCode("Building a column parser dispatches through MakeGenericMethod for the property's concrete type.")]
        private static ColumnParser<T>? BuildNullableParser<T>(PropertyInfo prop, Type innerType, bool textDates)
#if NET9_0_OR_GREATER
            where T : allows ref struct
#endif
        {
            if (innerType == typeof(bool))
            {
                return BuildNullableValue<T, bool>(prop, ReadBool);
            }
            if (innerType == typeof(DateTime))
            {
                return BuildNullableValue<T, DateTime>(prop, DateTimeReader(textDates));
            }
#if NET8_0
            if (innerType == typeof(Guid))
            {
                return BuildNullableValue<T, Guid>(prop, ReadGuid);
            }
#endif
            if (innerType == typeof(DateOnly))
            {
                return BuildNullableValue<T, DateOnly>(prop, DateOnlyReader(textDates));
            }
            if (innerType == typeof(TimeOnly))
            {
                // See BuildConcreteParser's TimeOnly case: always serial.
                return BuildNullableValue<T, TimeOnly>(prop, ReadTimeOnly);
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
#if NET9_0_OR_GREATER
            where T : allows ref struct
#endif
        {
            RefAction<T, string> setter = CompileSetter<T, string>(prop);
            return (ref model, in cell, _, _) =>
            {
                setter(ref model, cell.GetString());
                return true;
            };
        }

#if NET9_0_OR_GREATER
        // Zero-copy text binding: aliases Cell.Value directly instead of allocating via GetString().
        // Valid only until the enumerator's next MoveNext(); a caller needing it past that must copy.
        private static ColumnParser<T> BuildSpanParser<T>(PropertyInfo prop)
            where T : allows ref struct
        {
            RefAction<T, ReadOnlySpan<byte>> setter = CompileRefStructSetter<T, ReadOnlySpan<byte>>(prop);
            return (ref model, in cell, _, _) =>
            {
                setter(ref model, cell.Value);
                return true;
            };
        }
#endif

        // Shared shape behind every value-type column parser: read the cell into a V, then assign
        // through the compiled setter.
        private delegate bool CellReader<V>(in Cell cell, bool isDate1904, IFormatProvider provider, out V value);

        private static ColumnParser<T> BuildValue<T, V>(PropertyInfo prop, CellReader<V> read)
#if NET9_0_OR_GREATER
            where T : allows ref struct
#endif
        {
            RefAction<T, V> setter = CompileSetter<T, V>(prop);
            return (ref model, in cell, isDate1904, provider) =>
            {
                if (!read(in cell, isDate1904, provider, out V value))
                {
                    return false;
                }
                setter(ref model, value);
                return true;
            };
        }

        private static ColumnParser<T> BuildNullableValue<T, V>(PropertyInfo prop, CellReader<V> read)
            where V : struct
#if NET9_0_OR_GREATER
            where T : allows ref struct
#endif
        {
            RefAction<T, V?> setter = CompileSetter<T, V?>(prop);
            return (ref model, in cell, isDate1904, provider) =>
            {
                if (!read(in cell, isDate1904, provider, out V value))
                {
                    return false;
                }
                setter(ref model, value);
                return true;
            };
        }

#pragma warning disable S1172 // CellReader has one fixed signature for all typed cell readers.
        internal static bool ReadBool(in Cell cell, bool isDate1904, IFormatProvider provider, out bool value)
        {
            return TryParseBool(in cell, out value);
        }

        internal static bool ReadDateTime(in Cell cell, bool isDate1904, IFormatProvider _, out DateTime value)
        {
            return cell.TryGetDateTime(isDate1904, out value);
        }

        internal static bool ReadDateOnly(in Cell cell, bool isDate1904, IFormatProvider _, out DateOnly value)
        {
            if (!cell.TryGetDateTime(isDate1904, out DateTime dt))
            {
                value = default;
                return false;
            }
            value = DateOnly.FromDateTime(dt);
            return true;
        }

        // TryGetDouble reads the binary double (XLS/XLSB) or parses the text invariantly (XLSX),
        // matching how the serial is written; a culture-aware parse would misread "0.5" cells.
        internal static bool ReadTimeOnly(in Cell cell, bool isDate1904, IFormatProvider provider, out TimeOnly value)
        {
            if (!cell.TryGetDouble(out double serial))
            {
                value = default;
                return false;
            }
            value = TimeOnlyFromSerial(serial);
            return true;
        }

        internal static bool ReadTextDateTime(in Cell cell, bool _, IFormatProvider provider, out DateTime value)
        {
            return TryParseDateTimeText(in cell, provider, out value);
        }

        internal static bool ReadTextDateOnly(in Cell cell, bool _, IFormatProvider provider, out DateOnly value)
        {
            return TryParseDateOnlyText(in cell, provider, out value);
        }

        internal static bool ReadTextTimeOnly(in Cell cell, bool _, IFormatProvider provider, out TimeOnly value)
        {
            return TryParseTimeOnlyText(in cell, provider, out value);
        }
#pragma warning restore S1172

        internal static bool TryParseEnum<TEnum>(in Cell cell, out TEnum value)
            where TEnum : struct, Enum
        {
            return EnumCache<TEnum>.TryParse(in cell, out value);
        }

        private static CellReader<DateTime> DateTimeReader(bool textDates)
        {
            return textDates ? ReadTextDateTime : ReadDateTime;
        }

        private static CellReader<DateOnly> DateOnlyReader(bool textDates)
        {
            return textDates ? ReadTextDateOnly : ReadDateOnly;
        }

        // Excel time serial -> TimeOnly: the fractional part of the day, rounded to the nearest tick to
        // undo the double round-trip. A value that rounds up to a whole day wraps back to midnight.
        private static TimeOnly TimeOnlyFromSerial(double serial)
        {
            double fraction = serial - Math.Floor(serial);
            long ticks = (long)Math.Round(fraction * TimeSpan.TicksPerDay, MidpointRounding.AwayFromZero);
            return new TimeOnly(ticks == TimeSpan.TicksPerDay ? 0 : ticks);
        }

        // DateTime/DateOnly don't implement IUtf8SpanParsable, so decode to a stack (or pooled) char
        // buffer and parse culture-aware.
        [SkipLocalsInit]
        private static bool TryParseDateTimeText(in Cell cell, IFormatProvider provider, out DateTime value)
        {
            ReadOnlySpan<byte> utf8 = cell.Value;
            // Round-trip ISO 8601 ("O", 27 bytes exactly) parses straight from UTF-8. Offset/Z forms
            // fall through to keep their DateTimeKind semantics identical to the general parser.
            if (utf8.Length == 27 && utf8[10] == (byte)'T'
                && Utf8Parser.TryParse(utf8, out value, out int consumed, 'O') && consumed == 27)
            {
                return true;
            }
            Span<char> stack = stackalloc char[Utf8Text.StackChars];
            ReadOnlySpan<char> chars = Utf8Text.Decode(utf8, stack, out char[]? rented);
            try
            {
                return DateTime.TryParse(chars, provider, DateTimeStyles.None, out value);
            }
            finally
            {
                Utf8Text.Release(rented);
            }
        }

        [SkipLocalsInit]
        private static bool TryParseDateOnlyText(in Cell cell, IFormatProvider provider, out DateOnly value)
        {
            Span<char> stack = stackalloc char[Utf8Text.StackChars];
            ReadOnlySpan<char> chars = Utf8Text.Decode(cell.Value, stack, out char[]? rented);
            try
            {
                return DateOnly.TryParse(chars, provider, DateTimeStyles.None, out value);
            }
            finally
            {
                Utf8Text.Release(rented);
            }
        }

        [SkipLocalsInit]
        private static bool TryParseTimeOnlyText(in Cell cell, IFormatProvider provider, out TimeOnly value)
        {
            Span<char> stack = stackalloc char[Utf8Text.StackChars];
            ReadOnlySpan<char> chars = Utf8Text.Decode(cell.Value, stack, out char[]? rented);
            try
            {
                return TimeOnly.TryParse(chars, provider, DateTimeStyles.None, out value);
            }
            finally
            {
                Utf8Text.Release(rented);
            }
        }

        private static ColumnParser<T> BuildParsableCore<T, TProp>(PropertyInfo prop)
            where TProp : IUtf8SpanParsable<TProp>
#if NET9_0_OR_GREATER
            where T : allows ref struct
#endif
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

        private static ColumnParser<T> BuildNullableParsableCore<T, TProp>(PropertyInfo prop)
            where TProp : struct, IUtf8SpanParsable<TProp>
#if NET9_0_OR_GREATER
            where T : allows ref struct
#endif
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
        [SkipLocalsInit]
        private static bool TryParseGuid(in Cell cell, out Guid value)
        {
            ReadOnlySpan<byte> utf8 = cell.Value;
            if (!utf8.IsEmpty || !cell.TryGetDouble(out double d))
            {
                return TryParseGuidChars(utf8, out value);
            }
            Span<byte> doubleBuf = stackalloc byte[32];
            if (!Utf8Formatter.TryFormat(d, doubleBuf, out int byteWritten))
            {
                value = Guid.Empty;
                return false;
            }
            return TryParseGuidChars(doubleBuf[..byteWritten], out value);
        }

        [SkipLocalsInit]
        private static bool TryParseGuidChars(ReadOnlySpan<byte> utf8, out Guid value)
        {
            Span<char> stack = stackalloc char[Utf8Text.StackChars];
            ReadOnlySpan<char> chars = Utf8Text.Decode(utf8, stack, out char[]? rented);
            try
            {
                return Guid.TryParse(chars, out value);
            }
            finally
            {
                Utf8Text.Release(rented);
            }
        }

        internal static bool ReadGuid(in Cell cell, bool isDate1904, IFormatProvider provider, out Guid value)
        {
            return TryParseGuid(in cell, out value);
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

            [SkipLocalsInit]
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
                Span<char> stack = stackalloc char[Utf8Text.StackChars];
                ReadOnlySpan<char> chars = Utf8Text.Decode(cell.Value, stack, out char[]? rented);
                try
                {
                    return TryLookupName(chars, out value);
                }
                finally
                {
                    Utf8Text.Release(rented);
                }
            }

            private static bool TryLookupName(ReadOnlySpan<char> name, out TEnum value)
            {
#if NET8_0
                return TryLookupSpan(name, out value);
#else
                return _alternateLookup.TryGetValue(name, out value);
#endif
            }
        }

        private static ColumnParser<T> BuildEnumCore<T, TEnum>(PropertyInfo prop)
            where TEnum : struct, Enum
#if NET9_0_OR_GREATER
            where T : allows ref struct
#endif
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
#if NET9_0_OR_GREATER
            where T : allows ref struct
#endif
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

        // TConv (the concrete converter type, not just the interface) only devirtualizes TryConvert for
        // a value-type converter; CoreCLR's canonical generic sharing means a class converter still
        // resolves through the interface at runtime either way.
        private static ColumnParser<T> BuildConverterCore<T, TProp, TConv>(PropertyInfo prop, object converter)
            where TConv : IExcelCellConverter<TProp>
#if NET9_0_OR_GREATER
            where T : allows ref struct
#endif
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

        // Binds the property setter directly via CreateDelegate rather than compiling an Expression
        // tree, avoiding a dynamic-method emission per bound column.
        private static RefAction<T, TProp> CompileSetter<T, TProp>(PropertyInfo prop)
#if NET9_0_OR_GREATER
            where T : allows ref struct
#endif
        {
            MethodInfo setter = prop.GetSetMethod()!;
            if (typeof(T).IsValueType)
            {
                // A struct's implicit `this` is already `ref T` at the CLR level, so an open-instance
                // delegate binds directly.
                return setter.CreateDelegate<RefAction<T, TProp>>();
            }
            // Class model: `this` is a plain reference, so bind to Action<T, TProp> and wrap once.
            Action<T, TProp> act = setter.CreateDelegate<Action<T, TProp>>();
            return (ref model, value) => act(model, value);
        }

#if NET9_0_OR_GREATER
        // Separate from CompileSetter because Action<T,TProp> can't be written in a method generic
        // over a TProp that allows ref struct.
        private static RefAction<T, TProp> CompileRefStructSetter<T, TProp>(PropertyInfo prop)
            where T : allows ref struct
            where TProp : allows ref struct
        {
            return prop.GetSetMethod()!.CreateDelegate<RefAction<T, TProp>>();
        }
#endif

        // Matches "1"/"0" and "true"/"false" case-insensitively; garbage text fails rather than
        // silently becoming false.
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
