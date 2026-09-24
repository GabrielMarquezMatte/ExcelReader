using System.Collections.Frozen;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Core.Parser.Internal
{
    internal static class ColumnParserFactory
    {
        private static readonly MethodInfo _buildParsableMethod =
            typeof(ColumnParserFactory).GetMethod(
                nameof(BuildParsableCore),
                BindingFlags.NonPublic | BindingFlags.Static)!;

        private static readonly MethodInfo _buildNullableParsableMethod =
            typeof(ColumnParserFactory).GetMethod(
                nameof(BuildNullableParsableCore),
                BindingFlags.NonPublic | BindingFlags.Static)!;

        private static readonly MethodInfo _buildSpanParsableMethod =
            typeof(ColumnParserFactory).GetMethod(
                nameof(BuildSpanParsableCore),
                BindingFlags.NonPublic | BindingFlags.Static)!;

        private static readonly MethodInfo _buildNullableSpanParsableMethod =
            typeof(ColumnParserFactory).GetMethod(
                nameof(BuildNullableSpanParsableCore),
                BindingFlags.NonPublic | BindingFlags.Static)!;

        private static readonly MethodInfo _buildEnumMethod =
            typeof(ColumnParserFactory).GetMethod(
                nameof(BuildEnumCore),
                BindingFlags.NonPublic | BindingFlags.Static)!;

        private static readonly MethodInfo _buildNullableEnumMethod =
            typeof(ColumnParserFactory).GetMethod(
                nameof(BuildNullableEnumCore),
                BindingFlags.NonPublic | BindingFlags.Static)!;

        private static readonly MethodInfo _buildConverterMethod =
            typeof(ColumnParserFactory).GetMethod(
                nameof(BuildConverterCore),
                BindingFlags.NonPublic | BindingFlags.Static)!;

        private static readonly FrozenSet<Type> _parsableTypes = FrozenSet.ToFrozenSet(
        [
            typeof(int), typeof(long), typeof(double), typeof(float), typeof(decimal),
            typeof(short), typeof(byte), typeof(uint), typeof(ulong), typeof(ushort),
            typeof(sbyte), typeof(char), typeof(Half), typeof(Int128), typeof(UInt128),
            typeof(Guid),
        ]);

        private static readonly FrozenSet<Type> _spanParsableTypes = FrozenSet.ToFrozenSet(
        [
            typeof(TimeSpan), typeof(DateTimeOffset),
        ]);

        [RequiresUnreferencedCode("Building a column parser reflects over the property's type and setter, which trimming may remove.")]
        [RequiresDynamicCode("Building a column parser dispatches through MakeGenericMethod for the property's concrete type.")]
        internal static ColumnParser<T>? Build<T>(PropertyInfo prop, bool csvTextDates = false)
            where T : allows ref struct
        {
            Type propType = prop.PropertyType;
            Type? innerNullable = Nullable.GetUnderlyingType(propType);
            if (innerNullable is not null)
            {
                return BuildNullableParser<T>(prop, innerNullable, csvTextDates);
            }
            return BuildConcreteParser<T>(prop, propType, csvTextDates);
        }

        [RequiresUnreferencedCode("Building a converter-backed parser instantiates converterType and dispatches through MakeGenericMethod, which trimming may remove.")]
        [RequiresDynamicCode("Building a converter-backed parser calls MakeGenericType/MakeGenericMethod for the converter's concrete type.")]
        internal static ColumnParser<T> BuildConverter<T>(PropertyInfo prop, Type converterType)
            where T : allows ref struct
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
            where T : allows ref struct
        {
            if (propType == typeof(string))
            {
                return BuildStringParser<T>(prop);
            }
            if (propType == typeof(ReadOnlySpan<byte>))
            {
                return BuildSpanParser<T>(prop);
            }
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
                return BuildValue<T, TimeOnly>(prop, ReadTimeOnly);
            }
            if (propType.IsEnum)
            {
                return (ColumnParser<T>?)_buildEnumMethod
                    .MakeGenericMethod(typeof(T), propType)
                    .Invoke(null, [prop]);
            }
            if (_spanParsableTypes.Contains(propType))
            {
                return (ColumnParser<T>?)_buildSpanParsableMethod
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
            where T : allows ref struct
        {
            if (innerType == typeof(bool))
            {
                return BuildNullableValue<T, bool>(prop, ReadBool);
            }
            if (innerType == typeof(DateTime))
            {
                return BuildNullableValue<T, DateTime>(prop, DateTimeReader(textDates));
            }
            if (innerType == typeof(DateOnly))
            {
                return BuildNullableValue<T, DateOnly>(prop, DateOnlyReader(textDates));
            }
            if (innerType == typeof(TimeOnly))
            {
                return BuildNullableValue<T, TimeOnly>(prop, ReadTimeOnly);
            }
            if (innerType.IsEnum)
            {
                return (ColumnParser<T>?)_buildNullableEnumMethod
                    .MakeGenericMethod(typeof(T), innerType)
                    .Invoke(null, [prop]);
            }
            if (_spanParsableTypes.Contains(innerType))
            {
                return (ColumnParser<T>?)_buildNullableSpanParsableMethod
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
            where T : allows ref struct
        {
            RefAction<T, string> setter = CompileSetter<T, string>(prop);
            return (ref model, in cell, _, _) =>
            {
                setter(ref model, cell.GetString());
                return true;
            };
        }

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

        private delegate bool CellReader<V>(in Cell cell, bool isDate1904, IFormatProvider provider, out V value);

        private static ColumnParser<T> BuildValue<T, V>(PropertyInfo prop, CellReader<V> read)
            where T : allows ref struct
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
            where T : allows ref struct
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

#pragma warning disable S1172 
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

        internal static bool ReadTimeOnly(in Cell cell, bool isDate1904, IFormatProvider provider, out TimeOnly value)
        {
            if (!cell.TryGetDouble(out double serial) || !double.IsFinite(serial))
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

        [SkipLocalsInit]
        internal static bool TryParseSpanParsable<TValue>(in Cell cell, IFormatProvider provider, [MaybeNullWhen(false)] out TValue value)
            where TValue : ISpanParsable<TValue>
        {
            Span<char> stack = stackalloc char[Utf8Text.StackChars];
            ReadOnlySpan<char> chars = Utf8Text.Decode(cell.Value, stack, out char[]? rented);
            try
            {
                return TValue.TryParse(chars, provider, out value);
            }
            finally
            {
                Utf8Text.Release(rented);
            }
        }

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

        private static TimeOnly TimeOnlyFromSerial(double serial)
        {
            double fraction = serial - Math.Floor(serial);
            long ticks = (long)Math.Round(fraction * TimeSpan.TicksPerDay, MidpointRounding.AwayFromZero);
            return new TimeOnly(ticks == TimeSpan.TicksPerDay ? 0 : ticks);
        }

        [SkipLocalsInit]
        private static bool TryParseDateTimeText(in Cell cell, IFormatProvider provider, out DateTime value)
        {
            ReadOnlySpan<byte> utf8 = cell.Value;
            if (FastDate.TryParse(utf8, out value))
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
            if (cell.Value.Length == 10 && FastDate.TryParse(cell.Value, out DateTime date))
            {
                value = DateOnly.FromDateTime(date);
                return true;
            }
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
            where T : allows ref struct
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
            where T : allows ref struct
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

        private static ColumnParser<T> BuildSpanParsableCore<T, TProp>(PropertyInfo prop)
            where TProp : ISpanParsable<TProp>
            where T : allows ref struct
        {
            RefAction<T, TProp> setter = CompileSetter<T, TProp>(prop);
            return (ref model, in cell, _, provider) =>
            {
                if (!TryParseSpanParsable<TProp>(in cell, provider, out var value))
                {
                    return false;
                }
                setter(ref model, value);
                return true;
            };
        }

        private static ColumnParser<T> BuildNullableSpanParsableCore<T, TProp>(PropertyInfo prop)
            where TProp : struct, ISpanParsable<TProp>
            where T : allows ref struct
        {
            RefAction<T, TProp?> setter = CompileSetter<T, TProp?>(prop);
            return (ref model, in cell, _, provider) =>
            {
                if (!TryParseSpanParsable(in cell, provider, out TProp parsed))
                {
                    return false;
                }
                setter(ref model, parsed);
                return true;
            };
        }

        private static class EnumCache<TEnum>
            where TEnum : struct, Enum
        {
            private static readonly FrozenDictionary<string, TEnum> _nameMap = BuildNameMap();
            private static readonly FrozenDictionary<string, TEnum>.AlternateLookup<ReadOnlySpan<char>> _alternateLookup = _nameMap.GetAlternateLookup<ReadOnlySpan<char>>();
            private static readonly FrozenDictionary<long, TEnum> _valueMap = BuildValueMap();

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
                return _alternateLookup.TryGetValue(name, out value);
            }
        }

        private static ColumnParser<T> BuildEnumCore<T, TEnum>(PropertyInfo prop)
            where TEnum : struct, Enum
            where T : allows ref struct
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
            where T : allows ref struct
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

        private static ColumnParser<T> BuildConverterCore<T, TProp, TConv>(PropertyInfo prop, object converter)
            where TConv : IExcelCellConverter<TProp>
            where T : allows ref struct
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
            where T : allows ref struct
        {
            MethodInfo setter = prop.GetSetMethod()!;
            if (typeof(T).IsValueType)
            {
                return setter.CreateDelegate<RefAction<T, TProp>>();
            }
            Action<T, TProp> act = setter.CreateDelegate<Action<T, TProp>>();
            return (ref model, value) => act(model, value);
        }

        private static RefAction<T, TProp> CompileRefStructSetter<T, TProp>(PropertyInfo prop)
            where T : allows ref struct
            where TProp : allows ref struct
        {
            return prop.GetSetMethod()!.CreateDelegate<RefAction<T, TProp>>();
        }

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
