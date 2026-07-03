using System.Buffers.Text;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using ExcelReader.Core.Enums;
using ExcelReader.Core.ValueObjects;
using FastEnumUtility;

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
                .MakeGenericMethod(typeof(T), propType)
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
                if (cell.Type == CellType.Empty)
                {
                    return true;
                }
                setter(ref model, cell.GetString());
                return true;
            };
        }

        private static ColumnParser<T> BuildBoolParser<T>(PropertyInfo prop)
        {
            RefAction<T, bool> setter = CompileSetter<T, bool>(prop);
            return (ref model, in cell, _, _) =>
            {
                if (cell.Type == CellType.Empty)
                {
                    return true;
                }
                setter(ref model, IsTruthy(in cell));
                return true;
            };
        }

        private static ColumnParser<T> BuildDateTimeParser<T>(PropertyInfo prop)
        {
            RefAction<T, DateTime> setter = CompileSetter<T, DateTime>(prop);
            return (ref model, in cell, isDate1904, _) =>
            {
                if (cell.Type == CellType.Empty)
                {
                    return true;
                }
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
                if (cell.Type == CellType.Empty)
                {
                    return true;
                }
                if (!cell.TryGetDateTime(isDate1904, out DateTime dt))
                {
                    return false;
                }
                setter(ref model, DateOnly.FromDateTime(dt));
                return true;
            };
        }

        // CSV text-date parsers: the cell holds a date string (e.g. "2026-07-02" or ISO "O" form).
        // DateTime/DateOnly implement ISpanParsable (char) but not IUtf8SpanParsable, so decode the
        // short field to a stack char buffer and parse culture-aware — no heap allocation.
        private static ColumnParser<T> BuildTextDateTimeParser<T>(PropertyInfo prop)
        {
            RefAction<T, DateTime> setter = CompileSetter<T, DateTime>(prop);
            return (ref model, in cell, _, provider) =>
            {
                if (cell.Type == CellType.Empty)
                {
                    return true;
                }
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
                if (cell.Type == CellType.Empty)
                {
                    return true;
                }
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
                if (cell.Type == CellType.Empty)
                {
                    return true;
                }
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
                if (cell.Type == CellType.Empty)
                {
                    return true;
                }
                if (!TryParseDateOnlyText(in cell, provider, out DateOnly d))
                {
                    return false;
                }
                setter(ref model, d);
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

        private static ColumnParser<T> BuildParsableCore<T, TProp>(PropertyInfo prop)
            where TProp : IUtf8SpanParsable<TProp>
        {
            RefAction<T, TProp> setter = CompileSetter<T, TProp>(prop);
            return (ref model, in cell, _, provider) =>
            {
                if (cell.Type == CellType.Empty)
                {
                    return true;
                }
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
                if (cell.Type == CellType.Empty)
                {
                    return true;
                }
                setter(ref model, IsTruthy(in cell));
                return true;
            };
        }

        private static ColumnParser<T> BuildNullableDateTimeParser<T>(PropertyInfo prop)
        {
            RefAction<T, DateTime?> setter = CompileSetter<T, DateTime?>(prop);
            return (ref model, in cell, isDate1904, _) =>
            {
                if (cell.Type == CellType.Empty)
                {
                    return true;
                }
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
                if (cell.Type == CellType.Empty)
                {
                    return true;
                }
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
                if (cell.Type == CellType.Empty)
                {
                    return true;
                }
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
        // Guid does not implement IUtf8SpanParsable<Guid> on all targets, so parse from the string
        // form rather than the UTF-8 generic dispatch. Culture is irrelevant for Guid.
        private static ColumnParser<T> BuildGuidParser<T>(PropertyInfo prop)
        {
            RefAction<T, Guid> setter = CompileSetter<T, Guid>(prop);
            return (ref model, in cell, _, _) =>
            {
                if (cell.Type == CellType.Empty)
                {
                    return true;
                }
                if (!Guid.TryParse(cell.GetString(), out Guid value))
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
                if (cell.Type == CellType.Empty)
                {
                    return true;
                }
                if (!Guid.TryParse(cell.GetString(), out Guid value))
                {
                    return false;
                }
                setter(ref model, value);
                return true;
            };
        }
#endif

        // Enum.TryParse accepts both member names ("Active") and their numeric form ("2"),
        // so binary-numeric and text cells both resolve. Culture is irrelevant for enums.
        private static ColumnParser<T> BuildEnumCore<T, TEnum>(PropertyInfo prop)
            where TEnum : struct, Enum
        {
            RefAction<T, TEnum> setter = CompileSetter<T, TEnum>(prop);
            return (ref model, in cell, _, _) =>
            {
                if (cell.Type == CellType.Empty)
                {
                    return true;
                }
                var valueString = cell.GetString();
                if (!FastEnum.TryParse(valueString, ignoreCase: true, out TEnum value)
                    && !Enum.TryParse(valueString, ignoreCase: true, out value))
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
                if (cell.Type == CellType.Empty)
                {
                    return true;
                }
                var valueString = cell.GetString();
                if (!FastEnum.TryParse(valueString, ignoreCase: true, out TEnum parsed)
                    && !Enum.TryParse(valueString, ignoreCase: true, out parsed))
                {
                    return false;
                }
                setter(ref model, parsed);
                return true;
            };
        }

        // Empty cells are short-circuited here (keep default), matching every built-in parser, so the
        // converter only ever sees a populated cell.
        private static ColumnParser<T> BuildConverterCore<T, TProp>(PropertyInfo prop, object converter)
        {
            var typed = (IExcelCellConverter<TProp>)converter;
            RefAction<T, TProp> setter = CompileSetter<T, TProp>(prop);
            return (ref model, in cell, isDate1904, provider) =>
            {
                if (cell.Type == CellType.Empty)
                {
                    return true;
                }
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

        private static bool IsTruthy(in Cell cell)
        {
            return cell.Value.SequenceEqual("1"u8)
                || cell.Value.SequenceEqual("TRUE"u8)
                || cell.Value.SequenceEqual("true"u8);
        }
    }
}
