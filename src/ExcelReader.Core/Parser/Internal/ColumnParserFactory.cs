using System.Diagnostics.CodeAnalysis;
using System.Linq.Expressions;
using System.Reflection;
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

        private static readonly HashSet<Type> _parsableTypes =
        [
            typeof(int), typeof(long), typeof(double), typeof(float), typeof(decimal),
            typeof(short), typeof(byte), typeof(uint), typeof(ulong), typeof(ushort),
            typeof(Guid),
        ];

        internal static ColumnParser<T>? Build<T>(PropertyInfo prop) where T : new()
        {
            Type propType = prop.PropertyType;
            Type? innerNullable = Nullable.GetUnderlyingType(propType);
            if (innerNullable is not null)
            {
                return BuildNullableParser<T>(prop, innerNullable);
            }
            return BuildConcreteParser<T>(prop, propType);
        }

        private static ColumnParser<T>? BuildConcreteParser<T>(PropertyInfo prop, Type propType) where T : new()
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

        private static ColumnParser<T>? BuildNullableParser<T>(PropertyInfo prop, Type innerType) where T : new()
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

        private static ColumnParser<T> BuildStringParser<T>(PropertyInfo prop) where T : new()
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

        private static ColumnParser<T> BuildBoolParser<T>(PropertyInfo prop) where T : new()
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

        private static ColumnParser<T> BuildDateTimeParser<T>(PropertyInfo prop) where T : new()
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

        private static ColumnParser<T> BuildDateOnlyParser<T>(PropertyInfo prop) where T : new()
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

        private static ColumnParser<T> BuildParsableCore<T, TProp>(PropertyInfo prop)
            where T : new()
            where TProp : IUtf8SpanParsable<TProp>
        {
            RefAction<T, TProp> setter = CompileSetter<T, TProp>(prop);
            return (ref model, in cell, _, provider) =>
            {
                if (cell.Type == CellType.Empty)
                {
                    return true;
                }
                if (!cell.TryParse<TProp>(provider, out TProp? value))
                {
                    return false;
                }
                setter(ref model, value);
                return true;
            };
        }

        private static ColumnParser<T> BuildNullableBoolParser<T>(PropertyInfo prop) where T : new()
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

        private static ColumnParser<T> BuildNullableDateTimeParser<T>(PropertyInfo prop) where T : new()
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

        private static ColumnParser<T> BuildNullableDateOnlyParser<T>(PropertyInfo prop) where T : new()
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
            where T : new()
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
        private static ColumnParser<T> BuildGuidParser<T>(PropertyInfo prop) where T : new()
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

        private static ColumnParser<T> BuildNullableGuidParser<T>(PropertyInfo prop) where T : new()
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
            where T : new()
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
                if (!FastEnum.TryParse(valueString, out TEnum value)
                    || !Enum.TryParse(valueString, ignoreCase: true, out value))
                {
                    return false;
                }
                setter(ref model, value);
                return true;
            };
        }

        private static ColumnParser<T> BuildNullableEnumCore<T, TEnum>(PropertyInfo prop)
            where T : new()
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
                if (!FastEnum.TryParse(valueString, ignoreCase: true, out TEnum parsed) ||
                    !Enum.TryParse(valueString, ignoreCase: true, out parsed))
                {
                    return false;
                }
                setter(ref model, parsed);
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
