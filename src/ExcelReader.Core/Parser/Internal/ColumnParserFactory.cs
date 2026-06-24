using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;
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

        private static readonly HashSet<Type> _parsableTypes =
        [
            typeof(int), typeof(long), typeof(double), typeof(float), typeof(decimal),
            typeof(short), typeof(byte), typeof(uint), typeof(ulong), typeof(ushort),
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
            return (ref model, in cell, isDate1904) =>
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
            return (ref model, in cell, isDate1904) =>
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
            return (ref model, in cell, isDate1904) =>
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

        private static ColumnParser<T> BuildParsableCore<T, TProp>(PropertyInfo prop)
            where T : new()
            where TProp : IUtf8SpanParsable<TProp>
        {
            RefAction<T, TProp> setter = CompileSetter<T, TProp>(prop);
            return (ref model, in cell, _) =>
            {
                if (cell.Type == CellType.Empty)
                {
                    return true;
                }
                if (!cell.TryParse<TProp>(CultureInfo.InvariantCulture, out TProp? value))
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
            return (ref model, in cell, _) =>
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
            return (ref model, in cell, isDate1904) =>
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

        [SuppressMessage("Blocker Code Smell", "S3011:Reflection should not be used to increase accessibility of classes, methods, or fields",
            Justification = "Called via MakeGenericMethod dispatch; private access is intentional and type-safe.")]
        private static ColumnParser<T> BuildNullableParsableCore<T, TProp>(PropertyInfo prop)
            where T : new()
            where TProp : struct, IUtf8SpanParsable<TProp>
        {
            RefAction<T, TProp?> setter = CompileSetter<T, TProp?>(prop);
            return (ref model, in cell, _) =>
            {
                if (cell.Type == CellType.Empty)
                {
                    return true;
                }
                if (!cell.TryParse(CultureInfo.InvariantCulture, out TProp parsed))
                {
                    return false;
                }
                TProp? value = parsed;
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
