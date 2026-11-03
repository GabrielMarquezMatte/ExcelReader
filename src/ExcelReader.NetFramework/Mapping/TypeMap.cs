using System.Reflection;
using ExcelReader.NetFramework.Interop;

namespace ExcelReader.NetFramework.Mapping
{
    internal sealed unsafe class TypeMap<T>
        where T : new()
    {
        private static TypeMap<T>? _instance;

        private TypeMap()
        {
            if (typeof(T).IsValueType)
            {
                throw new NotSupportedException($"'{typeof(T).Name}' is a struct; typed reads fill classes only.");
            }

            List<PropertyBinder<T>> properties = new List<PropertyBinder<T>>();
            foreach (PropertyInfo property in typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                if (Attribute.IsDefined(property, typeof(ExcelIgnoreAttribute)) || property.GetSetMethod() is null)
                {
                    continue;
                }

                properties.Add(Bind(property));
            }

            Properties = properties.AsReadOnly();
        }

        internal static TypeMap<T> Instance => _instance ??= new TypeMap<T>();

        internal IReadOnlyList<PropertyBinder<T>> Properties { get; }

        private static PropertyBinder<T> Bind(PropertyInfo property)
        {
            string[] names = property.GetCustomAttributes<ExcelColumnAttribute>().Select(a => a.Name).ToArray();
            if (names.Length == 0)
            {
                names = new[] { property.Name };
            }

            ExcelRequiredAttribute? required = property.GetCustomAttribute<ExcelRequiredAttribute>();
            Type type = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;
            PropertyBinder<T>? binder = Create(property, type, names, required);
            if (binder is null)
            {
                throw new NotSupportedException($"Property '{typeof(T).Name}.{property.Name}' has type '{property.PropertyType}', which typed reads do not support.");
            }

            return binder;
        }

        private static PropertyBinder<T>? Create(PropertyInfo property, Type type, string[] names, ExcelRequiredAttribute? required)
        {
            if (type == typeof(TimeSpan))
            {
                return Bind<TimeSpan>(property, names, Xl.XL_T_TIME, required, ColumnReaders.ReadTimeSpan);
            }

            if (type.IsEnum)
            {
                return null;
            }

            return Type.GetTypeCode(type) switch
            {
                TypeCode.String => new PropertyBinder<T, string?>(property, names, Xl.XL_T_STRING, required, ColumnReaders.ReadString),
                TypeCode.Int64 => Bind<long>(property, names, Xl.XL_T_I64, required, ColumnReaders.ReadInt64),
                TypeCode.Int32 => Bind<int>(property, names, Xl.XL_T_I64, required, ColumnReaders.ReadInt32),
                TypeCode.Int16 => Bind<short>(property, names, Xl.XL_T_I64, required, ColumnReaders.ReadInt16),
                TypeCode.Byte => Bind<byte>(property, names, Xl.XL_T_I64, required, ColumnReaders.ReadByte),
                TypeCode.Double => Bind<double>(property, names, Xl.XL_T_F64, required, ColumnReaders.ReadDouble),
                TypeCode.Single => Bind<float>(property, names, Xl.XL_T_F64, required, ColumnReaders.ReadSingle),
                TypeCode.Decimal => Bind<decimal>(property, names, Xl.XL_T_F64, required, ColumnReaders.ReadDecimal),
                TypeCode.Boolean => Bind<bool>(property, names, Xl.XL_T_BOOL, required, ColumnReaders.ReadBoolean),
                TypeCode.DateTime => Bind<DateTime>(property, names, Xl.XL_T_TIMESTAMP, required, ColumnReaders.ReadDateTime),
                _ => null,
            };
        }

        private static PropertyBinder<T> Bind<TValue>(PropertyInfo property, string[] names, int nativeType, ExcelRequiredAttribute? required, ColumnReaders.Reader<TValue> read)
            where TValue : struct
        {
            if (property.PropertyType == typeof(TValue))
            {
                return new PropertyBinder<T, TValue>(property, names, nativeType, required, read);
            }

            return new PropertyBinder<T, TValue?>(property, names, nativeType, required, ColumnReaders.Lift(read));
        }
    }
}
