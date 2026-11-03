using System.Reflection;
using ExcelReader.NetFramework.Interop;
using ExcelReader.NetFramework.Parser;

namespace ExcelReader.NetFramework.Writer.Internal
{
    internal sealed class WriteMap<T>
    {
        private static WriteMap<T>? _instance;

        private WriteMap()
        {
            IEnumerable<PropertyInfo> properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.GetGetMethod() is not null && p.GetIndexParameters().Length == 0 && !Attribute.IsDefined(p, typeof(ExcelIgnoreAttribute)))
                .OrderBy(p => Depth(p.DeclaringType!))
                .ThenBy(p => p.MetadataToken);
            Columns = properties.Select(Create).ToArray();
        }

        internal static WriteMap<T> Instance
        {
            get
            {
                return _instance ??= new WriteMap<T>();
            }
        }

        internal ColumnWriter<T>[] Columns { get; }

        private static int Depth(Type type)
        {
            int depth = 0;
            for (Type? current = type.BaseType; current is not null; current = current.BaseType)
            {
                depth++;
            }

            return depth;
        }

        private static ColumnWriter<T> Create(PropertyInfo property)
        {
            ExcelColumnAttribute? column = property.GetCustomAttributes<ExcelColumnAttribute>().FirstOrDefault();
            string header = column?.Name ?? property.Name;
            Type type = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;
            bool nullable = type != property.PropertyType;
            ColumnWriter<T>? writer = Create(property, header, type, nullable);
            if (writer is null)
            {
                throw new NotSupportedException($"Property '{typeof(T).Name}.{property.Name}' has type '{property.PropertyType}', which typed writes do not support.");
            }

            return writer;
        }

        private static ColumnWriter<T>? Create(PropertyInfo property, string header, Type type, bool nullable)
        {
            if (type == typeof(TimeSpan))
            {
                return new ValueColumnWriter<T, TimeSpan, long>(property, header, Xl.XL_T_TIME, nullable, ColumnWriters.Time);
            }

            if (type.IsEnum)
            {
                return null;
            }

            return Type.GetTypeCode(type) switch
            {
                TypeCode.String => new StringColumnWriter<T>(property, header),
                TypeCode.Int64 or TypeCode.Int32 or TypeCode.Int16 or TypeCode.Byte => new ValueColumnWriter<T, long, long>(property, header, Xl.XL_T_I64, nullable, ColumnWriters.Int64),
                TypeCode.Double or TypeCode.Single or TypeCode.Decimal => new ValueColumnWriter<T, double, double>(property, header, Xl.XL_T_F64, nullable, ColumnWriters.Float64),
                TypeCode.Boolean => new ValueColumnWriter<T, bool, byte>(property, header, Xl.XL_T_BOOL, nullable, ColumnWriters.Bool),
                TypeCode.DateTime => new ValueColumnWriter<T, DateTime, long>(property, header, Xl.XL_T_TIMESTAMP, nullable, ColumnWriters.Timestamp),
                _ => null,
            };
        }
    }
}
