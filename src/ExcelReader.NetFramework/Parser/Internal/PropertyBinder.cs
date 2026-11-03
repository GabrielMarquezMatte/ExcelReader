using System.Linq.Expressions;
using System.Reflection;
using ExcelReader.NetFramework.Interop;
using ExcelReader.NetFramework.Reader;

namespace ExcelReader.NetFramework.Parser.Internal
{
    internal abstract class PropertyBinder<T>
    {
        protected PropertyBinder(string propertyName, string[] names, int nativeType, bool required, bool allowEmpty)
        {
            PropertyName = propertyName;
            Names = names;
            NativeType = nativeType;
            Required = required;
            AllowEmpty = allowEmpty;
        }

        internal string PropertyName { get; }

        internal string[] Names { get; }

        internal int NativeType { get; }

        internal bool Required { get; }

        internal bool AllowEmpty { get; }

        internal abstract unsafe void Fill(T[] items, int offset, int rowBase, XlColumn* column);
    }

    internal sealed unsafe class PropertyBinder<T, TValue> : PropertyBinder<T>
    {
        private readonly Action<T, TValue> _set;
        private readonly ColumnReaders.Reader<TValue> _read;

        internal PropertyBinder(PropertyInfo property, string[] names, int nativeType, ExcelRequiredAttribute? required, ColumnReaders.Reader<TValue> read)
            : base(property.Name, names, nativeType, required is not null, required?.AllowEmpty == true)
        {
            _read = read;
            ParameterExpression target = Expression.Parameter(typeof(T), "target");
            ParameterExpression value = Expression.Parameter(typeof(TValue), "value");
            _set = Expression.Lambda<Action<T, TValue>>(Expression.Assign(Expression.Property(target, property), value), target, value).Compile();
        }

        internal override void Fill(T[] items, int offset, int rowBase, XlColumn* column)
        {
            int count = checked((int)column->Length);
            int row = 0;
            try
            {
                for (; row < count; row++)
                {
                    if (column->IsValid(row))
                    {
                        _set(items[offset + row], _read(column, row));
                    }
                }
            }
            catch (OverflowException)
            {
                throw new ExcelReaderException(ExcelStatus.InvalidArgument, $"The value in data row {rowBase + row + 1} does not fit property '{PropertyName}' ({typeof(TValue).Name}).");
            }
        }
    }
}
