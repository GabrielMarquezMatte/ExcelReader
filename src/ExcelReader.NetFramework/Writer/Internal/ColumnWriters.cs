using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using ExcelReader.NetFramework.Interop;
using ExcelReader.NetFramework.Reader;

namespace ExcelReader.NetFramework.Writer.Internal
{
    internal static class ColumnWriters
    {
        private const long UnixEpochTicks = 621355968000000000;
        private const long TicksPerMicrosecond = 10;

        internal static long Int64(long value)
        {
            return value;
        }

        internal static double Float64(double value)
        {
            return value;
        }

        internal static byte Bool(bool value)
        {
            return value ? (byte)1 : (byte)0;
        }

        internal static long Timestamp(DateTime value)
        {
            return (value.Ticks - UnixEpochTicks) / TicksPerMicrosecond;
        }

        internal static long Time(TimeSpan value)
        {
            if (value < TimeSpan.Zero || value >= TimeSpan.FromDays(1))
            {
                throw new ArgumentOutOfRangeException(nameof(value), value, "A time of day is at least zero and less than one day.");
            }

            return value.Ticks / TicksPerMicrosecond;
        }
    }

    internal abstract unsafe class ColumnWriter<T>
    {
        protected ColumnWriter(PropertyInfo property, string header, int nativeType, bool nullable)
        {
            PropertyName = property.Name;
            Header = header;
            NativeType = nativeType;
            Nullable = nullable;
            Style = property.GetCustomAttribute<ExcelStyleAttribute>()?.ToCellStyle();
        }

        internal string PropertyName { get; }

        internal string Header { get; }

        internal int NativeType { get; }

        internal bool Nullable { get; }

        internal CellStyle? Style { get; }

        internal abstract void Fill(T[] rows, XlColumn* column, NativeWriteTable owner);

        protected static Func<T, TResult> Getter<TResult>(PropertyInfo property)
        {
            ParameterExpression target = Expression.Parameter(typeof(T), "target");
            Expression value = Expression.Convert(Expression.Property(target, property), typeof(TResult));
            return Expression.Lambda<Func<T, TResult>>(value, target).Compile();
        }

        protected static void MarkValid(byte* validity, int row)
        {
            validity[row >> 3] |= (byte)(1 << (row & 7));
        }
    }

    internal sealed unsafe class ValueColumnWriter<T, TValue, TNative> : ColumnWriter<T>
        where TValue : struct
        where TNative : unmanaged
    {
        private readonly Func<T, TValue?> _get;
        private readonly Func<TValue, TNative> _convert;

        internal ValueColumnWriter(PropertyInfo property, string header, int nativeType, bool nullable, Func<TValue, TNative> convert)
            : base(property, header, nativeType, nullable)
        {
            _get = Getter<TValue?>(property);
            _convert = convert;
        }

        internal override void Fill(T[] rows, XlColumn* column, NativeWriteTable owner)
        {
            int count = rows.Length;
            TNative* values = (TNative*)owner.Allocate((long)count * sizeof(TNative));
            byte* validity = null;
            if (Nullable)
            {
                validity = owner.AllocateValidity(count);
            }

            for (int row = 0; row < count; row++)
            {
                TValue? value = _get(rows[row]);
                if (!value.HasValue)
                {
                    values[row] = default;
                    continue;
                }

                values[row] = _convert(value.GetValueOrDefault());
                if (validity is not null)
                {
                    MarkValid(validity, row);
                }
            }

            *column = new XlColumn { Type = NativeType, Length = count, Values = values, Validity = validity };
        }
    }

    internal sealed unsafe class StringColumnWriter<T> : ColumnWriter<T>
    {
        private readonly Func<T, string?> _get;

        internal StringColumnWriter(PropertyInfo property, string header)
            : base(property, header, Xl.XL_T_STRING, true)
        {
            _get = Getter<string?>(property);
        }

        internal override void Fill(T[] rows, XlColumn* column, NativeWriteTable owner)
        {
            int count = rows.Length;
            string?[] texts = new string?[count];
            int total = Collect(rows, texts);
            int* offsets = (int*)owner.Allocate((count + 1L) * sizeof(int));
            byte* data = owner.Allocate(total);
            byte* validity = owner.AllocateValidity(count);
            int position = 0;
            offsets[0] = 0;
            for (int row = 0; row < count; row++)
            {
                string? text = texts[row];
                if (text is not null)
                {
                    position += Encode(text, data + position, total - position);
                    MarkValid(validity, row);
                }

                offsets[row + 1] = position;
            }

            *column = new XlColumn { Type = Xl.XL_T_STRING, Length = count, Values = offsets, Validity = validity, Data = data, DataLen = total };
        }

        private static int Encode(string text, byte* destination, int capacity)
        {
            fixed (char* chars = text)
            {
                return Encoding.UTF8.GetBytes(chars, text.Length, destination, capacity);
            }
        }

        private int Collect(T[] rows, string?[] texts)
        {
            long total = 0;
            for (int row = 0; row < rows.Length; row++)
            {
                string? text = _get(rows[row]);
                texts[row] = text;
                if (text is not null)
                {
                    total += Encoding.UTF8.GetByteCount(text);
                }
            }

            if (total > int.MaxValue)
            {
                throw new ExcelReaderException(ExcelStatus.InvalidArgument, $"Property '{typeof(T).Name}.{PropertyName}' holds {total} bytes of UTF-8 text; a column holds at most {int.MaxValue}.");
            }

            return (int)total;
        }
    }
}
