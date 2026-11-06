using System.Runtime.InteropServices;
using Apache.Arrow;
using Apache.Arrow.Types;
using ExcelReader.Core.Reader.Columnar;
using ExcelReader.Core.Reader.Schema;

namespace ExcelReader.Arrow
{
    /// <summary>
    /// Copies a built column's native blocks into Arrow arrays. Apache.Arrow takes managed memory
    /// only through public API, so each block is copied once and freed with the table.
    /// </summary>
    internal static class ArrowColumns
    {
        private static readonly Time64Type MicrosecondTime = new(TimeUnit.Microsecond);
        private static readonly TimestampType MicrosecondTimestamp = new(TimeUnit.Microsecond, timezone: (string?)null);
        private static readonly DictionaryType StringDictionary = new(Int32Type.Default, StringType.Default, ordered: false);

        internal static IArrowType TypeOf(ExcelColumnType type)
        {
            return type switch
            {
                ExcelColumnType.StringColumn => StringType.Default,
                ExcelColumnType.StringDictionaryColumn => StringDictionary,
                ExcelColumnType.Int64Column => Int64Type.Default,
                ExcelColumnType.Float64Column => DoubleType.Default,
                ExcelColumnType.BoolColumn => BooleanType.Default,
                ExcelColumnType.DateColumn => Date32Type.Default,
                ExcelColumnType.TimeColumn => MicrosecondTime,
                _ => MicrosecondTimestamp,
            };
        }

        internal static IArrowArray ToArray(in BuiltColumn column)
        {
            int length = checked((int)column.Length);
            ArrowBuffer validity = column.Validity == IntPtr.Zero ? ArrowBuffer.Empty : Copy(column.Validity, (length + 7) / 8);
            int nullCount = checked((int)ColumnBlocks.CountUnset(column.Validity, length));
            return column.Type switch
            {
                ExcelColumnType.StringColumn => new StringArray(length, Copy(column.Values, (length + 1) * sizeof(int)), Copy(column.Data, column.DataLen), validity, nullCount),
                ExcelColumnType.StringDictionaryColumn => Dictionary(in column, length, validity, nullCount),
                ExcelColumnType.Int64Column => new Int64Array(Copy(column.Values, length * sizeof(long)), validity, length, nullCount, 0),
                ExcelColumnType.Float64Column => new DoubleArray(Copy(column.Values, length * sizeof(double)), validity, length, nullCount, 0),
                ExcelColumnType.BoolColumn => new BooleanArray(PackBits(column.Values, length), validity, length, nullCount, 0),
                ExcelColumnType.DateColumn => new Date32Array(Copy(column.Values, length * sizeof(int)), validity, length, nullCount, 0),
                ExcelColumnType.TimeColumn => new Time64Array(MicrosecondTime, Copy(column.Values, length * sizeof(long)), validity, length, nullCount, 0),
                _ => new TimestampArray(MicrosecondTimestamp, Copy(column.Values, length * sizeof(long)), validity, length, nullCount, 0),
            };
        }

        private static DictionaryArray Dictionary(in BuiltColumn column, int length, ArrowBuffer validity, int nullCount)
        {
            int entries = checked((int)column.DictionaryLength);
            Int32Array indices = new(Copy(column.Values, length * sizeof(int)), validity, length, nullCount, 0);
            StringArray values = new(entries, Copy(column.DictionaryOffsets, (entries + 1) * sizeof(int)), Copy(column.Data, column.DataLen), ArrowBuffer.Empty, 0);
            return new DictionaryArray(StringDictionary, indices, values);
        }

        private static ArrowBuffer Copy(IntPtr block, long byteLength)
        {
            byte[] bytes = GC.AllocateUninitializedArray<byte>(checked((int)byteLength));
            Marshal.Copy(block, bytes, 0, bytes.Length);
            return new ArrowBuffer(bytes);
        }

        private static ArrowBuffer PackBits(IntPtr flags, int length)
        {
            byte[] values = new byte[length];
            Marshal.Copy(flags, values, 0, length);
            byte[] bits = new byte[(length + 7) / 8];
            for (int i = 0; i < length; i++)
            {
                if (values[i] != 0)
                {
                    bits[i >> 3] |= (byte)(1 << (i & 7));
                }
            }
            return new ArrowBuffer(bits);
        }
    }
}
