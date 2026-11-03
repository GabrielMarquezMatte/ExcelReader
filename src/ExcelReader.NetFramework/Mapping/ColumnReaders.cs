using ExcelReader.NetFramework.Interop;

namespace ExcelReader.NetFramework.Mapping
{
    internal static unsafe class ColumnReaders
    {
        private const long UnixEpochTicks = 621355968000000000;
        private const long TicksPerMicrosecond = 10;

        internal delegate TValue Reader<TValue>(XlColumn* column, int row);

        internal static Reader<TValue?> Lift<TValue>(Reader<TValue> read)
            where TValue : struct
        {
            return (column, row) => read(column, row);
        }

        internal static string? ReadString(XlColumn* column, int row)
        {
            int* offsets = (int*)column->Values;
            int start = offsets[row];
            return Utf8.Decode(column->Data + start, offsets[row + 1] - start);
        }

        internal static long ReadInt64(XlColumn* column, int row)
        {
            return ((long*)column->Values)[row];
        }

        internal static int ReadInt32(XlColumn* column, int row)
        {
            return checked((int)ReadInt64(column, row));
        }

        internal static short ReadInt16(XlColumn* column, int row)
        {
            return checked((short)ReadInt64(column, row));
        }

        internal static byte ReadByte(XlColumn* column, int row)
        {
            return checked((byte)ReadInt64(column, row));
        }

        internal static double ReadDouble(XlColumn* column, int row)
        {
            return ((double*)column->Values)[row];
        }

        internal static float ReadSingle(XlColumn* column, int row)
        {
            return (float)ReadDouble(column, row);
        }

        internal static decimal ReadDecimal(XlColumn* column, int row)
        {
            return (decimal)ReadDouble(column, row);
        }

        internal static bool ReadBoolean(XlColumn* column, int row)
        {
            return ((byte*)column->Values)[row] != 0;
        }

        internal static DateTime ReadDateTime(XlColumn* column, int row)
        {
            return new DateTime(checked((ReadInt64(column, row) * TicksPerMicrosecond) + UnixEpochTicks), DateTimeKind.Unspecified);
        }

        internal static TimeSpan ReadTimeSpan(XlColumn* column, int row)
        {
            return new TimeSpan(checked(ReadInt64(column, row) * TicksPerMicrosecond));
        }
    }
}
