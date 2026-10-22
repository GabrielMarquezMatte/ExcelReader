using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Reader;

namespace ExcelReader.Native.Reading
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeTableInfo
    {
        public int Sheet;
        public int FirstRow;
        public int LastRow;
        public int FirstColumn;
        public int LastColumn;
        public int HeaderRowCount;
        public int TotalsRowCount;
    }

    internal static partial class ReadApi
    {
        // XL_TABLE_SHEET(table) in excelreader.h: tables take the sheet indexes below -1.
        internal const int FirstTableSheet = -2;

        private static ExcelTable TableAt(NativeHandle handle, int table)
        {
            IReadOnlyList<ExcelTable> tables = handle.Workbook.Tables;
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(table, tables.Count, nameof(table));
            return tables[table];
        }

        internal static int TableCount(NativeHandle? handle, out int count)
        {
            count = 0;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }
            NativeApi.ClearLastError();
            try
            {
                count = handle.Workbook.Tables.Count;
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int TableInfoAt(NativeHandle? handle, int index, out NativeTableInfo info)
        {
            info = default;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }
            NativeApi.ClearLastError();
            if (index < 0)
            {
                NativeApi.SetLastError($"table must be a zero-based index; got {index}.");
                return NativeStatus.InvalidArgument;
            }
            try
            {
                ExcelTable table = TableAt(handle, index);
                info = new NativeTableInfo
                {
                    Sheet = table.Sheet.Index,
                    FirstRow = table.FirstRow,
                    LastRow = table.LastRow,
                    FirstColumn = table.FirstColumn,
                    LastColumn = table.LastColumn,
                    HeaderRowCount = table.HeaderRowCount,
                    TotalsRowCount = table.TotalsRowCount,
                };
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int TableNameAt(NativeHandle? handle, int index, Span<byte> buffer, out int length)
        {
            length = 0;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }
            NativeApi.ClearLastError();
            if (index < 0)
            {
                NativeApi.SetLastError($"table must be a zero-based index; got {index}.");
                return NativeStatus.InvalidArgument;
            }
            try
            {
                return CopyUtf8(TableAt(handle, index).Name, buffer, out length);
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int TableIndex(NativeHandle? handle, ReadOnlySpan<byte> utf8Name, out int index)
        {
            index = -1;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }
            NativeApi.ClearLastError();
            try
            {
                IExcelWorkbook workbook = handle.Workbook;
                if (workbook.TryGetTable(Encoding.UTF8.GetString(utf8Name), out ExcelTable? table))
                {
                    index = IndexOf(workbook.Tables, table);
                }
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        private static int IndexOf(IReadOnlyList<ExcelTable> tables, ExcelTable table)
        {
            for (int i = 0; i < tables.Count; i++)
            {
                if (ReferenceEquals(tables[i], table))
                {
                    return i;
                }
            }
            return -1;
        }
    }
}
