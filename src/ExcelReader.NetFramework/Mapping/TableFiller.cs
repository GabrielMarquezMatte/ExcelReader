using ExcelReader.NetFramework.Interop;

namespace ExcelReader.NetFramework.Mapping
{
    internal static unsafe class TableFiller
    {
        internal static void Fill<T>(PropertyBinder<T>[] bound, in XlTable table, T[] items, int offset, int rowBase)
        {
            if (table.ColumnCount != bound.Length)
            {
                throw new ExcelReaderException(ExcelStatus.Error, $"The native library returned {table.ColumnCount} columns for {bound.Length} column specs.");
            }

            for (int i = 0; i < bound.Length; i++)
            {
                XlColumn* column = table.Columns + i;
                if (column->Length != table.RowCount)
                {
                    throw new ExcelReaderException(ExcelStatus.Error, $"The native library returned {column->Length} rows for property '{bound[i].PropertyName}', expected {table.RowCount}.");
                }

                if (column->Type != bound[i].NativeType)
                {
                    throw new ExcelReaderException(ExcelStatus.Error, $"The native library returned column type {column->Type} for property '{bound[i].PropertyName}', expected {bound[i].NativeType}.");
                }
            }

            for (int i = 0; i < bound.Length; i++)
            {
                bound[i].Fill(items, offset, rowBase, table.Columns + i);
            }

            for (int i = 0; i < bound.Length; i++)
            {
                if (bound[i].Required && !bound[i].AllowEmpty)
                {
                    ThrowOnFirstEmpty(bound[i].PropertyName, table.Columns + i, rowBase);
                }
            }
        }

        private static void ThrowOnFirstEmpty(string propertyName, XlColumn* column, int rowBase)
        {
            int count = checked((int)column->Length);
            for (int row = 0; row < count; row++)
            {
                if (IsEmpty(column, row))
                {
                    throw new ExcelReaderException(ExcelStatus.InvalidArgument, $"Required property '{propertyName}' is empty in data row {rowBase + row + 1}.");
                }
            }
        }

        private static bool IsEmpty(XlColumn* column, int row)
        {
            if (!column->IsValid(row))
            {
                return true;
            }

            if (column->Type != Xl.XL_T_STRING)
            {
                return false;
            }

            int* offsets = (int*)column->Values;
            return offsets[row] == offsets[row + 1];
        }
    }
}
