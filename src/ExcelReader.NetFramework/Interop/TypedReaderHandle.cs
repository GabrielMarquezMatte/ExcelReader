using ExcelReader.NetFramework.Mapping;
using Microsoft.Win32.SafeHandles;

namespace ExcelReader.NetFramework.Interop
{
    internal sealed unsafe class TypedReaderHandle : SafeHandleZeroOrMinusOneIsInvalid
    {
        private int _rowsRead;

        internal TypedReaderHandle() : base(true)
        {
        }

        internal static TypedReaderHandle Open(ExcelWorkbook workbook, int sheet, NativeSpecs specs, int headerRow, int batchRows)
        {
            WorkbookHandle workbookHandle = workbook.Handle;
            bool added = false;
            try
            {
                workbookHandle.DangerousAddRef(ref added);
                IntPtr reader = IntPtr.Zero;
                NativeError.ThrowIfFailed(NativeMethods.xl_typed_reader_open(workbookHandle.DangerousGetHandle(), sheet, specs.Pointer, specs.Count, headerRow, batchRows, &reader));
                TypedReaderHandle result = new TypedReaderHandle();
                result.SetHandle(reader);
                return result;
            }
            finally
            {
                if (added)
                {
                    workbookHandle.DangerousRelease();
                }
            }
        }

        internal T[]? Next<T>(PropertyBinder<T>[] bound)
            where T : new()
        {
            bool added = false;
            XlTable table = default;
            try
            {
                DangerousAddRef(ref added);
                int status = NativeMethods.xl_typed_reader_next(handle, &table);
                if (status == (int)ExcelStatus.Eof)
                {
                    return null;
                }

                NativeError.ThrowIfFailed(status);
                T[] items = new T[checked((int)table.RowCount)];
                for (int i = 0; i < items.Length; i++)
                {
                    items[i] = new T();
                }

                TableFiller.Fill(bound, in table, items, 0, _rowsRead);
                _rowsRead = checked(_rowsRead + items.Length);
                return items;
            }
            finally
            {
                NativeMethods.xl_free_table(&table);
                if (added)
                {
                    DangerousRelease();
                }
            }
        }

        protected override bool ReleaseHandle()
        {
            NativeMethods.xl_typed_reader_close(handle);
            return true;
        }
    }
}
