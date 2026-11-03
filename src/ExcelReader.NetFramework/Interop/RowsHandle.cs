using ExcelReader.NetFramework.Reader;
using Microsoft.Win32.SafeHandles;

namespace ExcelReader.NetFramework.Interop
{
    internal sealed unsafe class RowsHandle : SafeHandleZeroOrMinusOneIsInvalid
    {
        internal RowsHandle() : base(true)
        {
        }

        internal static RowsHandle Open(ExcelWorkbook workbook, int sheet)
        {
            WorkbookHandle workbookHandle = workbook.Handle;
            bool added = false;
            try
            {
                workbookHandle.DangerousAddRef(ref added);
                IntPtr rows = IntPtr.Zero;
                NativeError.ThrowIfFailed(NativeMethods.xl_rows_open(workbookHandle.DangerousGetHandle(), sheet, &rows));
                RowsHandle result = new RowsHandle();
                result.SetHandle(rows);
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

        internal ExcelRow? Next()
        {
            bool added = false;
            try
            {
                DangerousAddRef(ref added);
                XlRow view = default;
                int status = NativeMethods.xl_rows_next_view(handle, &view);
                if (status == (int)ExcelStatus.Eof)
                {
                    return null;
                }

                NativeError.ThrowIfFailed(status);
                return Copy(view);
            }
            finally
            {
                if (added)
                {
                    DangerousRelease();
                }
            }
        }

        protected override bool ReleaseHandle()
        {
            return NativeMethods.xl_rows_close(handle) == (int)ExcelStatus.Ok;
        }

        private static ExcelRow Copy(XlRow view)
        {
            ExcelCell[] cells = new ExcelCell[view.CellCount];
            for (int i = 0; i < cells.Length; i++)
            {
                XlRowCell cell = view.Cells[i];
                cells[i] = new ExcelCell(cell.Column, (ExcelCellType)cell.Type, Utf8.Decode(cell.Value, cell.ValueLen));
            }

            return new ExcelRow(cells);
        }
    }
}
