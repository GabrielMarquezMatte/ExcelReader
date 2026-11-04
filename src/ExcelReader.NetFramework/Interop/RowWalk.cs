using System.Diagnostics.CodeAnalysis;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using ExcelReader.NetFramework.Reader;

namespace ExcelReader.NetFramework.Interop
{
    internal sealed unsafe class RowWalk
    {
        private const int CallbackFailed = int.MaxValue;

        private static readonly RowFn RowCallback = OnRow;
        private static readonly IntPtr RowPointer = Marshal.GetFunctionPointerForDelegate(RowCallback);

        private readonly ExcelRowAction _action;
        private ExceptionDispatchInfo? _error;

        private RowWalk(ExcelRowAction action)
        {
            _action = action;
        }

        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate int RowFn(XlCellView* cells, int cellCount, IntPtr userData);

        internal static void Run(ExcelWorkbook workbook, int sheet, ExcelRowAction action)
        {
            RowWalk walk = new RowWalk(action);
            GCHandle self = GCHandle.Alloc(walk);
            WorkbookHandle handle = workbook.Handle;
            bool added = false;
            try
            {
                handle.DangerousAddRef(ref added);
                int status = NativeMethods.xl_rows_foreach(handle.DangerousGetHandle(), sheet, RowPointer, GCHandle.ToIntPtr(self));
                walk._error?.Throw();
                NativeError.ThrowIfFailed(status);
            }
            finally
            {
                self.Free();
                if (added)
                {
                    handle.DangerousRelease();
                }
            }
        }

        [SuppressMessage("Design", "CA1031:Do not catch general exception types",
            Justification = "An exception must not cross the native callback boundary; it is rethrown on return.")]
        private static int OnRow(XlCellView* cells, int cellCount, IntPtr userData)
        {
            RowWalk walk = (RowWalk)GCHandle.FromIntPtr(userData).Target!;
            try
            {
                walk._action(new ExcelRowView(cells, cellCount));
                return (int)ExcelStatus.Ok;
            }
            catch (Exception exception)
            {
                walk._error = ExceptionDispatchInfo.Capture(exception);
                return CallbackFailed;
            }
        }
    }
}
