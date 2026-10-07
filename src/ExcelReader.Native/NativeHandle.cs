using ExcelReader.Core.Reader;
using ExcelReader.Native.Reading;

namespace ExcelReader.Native
{
    internal sealed class NativeHandle : IDisposable
    {
        internal NativeHandle(IExcelWorkbook workbook)
        {
            Workbook = workbook;
            Scratch = new byte[4096];
        }

        internal IExcelWorkbook Workbook { get; }

        internal int CurrentSheet { get; private set; }

        internal IExcelSheet Sheet
        {
            get
            {
                return Workbook.SheetAt(CurrentSheet);
            }
        }

        internal void MoveToSheet(int index)
        {
            _ = Workbook.SheetAt(index);
            CurrentSheet = index;
            ResetRows();
        }

        internal IExcelRowEnumerator? Rows { get; set; }

        internal byte[] Scratch { get; set; }

        internal int PendingLength { get; set; }

        internal bool HasPending { get; set; }

        internal RowViewBuffer? View { get; set; }

        internal ChunkedBuffer<byte>? AllRowsScratch { get; set; }

        internal int AllRowsCount { get; set; }

        internal int AllRowsLength { get; set; }

        internal bool AllRowsPending { get; set; }

        internal void ResetRows()
        {
            Rows?.Dispose();
            Rows = null;
            HasPending = false;
            PendingLength = 0;
            AllRowsPending = false;
            AllRowsLength = 0;
            AllRowsCount = 0;
            AllRowsScratch = null;
        }

        public void Dispose()
        {
            ResetRows();
            View?.Dispose();
            Workbook.Dispose();
        }
    }
}
