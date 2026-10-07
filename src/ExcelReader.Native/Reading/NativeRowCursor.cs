using ExcelReader.Core.Reader;

namespace ExcelReader.Native.Reading
{
    internal sealed class NativeRowCursor : IDisposable
    {
        internal NativeRowCursor(IExcelRowEnumerator rows)
        {
            Rows = rows;
            Scratch = new byte[4096];
        }

        internal IExcelRowEnumerator Rows { get; }

        internal byte[] Scratch { get; set; }

        internal int PendingLength { get; set; }

        internal bool HasPending { get; set; }

        internal RowViewBuffer? View { get; set; }

        internal ChunkedBuffer<byte>? AllRowsScratch { get; set; }

        internal int AllRowsCount { get; set; }

        internal int AllRowsLength { get; set; }

        internal bool AllRowsPending { get; set; }

        public void Dispose()
        {
            Rows.Dispose();
            View?.Dispose();
        }
    }
}
