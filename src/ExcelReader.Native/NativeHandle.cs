using ExcelReader.Core.Reader;

namespace ExcelReader.Native
{
    internal sealed class NativeHandle : IDisposable
    {
        internal NativeHandle(IExcelRowReader reader)
        {
            Reader = reader;
            Scratch = new byte[4096];
        }

        internal IExcelRowReader Reader { get; }

        internal IExcelRowEnumerator? Rows { get; set; }

        internal byte[] Scratch { get; set; }

        internal int PendingLength { get; set; }

        internal bool HasPending { get; set; }

        internal ChunkedBuffer<byte>? AllRowsScratch { get; set; }

        internal int AllRowsCount { get; set; }

        internal int AllRowsLength { get; set; }

        internal bool AllRowsPending { get; set; }

        internal NativeApi.TypedParseSession? LiveSession { get; set; }

        internal void FaultLiveSession(string cause)
        {
            LiveSession?.Fault($"this chunked read was invalidated by {cause} on the same workbook: a " +
                "workbook serves one row cursor at a time, so this read's position is no longer defined. " +
                "Finish or close the read before using the workbook for anything else.");
        }

        internal void ReleaseLiveSession(NativeApi.TypedParseSession session)
        {
            if (ReferenceEquals(LiveSession, session))
            {
                LiveSession = null;
            }
        }

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
            FaultLiveSession("xl_close");
            ResetRows();
            Reader.Dispose();
        }
    }
}
