using ExcelReader.Core.Reader;

namespace ExcelReader.Native
{
    /// <summary>
    /// Everything one open workbook needs on the managed side of the boundary. The caller only ever
    /// sees an opaque pointer to a GCHandle wrapping this object.
    /// </summary>
    /// <remarks>
    /// <see cref="Scratch"/> holds the most recently serialized row. A row is serialized once and
    /// then copied out, so a caller whose buffer was too small can retry with a bigger one without
    /// losing the row — the reader has already advanced past it and cannot rewind.
    /// </remarks>
    internal sealed class NativeHandle : IDisposable
    {
        internal NativeHandle(IExcelRowReader reader)
        {
            Reader = reader;
            Scratch = new byte[4096];
        }

        internal IExcelRowReader Reader { get; }

        /// <summary>Row cursor over the current sheet. Created lazily on the first row request, dropped on sheet change.</summary>
        internal IExcelRowEnumerator? Rows { get; set; }

        internal byte[] Scratch { get; set; }

        internal int PendingLength { get; set; }

        internal bool HasPending { get; set; }

        internal void ResetRows()
        {
            Rows?.Dispose();
            Rows = null;
            HasPending = false;
            PendingLength = 0;
        }

        public void Dispose()
        {
            ResetRows();
            Reader.Dispose();
        }
    }
}
