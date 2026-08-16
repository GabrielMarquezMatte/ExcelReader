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

        /// <summary>Backs xl_read_all_blob: every remaining row of the sheet, concatenated, from the last
        /// accumulation — the repeated <c>int32 row_length, row blob</c> entries only, WITHOUT the
        /// leading row count, which is written straight into the caller's buffer on the way out (see
        /// <see cref="NativeApi.ReadAllBlob"/>) because a chunked buffer cannot be back-patched at
        /// offset 0. Held across a <see cref="NativeStatus.BufferTooSmall"/> return the same way
        /// <see cref="Scratch"/>/<see cref="HasPending"/> hold a single row — so a caller that retries
        /// with a bigger buffer loses nothing, even though accumulation has already fully drained the
        /// underlying row enumerator by the time the first too-small result comes back.</summary>
        internal ChunkedBuffer<byte>? AllRowsScratch { get; set; }

        /// <summary>Row count for <see cref="AllRowsScratch"/>'s entries, i.e. the blob's leading int32.</summary>
        internal int AllRowsCount { get; set; }

        /// <summary>Byte size of the whole blob a caller must supply room for: the leading count plus
        /// <see cref="AllRowsScratch"/>.</summary>
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
            Reader.Dispose();
        }
    }
}
