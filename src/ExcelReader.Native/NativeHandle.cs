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
    /// <para>
    /// <see cref="LiveSession"/> is the interlock behind the ABI's one-chunked-read-per-workbook rule.
    /// <see cref="IExcelRowReader"/> serves ONE usable <see cref="IExcelRowEnumerator"/> at a time (see
    /// its thread-safety remarks): <c>CsvReader.GetEnumerator</c> and <c>XlsReader.GetEnumerator</c>
    /// rewind the shared source stream, so a second enumerator silently invalidates the first.
    /// <see cref="Rows"/>, the <c>xl_next_row</c> cursor, is also held open across calls on this handle
    /// by design, so <see cref="LiveSession"/> is not the only such enumerator here — it is the only
    /// one that is CALLER-VISIBLE across calls (a <c>xl_typed_reader</c>/Arrow stream handed back to
    /// the caller keeps running until the caller drains or closes it), which is why it alone needs an
    /// explicit interlock: the slot exists here, on the workbook they all share, rather than inside
    /// <see cref="NativeApi.TypedParseSession"/>, so that every OTHER read on the workbook can reach it
    /// to fault a live session before taking the cursor over.
    /// </para>
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

        /// <summary>The one caller-visible chunked read open on this workbook — an
        /// <c>xl_typed_reader</c> or an <c>xl_parse_arrow_stream</c> — or <see langword="null"/> when
        /// there is none. Set by <see cref="NativeApi.TypedParseSession.Open"/>, cleared by that
        /// session's dispose or fault. See the class remarks for why one is the limit.</summary>
        internal NativeApi.TypedParseSession? LiveSession { get; set; }

        /// <summary>
        /// Invalidates the live chunked read, if any, because <paramref name="cause"/> is about to take
        /// over the workbook's row cursor. The session latches the reason and reports it from every
        /// later <c>xl_typed_reader_next</c>/<c>get_next</c>; it never silently resumes at whatever row
        /// <paramref name="cause"/> leaves the cursor on.
        /// </summary>
        /// <param name="cause">The ABI function taking over, named as the caller knows it.</param>
        internal void FaultLiveSession(string cause)
        {
            LiveSession?.Fault($"this chunked read was invalidated by {cause} on the same workbook: a " +
                "workbook serves one row cursor at a time, so this read's position is no longer defined. " +
                "Finish or close the read before using the workbook for anything else.");
        }

        /// <summary>Releases the live-session slot if <paramref name="session"/> still owns it. Identity-checked
        /// because a faulted session is detached immediately, so the slot may already belong to a newer one by
        /// the time the faulted one is closed.</summary>
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
            // A chunked read outliving its workbook is the documented "close the workbook first" case
            // (excelreader.h, chunked typed reading): faulted here, explicitly, so the next call on it
            // is a clean XL_ERROR with a message. Leaving it to the disposed reader to throw is not
            // enough — a CSV enumerator sitting on already-buffered bytes keeps returning rows from a
            // closed workbook instead.
            FaultLiveSession("xl_close");
            ResetRows();
            Reader.Dispose();
        }
    }
}
