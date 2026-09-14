using ExcelReader.Core.Reader;

namespace ExcelReader.Native
{
    // Everything one open workbook needs on the managed side of the boundary. The caller only ever
    // sees an opaque pointer to a GCHandle wrapping this object.
    //
    // Scratch holds the most recently serialized row. A row is serialized once and
    // then copied out, so a caller whose buffer was too small can retry with a bigger one without
    // losing the row — the reader has already advanced past it and cannot rewind.
    //
    // LiveSession is the interlock behind the ABI's one-chunked-read-per-workbook rule.
    // IExcelRowReader serves ONE usable IExcelRowEnumerator at a time (see
    // its thread-safety remarks): CsvReader.GetEnumerator and XlsReader.GetEnumerator
    // rewind the shared source stream, so a second enumerator silently invalidates the first.
    // Rows, the xl_next_row cursor, is also held open across calls on this handle
    // by design, so LiveSession is not the only such enumerator here — it is the only
    // one that is CALLER-VISIBLE across calls (a xl_typed_reader/Arrow stream handed back to
    // the caller keeps running until the caller drains or closes it), which is why it alone needs an
    // explicit interlock: the slot exists here, on the workbook they all share, rather than inside
    // NativeApi.TypedParseSession, so that every OTHER read on the workbook can reach it
    // to fault a live session before taking the cursor over.
    internal sealed class NativeHandle : IDisposable
    {
        internal NativeHandle(IExcelRowReader reader)
        {
            Reader = reader;
            Scratch = new byte[4096];
        }

        internal IExcelRowReader Reader { get; }

        // Row cursor over the current sheet. Created lazily on the first row request, dropped on sheet change.
        internal IExcelRowEnumerator? Rows { get; set; }

        internal byte[] Scratch { get; set; }

        internal int PendingLength { get; set; }

        internal bool HasPending { get; set; }

        // Backs xl_read_all_blob: every remaining row of the sheet, concatenated, from the last
        // accumulation — the repeated int32 row_length, row blob entries only, WITHOUT the
        // leading row count, which is written straight into the caller's buffer on the way out (see
        // NativeApi.ReadAllBlob) because a chunked buffer cannot be back-patched at
        // offset 0. Held across a NativeStatus.BufferTooSmall return the same way
        // Scratch/HasPending hold a single row — so a caller that retries
        // with a bigger buffer loses nothing, even though accumulation has already fully drained the
        // underlying row enumerator by the time the first too-small result comes back.
        internal ChunkedBuffer<byte>? AllRowsScratch { get; set; }

        // Row count for AllRowsScratch's entries, i.e. the blob's leading int32.
        internal int AllRowsCount { get; set; }

        // Byte size of the whole blob a caller must supply room for: the leading count plus
        // AllRowsScratch.
        internal int AllRowsLength { get; set; }

        internal bool AllRowsPending { get; set; }

        // The one caller-visible chunked read open on this workbook — an
        // xl_typed_reader or an xl_parse_arrow_stream — or null when
        // there is none. Set by NativeApi.TypedParseSession.Open, cleared by that
        // session's dispose or fault. See the class remarks for why one is the limit.
        internal NativeApi.TypedParseSession? LiveSession { get; set; }

        // Invalidates the live chunked read, if any, because cause is about to take
        // over the workbook's row cursor. The session latches the reason and reports it from every
        // later xl_typed_reader_next/get_next; it never silently resumes at whatever row
        // cause leaves the cursor on.
        //
        // cause: The ABI function taking over, named as the caller knows it.
        internal void FaultLiveSession(string cause)
        {
            LiveSession?.Fault($"this chunked read was invalidated by {cause} on the same workbook: a " +
                "workbook serves one row cursor at a time, so this read's position is no longer defined. " +
                "Finish or close the read before using the workbook for anything else.");
        }

        // Releases the live-session slot if session still owns it. Identity-checked
        // because a faulted session is detached immediately, so the slot may already belong to a newer one by
        // the time the faulted one is closed.
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
