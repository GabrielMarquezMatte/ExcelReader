using ExcelReader.Core.Reader;

namespace ExcelReader.Native
{
    internal static unsafe partial class NativeApi
    {
        /// <summary>
        /// A resumable <see cref="ParseTyped"/>: the same schema-driven columnar read, cut into
        /// batches of at most <c>maxRows</c> rows. Column resolution happens once, at open — the
        /// header row is consumed there and never re-read — so a batch is purely the row loop.
        /// </summary>
        /// <remarks>
        /// Fresh <see cref="ColumnBuilder"/>s per batch are the entire memory ceiling: each
        /// builder's <see cref="ChunkedBuffer{T}"/> chain is dropped once <see cref="BuildTable"/>
        /// has copied it out, so the working set is one batch's columns plus one batch's output
        /// block rather than the whole sheet.
        ///
        /// Holds a strong reference to the <see cref="NativeHandle"/> so the reader cannot be
        /// collected underneath it. This is also the only thing in this ABI that keeps an
        /// <see cref="IExcelRowEnumerator"/> open ACROSS calls, which a workbook cannot serve twice at
        /// once — so a caller-visible session takes <see cref="NativeHandle.LiveSession"/> for its whole
        /// lifetime (a second one is refused) and any other read on that workbook, including
        /// <c>xl_close</c> and <c>xl_move_to_sheet</c>, <see cref="Fault"/>s it instead of rewinding the
        /// reader underneath it. A faulted session reports its latched message from every later
        /// <see cref="NextBatch"/> and never returns rows again.
        /// </remarks>
        internal sealed class TypedParseSession : IDisposable
        {
            private readonly NativeHandle _handle;
            private readonly NativeColumnSpec[] _specs;
            private readonly int[] _columnIndices;
            private readonly long _maxRows;
            private readonly bool _isDate1904;
            private IExcelRowEnumerator? _rows;
            // Latched together: _faultMessage is the reason for the FIRST fault and is re-reported by
            // every later call, so a consumer that only reads the error after the second one still
            // learns what actually happened.
            private bool _faulted;
            private string? _faultMessage;
            private bool _disposed;

            private TypedParseSession(NativeHandle handle, NativeColumnSpec[] specs, int[] columnIndices,
                long maxRows, bool isDate1904, IExcelRowEnumerator rows)
            {
                _handle = handle;
                _specs = specs;
                _columnIndices = columnIndices;
                _maxRows = maxRows;
                _isDate1904 = isDate1904;
                _rows = rows;
            }

            /// <summary>
            /// Opens a CALLER-VISIBLE session: resolves the columns once, positions the reader at the
            /// first data row, and takes the workbook's single live-session slot for as long as the
            /// session lives.
            /// </summary>
            /// <param name="maxRows">Rows per batch. 0 means unbounded — one batch holding every row.</param>
            /// <remarks>
            /// Refused with <see cref="NativeStatus.Error"/> when the workbook already has one open —
            /// a second <c>xl_typed_reader_open</c>, an <c>xl_parse_arrow_stream</c> while a typed
            /// reader is open, or the reverse. No new status code for it: it is a workbook-state
            /// failure whose detail belongs in <c>xl_last_error</c>, exactly what XL_ERROR means
            /// everywhere else in this ABI (the arguments are all perfectly valid, so
            /// XL_INVALID_ARGUMENT would misdescribe it, and the handle is live, so XL_INVALID_HANDLE
            /// would too).
            /// </remarks>
            internal static int Open(NativeHandle? handle, NativeColumnSpec[] specs, int headerRow, long maxRows,
                out TypedParseSession? session)
            {
                session = null;
                if (handle is null)
                {
                    return NativeStatus.InvalidHandle;
                }
                if (handle.LiveSession is not null)
                {
                    SetLastError("this workbook already has a chunked read open; close it "
                        + "(xl_typed_reader_close, or the Arrow stream's own release) before opening another.");
                    return NativeStatus.Error;
                }

                int status = OpenCore(handle, specs, headerRow, maxRows, faultCause: null, out session);
                if (status == NativeStatus.Ok)
                {
                    handle.LiveSession = session;
                }
                return status;
            }

            /// <summary>
            /// Opens the short-lived session behind <see cref="ParseTyped"/>/<see cref="ParseArrow"/>:
            /// one unbounded batch, drained and disposed inside the same ABI call.
            /// </summary>
            /// <param name="cause">The ABI function opening it, for the fault message it leaves behind.</param>
            /// <remarks>
            /// Deliberately does NOT touch <see cref="NativeHandle.LiveSession"/>. That is what lets the
            /// whole-sheet functions coexist with the one-live-session rule instead of being refused by
            /// it: they cannot be interleaved with anything, because they hold the workbook's cursor
            /// only for the duration of their own call — and for the same reason they cannot fault
            /// themselves, since the session they fault is by definition someone else's. What they do
            /// have to do is fault any caller-visible session before taking the reader's enumerator —
            /// but not before <paramref name="specs"/>/<paramref name="headerRow"/> are known to be
            /// valid, or a call that was always going to fail XL_INVALID_ARGUMENT on pure argument
            /// checking would destroy an unrelated caller's open reader without ever touching the
            /// cursor. See <see cref="OpenCore"/>'s own ordering for where the line sits.
            /// </remarks>
            internal static int OpenTransient(NativeHandle? handle, NativeColumnSpec[] specs, int headerRow,
                string cause, out TypedParseSession? session)
            {
                session = null;
                if (handle is null)
                {
                    return NativeStatus.InvalidHandle;
                }
                return OpenCore(handle, specs, headerRow, maxRows: 0, faultCause: cause, out session);
            }

            /// <param name="faultCause">Non-null only for <see cref="OpenTransient"/>: the ABI function
            /// name to fault any live session with, once argument validation has passed and this call is
            /// actually about to take the reader's enumerator. Null for <see cref="Open"/>, which never
            /// reaches here with a live session already set (its own caller refuses that first).</param>
            private static int OpenCore(NativeHandle handle, NativeColumnSpec[] specs, int headerRow, long maxRows,
                string? faultCause, out TypedParseSession? session)
            {
                session = null;
                if (maxRows < 0)
                {
                    SetLastError($"max_rows must be 0 (unbounded) or positive; got {maxRows}.");
                    return NativeStatus.InvalidArgument;
                }
                if (!TryValidateArguments(specs, headerRow, out string? argumentError))
                {
                    SetLastError(argumentError);
                    return NativeStatus.InvalidArgument;
                }

                // Only past this point does the call actually commit to taking over the workbook's row
                // cursor, so only from here does an unrelated caller's live session need to be faulted.
                if (faultCause is not null)
                {
                    handle.FaultLiveSession(faultCause);
                }

                ClearLastError();
                IExcelRowEnumerator? rows = null;
                try
                {
                    rows = handle.Reader.GetEnumerator();
                    int[] columnIndices = new int[specs.Length];
                    if (!TryResolveColumns(rows, specs, headerRow, columnIndices, out string? resolveError))
                    {
                        SetLastError(resolveError);
                        rows.Dispose();
                        return NativeStatus.InvalidArgument;
                    }

                    session = new TypedParseSession(handle, specs, columnIndices, maxRows,
                        handle.Reader.IsDate1904, rows);
                    return NativeStatus.Ok;
                }
                catch (Exception exception)
                {
                    SetLastError(exception.Message);
                    rows?.Dispose();
                    return NativeStatus.Error;
                }
            }

            /// <summary>
            /// Fills <paramref name="table"/> with the next batch. Returns <see cref="NativeStatus.Eof"/>
            /// once the sheet is exhausted, zeroing <paramref name="table"/> so the caller's
            /// <see cref="FreeTable"/> stays safe either way.
            /// </summary>
            internal int NextBatch(out NativeTable table)
            {
                table = default;
                if (_disposed)
                {
                    SetLastError("this typed reader has been closed.");
                    return NativeStatus.InvalidHandle;
                }
                // A conversion failure — or any other loss of a defined resume point, including an
                // interleaved call on the same workbook — latches: every later call reports the same
                // thing rather than silently continuing past bad data or from a rewound cursor.
                if (_faulted)
                {
                    SetLastError(_faultMessage!);
                    return NativeStatus.Error;
                }
                if (_rows is null)
                {
                    return NativeStatus.Eof;
                }

                ClearLastError();
                try
                {
                    ColumnBuilder[] builders = new ColumnBuilder[_specs.Length];
                    for (int i = 0; i < _specs.Length; i++)
                    {
                        builders[i] = new ColumnBuilder(_specs[i].Type, _specs[i].Nullable);
                    }

                    long taken = 0;
                    while ((_maxRows == 0 || taken < _maxRows) && _rows.MoveNext())
                    {
                        if (!TryAppendRow(builders, _rows.Current, _columnIndices, _isDate1904, out int failedColumn))
                        {
                            return Fail(DescribeFailedColumn(_specs, failedColumn));
                        }
                        taken++;
                    }

                    if (taken == 0)
                    {
                        return NativeStatus.Eof;
                    }
                    table = BuildTable(builders);
                    return NativeStatus.Ok;
                }
                catch (Exception exception)
                {
                    table = default;
                    return Fail(exception.Message);
                }
            }

            /// <summary>The reason this session is faulted, or <see langword="null"/> while it is healthy.
            /// Owned here rather than read back out of the thread's <c>xl_last_error</c>, which any
            /// unrelated ExcelReader call on the same thread can overwrite between the fault and the
            /// report — see <see cref="ArrowStreamGetNextCore"/>, the one consumer that needs it.</summary>
            internal string? FaultMessage
            {
                get
                {
                    return _faultMessage;
                }
            }

            /// <summary>
            /// Permanently invalidates this session with <paramref name="message"/>, releasing its row
            /// cursor and the workbook's live-session slot. Idempotent, and the FIRST message wins: the
            /// original cause is more useful than whatever happened afterwards.
            /// </summary>
            internal void Fault(string message)
            {
                if (_disposed || _faulted)
                {
                    return;
                }
                // State first, cursor second: a throwing IExcelRowEnumerator.Dispose (plausibly an
                // IOException) must still leave this session latched and detached, not half-faulted.
                _faulted = true;
                _faultMessage = message;
                _handle.ReleaseLiveSession(this);
                ReleaseRows();
            }

            // Latches `message` and reports it now, for a fault raised inside NextBatch itself.
            // Re-entrant on purpose: Fault's own cursor release can throw, which lands in NextBatch's
            // catch and comes straight back here — the already-latched message then wins and the
            // second release is a no-op.
            private int Fail(string message)
            {
                Fault(message);
                SetLastError(_faultMessage!);
                return NativeStatus.Error;
            }

            // Nulled before the dispose, so a throwing dispose cannot leave a cursor this session
            // still believes it can read from.
            private void ReleaseRows()
            {
                IExcelRowEnumerator? rows = _rows;
                _rows = null;
                rows?.Dispose();
            }

            public void Dispose()
            {
                if (_disposed)
                {
                    return;
                }
                _disposed = true;
                _handle.ReleaseLiveSession(this);
                ReleaseRows();
                // _handle is deliberately NOT disposed: the session borrows the workbook, it does
                // not own it. xl_close remains the only thing that closes a workbook.
                GC.KeepAlive(_handle);
            }
        }
    }
}
