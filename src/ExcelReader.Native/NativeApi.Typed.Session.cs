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
        /// collected underneath it. A caller that closes the workbook first still gets a clean
        /// error, not undefined behavior: the disposed reader throws and
        /// <see cref="NextBatch"/> latches it as a fault.
        /// </remarks>
        internal sealed class TypedParseSession : IDisposable
        {
            private readonly NativeHandle _handle;
            private readonly NativeColumnSpec[] _specs;
            private readonly int[] _columnIndices;
            private readonly long _maxRows;
            private readonly bool _isDate1904;
            private IExcelRowEnumerator? _rows;
            private bool _faulted;
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
            /// Resolves the columns once and positions the reader at the first data row, leaving the
            /// row loop itself to <see cref="NextBatch"/>.
            /// </summary>
            /// <param name="maxRows">Rows per batch. 0 means unbounded — one batch holding every row.</param>
            internal static int Open(NativeHandle? handle, NativeColumnSpec[] specs, int headerRow, long maxRows,
                out TypedParseSession? session)
            {
                session = null;
                if (handle is null)
                {
                    return NativeStatus.InvalidHandle;
                }
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
                // A conversion failure leaves the enumerator mid-sheet with no defined resume point,
                // so the fault latches: every later call reports the same thing rather than silently
                // continuing past bad data.
                if (_faulted)
                {
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
                            _faulted = true;
                            SetLastError(DescribeFailedColumn(_specs, failedColumn));
                            return NativeStatus.Error;
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
                    _faulted = true;
                    SetLastError(exception.Message);
                    table = default;
                    return NativeStatus.Error;
                }
            }

            public void Dispose()
            {
                if (_disposed)
                {
                    return;
                }
                _disposed = true;
                _rows?.Dispose();
                _rows = null;
                // _handle is deliberately NOT disposed: the session borrows the workbook, it does
                // not own it. xl_close remains the only thing that closes a workbook.
                GC.KeepAlive(_handle);
            }
        }
    }
}
