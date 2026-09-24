using ExcelReader.Core.Reader;

namespace ExcelReader.Native.Typed
{
    internal static unsafe partial class TypedApi
    {
        internal sealed class TypedParseSession : IDisposable
        {
            private readonly NativeHandle _handle;
            private readonly NativeColumnSpec[] _specs;
            private readonly int[] _columnIndices;
            private readonly long _maxRows;
            private readonly bool _isDate1904;
            private IExcelRowEnumerator? _rows;
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
                    NativeApi.SetLastError("this workbook already has a chunked read open; close it "
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

            private static int OpenCore(NativeHandle handle, NativeColumnSpec[] specs, int headerRow, long maxRows,
                string? faultCause, out TypedParseSession? session)
            {
                session = null;
                if (maxRows < 0)
                {
                    NativeApi.SetLastError($"max_rows must be 0 (unbounded) or positive; got {maxRows}.");
                    return NativeStatus.InvalidArgument;
                }
                if (!TryValidateArguments(specs, headerRow, out string? argumentError))
                {
                    NativeApi.SetLastError(argumentError);
                    return NativeStatus.InvalidArgument;
                }

                if (faultCause is not null)
                {
                    handle.FaultLiveSession(faultCause);
                }

                NativeApi.ClearLastError();
                IExcelRowEnumerator? rows = null;
                try
                {
                    rows = handle.Reader.GetEnumerator();
                    int[] columnIndices = new int[specs.Length];
                    if (!TryResolveColumns(rows, specs, headerRow, columnIndices, out string? resolveError))
                    {
                        NativeApi.SetLastError(resolveError);
                        rows.Dispose();
                        return NativeStatus.InvalidArgument;
                    }

                    session = new TypedParseSession(handle, specs, columnIndices, maxRows,
                        handle.Reader.IsDate1904, rows);
                    return NativeStatus.Ok;
                }
                catch (Exception exception)
                {
                    NativeApi.SetLastError(exception.Message);
                    rows?.Dispose();
                    return NativeStatus.Error;
                }
            }

            internal int NextBatch(out NativeTable table)
            {
                table = default;
                if (_disposed)
                {
                    NativeApi.SetLastError("this typed reader has been closed.");
                    return NativeStatus.InvalidHandle;
                }
                if (_faulted)
                {
                    NativeApi.SetLastError(_faultMessage!);
                    return NativeStatus.Error;
                }
                if (_rows is null)
                {
                    return NativeStatus.Eof;
                }

                NativeApi.ClearLastError();
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

            internal string? FaultMessage
            {
                get
                {
                    return _faultMessage;
                }
            }

            internal void Fault(string message)
            {
                if (_disposed || _faulted)
                {
                    return;
                }
                _faulted = true;
                _faultMessage = message;
                _handle.ReleaseLiveSession(this);
                ReleaseRows();
            }

            private int Fail(string message)
            {
                Fault(message);
                NativeApi.SetLastError(_faultMessage!);
                return NativeStatus.Error;
            }

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
                GC.KeepAlive(_handle);
            }
        }
    }
}
