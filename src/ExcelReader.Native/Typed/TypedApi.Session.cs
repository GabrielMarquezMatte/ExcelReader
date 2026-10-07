using ExcelReader.Core.Reader;
using ExcelReader.Native.Reading;

namespace ExcelReader.Native.Typed
{
    internal static unsafe partial class TypedApi
    {
        internal sealed class TypedParseSession : IDisposable
        {
            private readonly NativeColumnSpec[] _specs;
            private readonly int[] _columnIndices;
            private readonly long _maxRows;
            private readonly bool _isDate1904;
            private IExcelRowEnumerator? _rows;
            private bool _faulted;
            private string? _faultMessage;
            private bool _disposed;

            private TypedParseSession(NativeColumnSpec[] specs, int[] columnIndices,
                long maxRows, bool isDate1904, IExcelRowEnumerator rows)
            {
                _specs = specs;
                _columnIndices = columnIndices;
                _maxRows = maxRows;
                _isDate1904 = isDate1904;
                _rows = rows;
            }

            internal static int Open(NativeHandle? handle, int sheet, NativeColumnSpec[] specs, int headerRow, long maxRows,
                out TypedParseSession? session)
            {
                session = null;
                if (handle is null)
                {
                    return NativeStatus.InvalidHandle;
                }
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

                NativeApi.ClearLastError();
                int status = ReadApi.ResolveSheet(handle, sheet, out IExcelSheet? resolved);
                if (status != NativeStatus.Ok)
                {
                    return status;
                }

                IExcelRowEnumerator? rows = null;
                try
                {
                    rows = resolved!.GetEnumerator();
                    int[] columnIndices = new int[specs.Length];
                    if (!TryResolveColumns(rows, specs, headerRow, columnIndices, out string? resolveError))
                    {
                        NativeApi.SetLastError(resolveError);
                        rows.Dispose();
                        return NativeStatus.InvalidArgument;
                    }

                    session = new TypedParseSession(specs, columnIndices, maxRows, handle.Workbook.IsDate1904, rows);
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

            private int Fail(string message)
            {
                _faulted = true;
                _faultMessage = message;
                ReleaseRows();
                NativeApi.SetLastError(message);
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
                ReleaseRows();
            }
        }
    }
}
