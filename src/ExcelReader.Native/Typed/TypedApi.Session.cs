using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Columnar;
using ExcelReader.Native.Reading;

namespace ExcelReader.Native.Typed
{
    internal static unsafe partial class TypedApi
    {
        internal sealed class TypedParseSession : IDisposable
        {
            private readonly TypedColumnSpec[] _specs;
            private readonly int[] _columnIndices;
            private readonly long _maxRows;
            private readonly bool _isDate1904;
            private IExcelRowEnumerator? _rows;
            private bool _faulted;
            private string? _faultMessage;
            private bool _disposed;

            private TypedParseSession(TypedColumnSpec[] specs, int[] columnIndices,
                long maxRows, bool isDate1904, IExcelRowEnumerator rows)
            {
                _specs = specs;
                _columnIndices = columnIndices;
                _maxRows = maxRows;
                _isDate1904 = isDate1904;
                _rows = rows;
            }

            internal static int Open(NativeHandle? handle, int sheet, NativeColumnSpec[] specs, int headerRow, long maxRows,
                out TypedParseSession? session, bool allowDictionary = false)
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
                int status = ValidateSpecs(specs, headerRow, allowDictionary, out TypedColumnSpec[] coreSpecs);
                if (status != NativeStatus.Ok)
                {
                    return status;
                }
                status = ReadApi.ResolveSheet(handle, sheet, out IExcelSheet? resolved);
                if (status != NativeStatus.Ok)
                {
                    return status;
                }
                return OpenRows(resolved!, coreSpecs, headerRow, maxRows, handle.Workbook.IsDate1904, out session);
            }

            private static int OpenRows(IExcelSheet sheet, TypedColumnSpec[] specs, int headerRow, long maxRows, bool isDate1904,
                out TypedParseSession? session)
            {
                session = null;
                IExcelRowEnumerator? rows = null;
                try
                {
                    rows = sheet.GetEnumerator();
                    int[] columnIndices = new int[specs.Length];
                    if (!TypedTableParser.TryResolveColumns(rows, specs, headerRow, columnIndices, out string? resolveError))
                    {
                        NativeApi.SetLastError(resolveError);
                        rows.Dispose();
                        return NativeStatus.InvalidArgument;
                    }

                    session = new TypedParseSession(specs, columnIndices, maxRows, isDate1904, rows);
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
                int status = NextColumns(out ColumnarTable? columns);
                if (status != NativeStatus.Ok)
                {
                    return status;
                }
                using ColumnarTable built = columns!;
                try
                {
                    table = ToNativeTable(built);
                    return NativeStatus.Ok;
                }
                catch (Exception exception)
                {
                    return Fail(exception.Message);
                }
            }

            internal int NextColumns(out ColumnarTable? table)
            {
                table = null;
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
                using TypedTableBuilder builder = new(_specs);
                try
                {
                    while ((_maxRows == 0 || builder.RowCount < _maxRows) && _rows.MoveNext())
                    {
                        if (!builder.TryAppend(_rows.Current, _columnIndices, _isDate1904, out int failedColumn))
                        {
                            return Fail(TypedTableParser.DescribeFailedColumn(_specs, failedColumn));
                        }
                    }

                    if (builder.RowCount == 0)
                    {
                        return NativeStatus.Eof;
                    }
                    table = builder.Build();
                    return NativeStatus.Ok;
                }
                catch (Exception exception)
                {
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
