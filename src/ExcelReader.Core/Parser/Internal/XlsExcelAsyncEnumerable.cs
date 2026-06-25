using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser.Internal
{
    public sealed class XlsExcelAsyncEnumerable<T> : IAsyncEnumerable<T> where T : new()
    {
        private readonly XlsReader _reader;
        private readonly ExcelParserConfig _config;
        private readonly CancellationToken _ct;

        internal XlsExcelAsyncEnumerable(XlsReader reader, ExcelParserConfig config, CancellationToken ct = default)
        {
            _reader = reader;
            _config = config;
            _ct = ct;
        }

        [SuppressMessage("Performance", "HLQ006:GetAsyncEnumerator should return a value type",
            Justification = "Async enumerator requires a class to host the async state machine.")]
        public IAsyncEnumerator<T> GetAsyncEnumerator(CancellationToken cancellationToken = default)
        {
            TypeMapInfo<T> info = TypeMapper<T>.GetInfo();
            CancellationToken effective = cancellationToken.CanBeCanceled ? cancellationToken : _ct;
            return new AsyncEnumerator(_reader, info, _config.ColumnNameComparer, _config.HeaderRow, effective);
        }

        private sealed class AsyncEnumerator : IAsyncEnumerator<T>
        {
            [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
                Justification = "XlsReader is injected and not owned by this enumerator; caller manages its lifetime.")]
            private readonly XlsReader _reader;
            private readonly int _headerRow;
            private readonly CancellationToken _ct;
            private RowProjector<T> _projector;
            private XlsReader.Enumerator? _rows;
            private int _rowNumber;
            private T _current = default!;

            internal AsyncEnumerator(
                XlsReader reader,
                TypeMapInfo<T> typeInfo,
                StringComparer comparer,
                int headerRow,
                CancellationToken ct)
            {
                _reader = reader;
                _headerRow = headerRow;
                _ct = ct;
                _projector = new RowProjector<T>(typeInfo, comparer, reader.IsDate1904);
            }

            public T Current => _current;

            public ValueTask<bool> MoveNextAsync()
            {
                return AdvanceAsync();
            }

            private async ValueTask<bool> AdvanceAsync()
            {
                _rows ??= _reader.GetAsyncEnumerator(_ct);
                while (true)
                {
                    bool hasMore = await _rows.MoveNextAsync().ConfigureAwait(false);
                    if (!hasMore)
                    {
                        return false;
                    }
                    _rowNumber++;
                    if (ProcessCurrentRow())
                    {
                        return true;
                    }
                }
            }

            private bool ProcessCurrentRow()
            {
                Row row = _rows!.Current;
                if (_rowNumber < _headerRow)
                {
                    return false;
                }
                if (_rowNumber == _headerRow)
                {
                    _projector.BuildColumnMap(in row);
                    return false;
                }
                if (!_projector.IsMapped)
                {
                    return false;
                }
                _current = new T();
                _projector.ParseCurrentRow(in row, ref _current);
                return true;
            }

            public ValueTask DisposeAsync()
            {
                if (_rows is null)
                {
                    return ValueTask.CompletedTask;
                }
                return _rows.DisposeAsync();
            }
        }
    }
}
