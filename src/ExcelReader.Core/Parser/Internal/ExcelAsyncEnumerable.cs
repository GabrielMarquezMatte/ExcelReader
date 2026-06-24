using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser.Internal
{
    public sealed class ExcelAsyncEnumerable<T> : IAsyncEnumerable<T> where T : new()
    {
        private readonly XlsxReader _reader;
        private readonly ExcelParserConfig _config;
        private readonly CancellationToken _ct;

        internal ExcelAsyncEnumerable(XlsxReader reader, ExcelParserConfig config, CancellationToken ct = default)
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
                Justification = "XlsxReader is injected and not owned by this enumerator; caller manages its lifetime.")]
            private readonly XlsxReader _reader;
            private readonly int _headerRow;
            private readonly CancellationToken _ct;
            private RowProjector<T> _projector;
            private XlsxReader.Enumerator? _rows;
            private int _rowNumber;
            private T _current = default!;

            internal AsyncEnumerator(
                XlsxReader reader,
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

            // async only awaits — no ref struct locals here.
            // Ref struct access is delegated to ProcessCurrentRow (sync helper).
            private async ValueTask<bool> AdvanceAsync()
            {
                _rows ??= await _reader.GetAsyncEnumeratorAsync(_ct).ConfigureAwait(false);
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

            // Sync helper — ref struct locals are legal here.
            // Called synchronously between awaits; Row spans are valid for this call's duration.
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
