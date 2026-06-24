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
            private readonly TypeMapInfo<T> _typeInfo;
            private readonly StringComparer _comparer;
            private readonly int _headerRow;
            private readonly CancellationToken _ct;
            private XlsxReader.Enumerator? _rows;
            private int _rowNumber;
            private ColumnParser<T>?[]? _columnParsers;
            private int _maxColumn;
            private T _current = default!;

            internal AsyncEnumerator(
                XlsxReader reader,
                TypeMapInfo<T> typeInfo,
                StringComparer comparer,
                int headerRow,
                CancellationToken ct)
            {
                _reader = reader;
                _typeInfo = typeInfo;
                _comparer = comparer;
                _headerRow = headerRow;
                _ct = ct;
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
                    BuildColumnMap(in row);
                    return false;
                }
                if (_columnParsers is null)
                {
                    return false;
                }
                _current = new T();
                ParseCurrentRow(in row);
                return true;
            }

            private void BuildColumnMap(in Row row)
            {
                _maxColumn = row.ColumnCount;
                _columnParsers = new ColumnParser<T>?[_maxColumn];
                for (int col = 0; col < _maxColumn; col++)
                {
                    Cell cell = row[col];
                    string header = cell.GetString();
                    if (string.IsNullOrEmpty(header))
                    {
                        continue;
                    }
                    int index = _typeInfo.FindIndex(header, _comparer);
                    if (index >= 0)
                    {
                        _columnParsers[col] = _typeInfo.Parsers[index];
                    }
                }
            }

            private void ParseCurrentRow(in Row row)
            {
                bool isDate1904 = _reader.IsDate1904;
                int bound = Math.Min(_maxColumn, row.ColumnCount);
                for (int col = 0; col < bound; col++)
                {
                    ColumnParser<T>? parser = _columnParsers![col];
                    if (parser is null)
                    {
                        continue;
                    }
                    parser(ref _current, in row, col, isDate1904);
                }
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
