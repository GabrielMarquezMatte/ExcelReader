using System.Collections;
using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser.Internal
{
    [SuppressMessage("Design", "CA1034:Nested types should not be visible",
        Justification = "Public nested struct Enumerator is the standard foreach pattern.")]
    public sealed class ExcelEnumerable<T> : IEnumerable<T>, IAsyncEnumerable<T> where T : new()
    {
        private readonly XlsxReader _reader;
        private readonly ExcelParserConfig _config;
        private readonly CancellationToken _ct;

        internal ExcelEnumerable(XlsxReader reader, ExcelParserConfig config, CancellationToken ct = default)
        {
            _reader = reader;
            _config = config;
            _ct = ct;
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP015:Member should not return created and cached instance",
            Justification = "Each call creates a fresh enumerator; no caching.")]
        public Enumerator GetEnumerator()
        {
            TypeMapInfo<T> info = TypeMapper<T>.GetInfo();
            XlsxReader.Enumerator rows = _reader.GetEnumerator();
            return new Enumerator(rows, info, _config.ColumnNameComparer, _config.HeaderRow, _reader.IsDate1904);
        }

        IEnumerator<T> IEnumerable<T>.GetEnumerator()
        {
            return GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        // Unlike the .xls path, the xlsx row enumerator awaits real stream I/O, so the async enumerator
        // must be a class hosting the state machine — it cannot share the sync struct Enumerator.
        [SuppressMessage("Performance", "HLQ006:GetAsyncEnumerator should return a value type",
            Justification = "Async enumerator requires a class to host the async state machine.")]
        public IAsyncEnumerator<T> GetAsyncEnumerator(CancellationToken cancellationToken = default)
        {
            TypeMapInfo<T> info = TypeMapper<T>.GetInfo();
            CancellationToken effective = cancellationToken.CanBeCanceled ? cancellationToken : _ct;
            return new AsyncEnumerator(_reader, info, _config.ColumnNameComparer, _config.HeaderRow, effective);
        }

        public struct Enumerator : IEnumerator<T>
        {
            [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP006:Implement IDisposable",
                Justification = "Struct implements IDisposable; rows disposed in Dispose().")]
            private readonly XlsxReader.Enumerator _rows;
            private RowProjector<T> _projector;
            private T _current = default!;

            internal Enumerator(
                XlsxReader.Enumerator rows,
                TypeMapInfo<T> typeInfo,
                StringComparer comparer,
                int headerRow,
                bool isDate1904)
            {
                _rows = rows;
                _projector = new RowProjector<T>(typeInfo, comparer, headerRow, isDate1904);
            }

            public readonly T Current => _current;

            readonly object? IEnumerator.Current => _current;

            public bool MoveNext()
            {
                while (_rows.MoveNext())
                {
                    Row row = _rows.Current;
                    switch (_projector.Advance(in row, ref _current))
                    {
                        case ProjectionStep.Yield:
                            return true;
                        case ProjectionStep.Stop:
                            return false;
                        case ProjectionStep.Skip:
                            break;
                    }
                }
                return false;
            }

            [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
                Justification = "The enumerator owns _rows — it was created by GetEnumerator, not injected from outside.")]
            public readonly void Dispose()
            {
                _rows.Dispose();
            }

            public void Reset()
            {
                throw new NotSupportedException();
            }
        }

        private sealed class AsyncEnumerator : IAsyncEnumerator<T>
        {
            [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
                Justification = "XlsxReader is injected and not owned by this enumerator; caller manages its lifetime.")]
            private readonly XlsxReader _reader;
            private readonly CancellationToken _ct;
            private RowProjector<T> _projector;
            private XlsxReader.Enumerator? _rows;
            private T _current = default!;

            internal AsyncEnumerator(
                XlsxReader reader,
                TypeMapInfo<T> typeInfo,
                StringComparer comparer,
                int headerRow,
                CancellationToken ct)
            {
                _reader = reader;
                _ct = ct;
                _projector = new RowProjector<T>(typeInfo, comparer, headerRow, reader.IsDate1904);
            }

            public T Current => _current;

            public ValueTask<bool> MoveNextAsync()
            {
                return AdvanceAsync();
            }

            // async only awaits — no ref struct locals here.
            // Ref struct access is delegated to Project (sync helper).
            private async ValueTask<bool> AdvanceAsync()
            {
                _rows ??= await _reader.GetAsyncEnumeratorAsync(_ct).ConfigureAwait(false);
                while (await _rows.MoveNextAsync().ConfigureAwait(false))
                {
                    switch (Project())
                    {
                        case ProjectionStep.Yield:
                            return true;
                        case ProjectionStep.Stop:
                            return false;
                        case ProjectionStep.Skip:
                            break;
                    }
                }
                return false;
            }

            // Sync helper — ref struct locals are legal here.
            // Called synchronously between awaits; Row spans are valid for this call's duration.
            private ProjectionStep Project()
            {
                Row row = _rows!.Current;
                return _projector.Advance(in row, ref _current);
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
