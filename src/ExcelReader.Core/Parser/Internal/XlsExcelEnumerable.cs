using System.Collections;
using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser.Internal
{
    [SuppressMessage("Design", "CA1034:Nested types should not be visible",
        Justification = "Public nested struct Enumerator is the standard foreach pattern.")]
    public sealed class XlsExcelEnumerable<T> : IEnumerable<T>, IAsyncEnumerable<T> where T : new()
    {
        private readonly XlsReader _reader;
        private readonly ExcelParserConfig _config;
        private readonly CancellationToken _ct;

        internal XlsExcelEnumerable(XlsReader reader, ExcelParserConfig config, CancellationToken ct = default)
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
            XlsReader.Enumerator rows = _reader.GetEnumerator();
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

        public Enumerator GetAsyncEnumerator(CancellationToken cancellationToken = default)
        {
            TypeMapInfo<T> info = TypeMapper<T>.GetInfo();
            CancellationToken effective = cancellationToken.CanBeCanceled ? cancellationToken : _ct;
            XlsReader.Enumerator rows = _reader.GetAsyncEnumerator(effective);
            return new Enumerator(rows, info, _config.ColumnNameComparer, _config.HeaderRow, _reader.IsDate1904);
        }

        [SuppressMessage("Performance", "HLQ006:GetAsyncEnumerator should return a value type",
            Justification = "Struct enumerator is boxed once on return; caller sees IAsyncEnumerator<T> which communicates disposal.")]
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP015:Member should not return created and cached instance",
            Justification = "Each call creates a fresh enumerator; no caching.")]
        IAsyncEnumerator<T> IAsyncEnumerable<T>.GetAsyncEnumerator(CancellationToken cancellationToken = default)
        {
            return GetAsyncEnumerator(cancellationToken);
        }

        public struct Enumerator : IEnumerator<T>, IAsyncEnumerator<T>
        {
            private readonly XlsReader.Enumerator _rows;
            private RowProjector<T> _projector;
            private T _current = default!;

            internal Enumerator(
                XlsReader.Enumerator rows,
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

            public ValueTask<bool> MoveNextAsync()
            {
                return new ValueTask<bool>(MoveNext());
            }

            [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
                Justification = "The enumerator owns _rows — it was created by GetEnumerator, not injected from outside.")]
            public readonly void Dispose()
            {
                _rows.Dispose();
            }

            public readonly ValueTask DisposeAsync()
            {
                return _rows.DisposeAsync();
            }

            public readonly void Reset()
            {
                throw new NotSupportedException();
            }
        }
    }
}
