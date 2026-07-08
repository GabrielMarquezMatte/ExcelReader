using System.Collections;
using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser.Internal
{
    public sealed class ExcelEnumerable<T> : ExcelEnumerable<T, XlsxReader, XlsxReader.Enumerator>
    {
        internal ExcelEnumerable(XlsxReader reader, ExcelParserConfig config, CancellationToken ct = default)
            : base(reader, config, ct)
        {
        }
    }

    [SuppressMessage("Design", "CA1034:Nested types should not be visible",
        Justification = "Public nested struct Enumerator is the standard foreach pattern.")]
    public class ExcelEnumerable<T, TReader, TEnumerator> : IEnumerable<T>, IAsyncEnumerable<T>
        where TReader : IExcelRowReader<TEnumerator>
        where TEnumerator : IExcelRowEnumerator
    {
        private readonly TReader _reader;
        private readonly ExcelParserConfig _config;
        private readonly CancellationToken _ct;

        internal ExcelEnumerable(TReader reader, ExcelParserConfig config, CancellationToken ct = default)
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
            TEnumerator rows = _reader.GetEnumerator();
            return new Enumerator(rows, info, _config.ColumnNameComparer, _config.HeaderNormalization, _config.HeaderRow, _reader.IsDate1904, _config.Culture);
        }

        IEnumerator<T> IEnumerable<T>.GetEnumerator()
        {
            return GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        IAsyncEnumerator<T> IAsyncEnumerable<T>.GetAsyncEnumerator(CancellationToken cancellationToken)
        {
            return GetAsyncEnumerator(cancellationToken);
        }

        public AsyncEnumerator GetAsyncEnumerator(CancellationToken cancellationToken = default)
        {
            TypeMapInfo<T> info = TypeMapper<T>.GetInfo();
            CancellationToken effective = cancellationToken.CanBeCanceled ? cancellationToken : _ct;
            return new AsyncEnumerator(_reader, info, _config.ColumnNameComparer, _config.HeaderNormalization, _config.HeaderRow, _config.Culture, effective);
        }

        public struct Enumerator : IEnumerator<T>
        {
            [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP006:Implement IDisposable",
                Justification = "Struct implements IDisposable; rows disposed in Dispose().")]
            private readonly TEnumerator _rows;
            private RowProjector<T> _projector;
            private T _current = default!;

            internal Enumerator(
                TEnumerator rows,
                TypeMapInfo<T> typeInfo,
                StringComparer comparer,
                HeaderNormalization normalization,
                int headerRow,
                bool isDate1904,
                IFormatProvider provider)
            {
                _rows = rows;
                _projector = new RowProjector<T>(typeInfo, comparer, normalization, headerRow, isDate1904, provider);
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
                    }
                }
                return false;
            }

            [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
                Justification = "The enumerator owns _rows; it was created by GetEnumerator, not injected from outside.")]
            public readonly void Dispose()
            {
                _rows.Dispose();
            }

            public void Reset()
            {
                throw new NotSupportedException();
            }
        }

        public sealed class AsyncEnumerator : IAsyncEnumerator<T>
        {
            private readonly TReader _reader;
            private readonly CancellationToken _ct;
            private RowProjector<T> _projector;
            private TEnumerator? _rows;
            private T _current = default!;

            internal AsyncEnumerator(
                TReader reader,
                TypeMapInfo<T> typeInfo,
                StringComparer comparer,
                HeaderNormalization normalization,
                int headerRow,
                IFormatProvider provider,
                CancellationToken ct)
            {
                _reader = reader;
                _ct = ct;
                _projector = new RowProjector<T>(typeInfo, comparer, normalization, headerRow, reader.IsDate1904, provider);
            }

            public T Current => _current;

            public ValueTask<bool> MoveNextAsync()
            {
                return AdvanceAsync();
            }

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
                    }
                }
                return false;
            }

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