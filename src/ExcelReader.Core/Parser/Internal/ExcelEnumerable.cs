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
        Justification = "Public nested Enumerator/AsyncEnumerator are the standard foreach/await-foreach pattern.")]
    public class ExcelEnumerable<T, TReader, TEnumerator> : IEnumerable<T>, IAsyncEnumerable<T>
        where TReader : IExcelRowReader<TEnumerator>
        where TEnumerator : class, IExcelRowEnumerator
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
        [SuppressMessage("Performance", "HLQ006:GetEnumerator should return a value type",
            Justification = "Enumerator is a class so the sync and async paths can share the SyncRowEnumerator/AsyncRowEnumerator base plumbing.")]
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

        [SuppressMessage("Performance", "HLQ006:GetAsyncEnumerator should return a value type",
            Justification = "Async enumerator requires a class to host the async state machine.")]
        public AsyncEnumerator GetAsyncEnumerator(CancellationToken cancellationToken = default)
        {
            TypeMapInfo<T> info = TypeMapper<T>.GetInfo();
            CancellationToken effective = cancellationToken.CanBeCanceled ? cancellationToken : _ct;
            return new AsyncEnumerator(_reader, info, _config.ColumnNameComparer, _config.HeaderNormalization, _config.HeaderRow, _config.Culture, effective);
        }

        public sealed class Enumerator : SyncRowEnumerator<T, TEnumerator>
        {
            private RowProjector<T> _projector;

            internal Enumerator(
                TEnumerator rows,
                TypeMapInfo<T> typeInfo,
                StringComparer comparer,
                HeaderNormalization normalization,
                int headerRow,
                bool isDate1904,
                IFormatProvider provider)
                : base(rows)
            {
                _projector = new RowProjector<T>(typeInfo, comparer, normalization, headerRow, isDate1904, provider);
            }

            private protected override ProjectionStep Project()
            {
                Row row = Rows.Current;
                return _projector.Advance(in row, ref CurrentValue);
            }
        }

        public sealed class AsyncEnumerator : AsyncRowEnumerator<T, TReader, TEnumerator>
        {
            private RowProjector<T> _projector;

            internal AsyncEnumerator(
                TReader reader,
                TypeMapInfo<T> typeInfo,
                StringComparer comparer,
                HeaderNormalization normalization,
                int headerRow,
                IFormatProvider provider,
                CancellationToken ct)
                : base(reader, ct)
            {
                _projector = new RowProjector<T>(typeInfo, comparer, normalization, headerRow, reader.IsDate1904, provider);
            }

            private protected override ProjectionStep Project()
            {
                Row row = Rows!.Current;
                return _projector.Advance(in row, ref CurrentValue);
            }
        }
    }
}
