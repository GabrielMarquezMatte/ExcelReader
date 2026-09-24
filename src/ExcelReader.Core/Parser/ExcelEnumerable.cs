using System.Collections;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsx;

namespace ExcelReader.Core.Parser
{
    /// <summary>Lazily projects XLSX rows into <typeparamref name="T"/> instances, for both synchronous and asynchronous enumeration.</summary>
    /// <typeparam name="T">The row model type to bind each row to.</typeparam>
    public sealed class ExcelEnumerable<T> : ExcelEnumerable<T, XlsxReader, XlsxReader.Enumerator>
        where T : allows ref struct
    {
        internal ExcelEnumerable(XlsxReader reader, ExcelParserConfig config, TypeMapInfo<T> explicitInfo)
            : base(reader, config, explicitInfo)
        {
        }
    }

    /// <summary>Lazily projects rows read by a given reader/enumerator pair into <typeparamref name="T"/> instances, for both synchronous and asynchronous enumeration.</summary>
    /// <typeparam name="T">The row model type to bind each row to.</typeparam>
    /// <typeparam name="TReader">The concrete row reader type this instance pulls rows from.</typeparam>
    /// <typeparam name="TEnumerator">The concrete row enumerator type <typeparamref name="TReader"/> produces.</typeparam>
    [EditorBrowsable(EditorBrowsableState.Never)]
    [SuppressMessage("Design", "CA1034:Nested types should not be visible",
        Justification = "Public nested Enumerator/AsyncEnumerator are the standard foreach/await-foreach pattern.")]
    public class ExcelEnumerable<T, TReader, TEnumerator> : IEnumerable<T>, IAsyncEnumerable<T>
        where T : allows ref struct
        where TReader : IExcelRowReader<TEnumerator>
        where TEnumerator : class, IExcelRowEnumerator
    {
        private readonly TReader _reader;
        private readonly ExcelParserConfig _config;
        private readonly TypeMapInfo<T> _info;

        internal ExcelEnumerable(TReader reader, ExcelParserConfig config, TypeMapInfo<T> explicitInfo)
        {
            _reader = reader;
            _config = config;
            _info = explicitInfo;
        }

        /// <inheritdoc cref="IEnumerable{T}.GetEnumerator"/>
        [SuppressMessage("ApiDesign", "RS0041:Public members should not use oblivious types",
            Justification = "T allows ref struct so a row model can be a class, a struct or a ref struct; constraining it would break that.")]
        public Enumerator GetEnumerator()
        {
            TEnumerator rows = _reader.GetEnumerator();
            return new Enumerator(rows, _info, _config.ColumnNameComparer, _config.HeaderNormalization, _config.HeaderRow, _reader.IsDate1904, _config.Culture, _config.ThrowOnParseFailure);
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

        /// <inheritdoc cref="IAsyncEnumerable{T}.GetAsyncEnumerator"/>
        [SuppressMessage("ApiDesign", "RS0041:Public members should not use oblivious types",
            Justification = "T allows ref struct so a row model can be a class, a struct or a ref struct; constraining it would break that.")]
        public AsyncEnumerator GetAsyncEnumerator(CancellationToken cancellationToken = default)
        {
            return new AsyncEnumerator(_reader, _info, _config.ColumnNameComparer, _config.HeaderNormalization, _config.HeaderRow, _config.Culture, _config.ThrowOnParseFailure, cancellationToken);
        }

        /// <summary>Enumerates rows synchronously, projecting each into a <typeparamref name="T"/> instance.</summary>
        [EditorBrowsable(EditorBrowsableState.Never)]
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
                IFormatProvider provider,
                bool throwOnParseFailure = false)
                : base(rows)
            {
                _projector = new RowProjector<T>(typeInfo, comparer, normalization, headerRow, isDate1904, provider, throwOnParseFailure);
            }

            private protected override ProjectionStep Classify()
            {
                Row row = Rows.Current;
                return _projector.Classify(in row);
            }

            private protected override T Project()
            {
                return _projector.Project(Rows.Current);
            }
        }

        /// <summary>Enumerates rows asynchronously, projecting each into a <typeparamref name="T"/> instance.</summary>
        [EditorBrowsable(EditorBrowsableState.Never)]
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
                bool throwOnParseFailure,
                CancellationToken ct)
                : base(reader, ct)
            {
                _projector = new RowProjector<T>(typeInfo, comparer, normalization, headerRow, reader.IsDate1904, provider, throwOnParseFailure);
            }

            private protected override ProjectionStep Classify()
            {
                Row row = Rows!.Current;
                return _projector.Classify(in row);
            }

            private protected override T Project()
            {
                return _projector.Project(Rows!.Current);
            }
        }
    }
}
