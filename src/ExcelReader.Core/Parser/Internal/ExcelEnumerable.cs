using System.Collections;
using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser.Internal
{
    /// <summary>Lazily projects XLSX rows into <typeparamref name="T"/> instances, for both synchronous and asynchronous enumeration.</summary>
    /// <typeparam name="T">The row model type to bind each row to.</typeparam>
    public sealed class ExcelEnumerable<T> : ExcelEnumerable<T, XlsxReader, XlsxReader.Enumerator>
    {
        [RequiresUnreferencedCode("Typed parsing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Typed parsing binds property setters at runtime (MethodInfo.CreateDelegate / MakeGenericMethod).")]
        internal ExcelEnumerable(XlsxReader reader, ExcelParserConfig config, CancellationToken ct = default)
            : base(reader, config, ct)
        {
        }

        internal ExcelEnumerable(XlsxReader reader, ExcelParserConfig config, TypeMapInfo<T> explicitInfo, CancellationToken ct = default)
            : base(reader, config, explicitInfo, ct)
        {
        }
    }

    /// <summary>Lazily projects rows read by a given reader/enumerator pair into <typeparamref name="T"/> instances, for both synchronous and asynchronous enumeration.</summary>
    /// <typeparam name="T">The row model type to bind each row to.</typeparam>
    /// <typeparam name="TReader">The concrete row reader type this instance pulls rows from.</typeparam>
    /// <typeparam name="TEnumerator">The concrete row enumerator type <typeparamref name="TReader"/> produces.</typeparam>
    [SuppressMessage("Design", "CA1034:Nested types should not be visible",
        Justification = "Public nested Enumerator/AsyncEnumerator are the standard foreach/await-foreach pattern.")]
    public class ExcelEnumerable<T, TReader, TEnumerator> : IEnumerable<T>, IAsyncEnumerable<T>
        where TReader : IExcelRowReader<TEnumerator>
        where TEnumerator : class, IExcelRowEnumerator
    {
        private readonly TReader _reader;
        private readonly ExcelParserConfig _config;
        private readonly CancellationToken _ct;
        // Resolved once here, in the constructor — never in GetEnumerator()/GetAsyncEnumerator() — so
        // that ExcelMappedParser<T>'s constructor overload below is the only path into this type whose
        // *method bodies* ever mention TypeMapper<T>. A trimmer/AOT analyzer decides reachability per
        // method, not per field: a `_info ?? TypeMapper<T>.GetInfo()` fallback living inside the shared
        // GetEnumerator() would make the reflection-based TypeMapper<T>.Build() reachable from every
        // caller, including ExcelMappedParser<T>'s, even though that field would always be non-null for
        // it at runtime — trimming can't see that far. Keeping the two constructors as the only two call
        // sites for the two info sources is what lets NativeAOT drop TypeMapper<T>'s reflection entirely
        // when only the explicit-info constructor is ever called from a published app's reachable code.
        private readonly TypeMapInfo<T> _info;

        [RequiresUnreferencedCode("Typed parsing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Typed parsing binds property setters at runtime (MethodInfo.CreateDelegate / MakeGenericMethod).")]
        internal ExcelEnumerable(TReader reader, ExcelParserConfig config, CancellationToken ct = default)
        {
            _reader = reader;
            _config = config;
            _info = TypeMapper<T>.GetInfo();
            _ct = ct;
        }

        internal ExcelEnumerable(TReader reader, ExcelParserConfig config, TypeMapInfo<T> explicitInfo, CancellationToken ct = default)
        {
            _reader = reader;
            _config = config;
            _info = explicitInfo;
            _ct = ct;
        }

        /// <inheritdoc cref="IEnumerable{T}.GetEnumerator"/>
        [SuppressMessage("Performance", "HLQ006:GetEnumerator should return a value type",
            Justification = "Enumerator is a class so the sync and async paths can share the SyncRowEnumerator/AsyncRowEnumerator base plumbing.")]
        [SuppressMessage("ApiDesign", "RS0041:Public members should not use oblivious types",
            Justification = "T is intentionally unconstrained so a row model can be a class or a struct (see RefParser/struct-binding support); constraining it would break that.")]
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
        [SuppressMessage("Performance", "HLQ006:GetAsyncEnumerator should return a value type",
            Justification = "Async enumerator requires a class to host the async state machine.")]
        [SuppressMessage("ApiDesign", "RS0041:Public members should not use oblivious types",
            Justification = "T is intentionally unconstrained so a row model can be a class or a struct (see RefParser/struct-binding support); constraining it would break that.")]
        public AsyncEnumerator GetAsyncEnumerator(CancellationToken cancellationToken = default)
        {
            CancellationToken effective = cancellationToken.CanBeCanceled ? cancellationToken : _ct;
            return new AsyncEnumerator(_reader, _info, _config.ColumnNameComparer, _config.HeaderNormalization, _config.HeaderRow, _config.Culture, _config.ThrowOnParseFailure, effective);
        }

        /// <summary>Enumerates rows synchronously, projecting each into a <typeparamref name="T"/> instance.</summary>
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

            private protected override ProjectionStep Project()
            {
                Row row = Rows.Current;
                return _projector.Advance(in row, ref CurrentValue);
            }
        }

        /// <summary>Enumerates rows asynchronously, projecting each into a <typeparamref name="T"/> instance.</summary>
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

            private protected override ProjectionStep Project()
            {
                Row row = Rows!.Current;
                return _projector.Advance(in row, ref CurrentValue);
            }
        }
    }
}
