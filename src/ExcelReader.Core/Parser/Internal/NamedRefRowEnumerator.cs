#if NET9_0_OR_GREATER
using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser.Internal
{
    // Reflection-based, attribute-driven ("[ExcelColumn]"/"[ExcelRequired]"/"[ExcelConverter]") ref
    // struct row projector. Reuses TypeMapper<TModel>/TypeMapInfo<TModel>/ColumnParserFactory (widened
    // to `allows ref struct` for this TFM) for reflection + Expression-tree byref setter compilation —
    // the exact same machinery ExcelParser<T> already uses for classes/structs, now extended to ref
    // structs too. NOT AOT/trim-safe (unlike IExcelRowModel<TSelf>'s manual FromRow path).
    //
    // Can't reuse RowProjector<T> directly: RowProjector<T>.Advance classifies AND parses a row in one
    // call, immediately writing into a caller-supplied `ref T model`. Every existing caller stores
    // that result in a class FIELD (SyncRowEnumerator<T,TRows>.CurrentValue) — illegal for a
    // ref-struct-constrained T (CS8345: a ref struct field is only legal inside another ref struct).
    // So this type splits the same two responsibilities instead: MoveNext()/MoveNextAsync() only
    // advance/classify (skip pre-header rows, build the column map once at the header row), and Current's
    // getter parses the row into a fresh LOCAL model on every access — never stored as a field, safe to
    // call repeatedly (idempotent) for the same row.
    //
    // A reference type (not a struct): MoveNextAsync's awaiting slow path mutates enumeration state
    // (_rowNumber, and the one-time column map) across an await, which a struct would silently lose to the
    // async state machine's by-value `this` copy. A class shares the one instance, exactly like
    // AsyncRowEnumerator<T,TReader,TRows>. The lazy-Current design above is still required regardless:
    // a ref-struct TModel can never be stored in a field (CS8345), class or struct.
    /// <summary>A forward-only, zero-model-allocation cursor over rows bound by attribute to a ref struct model type, driven by <c>MoveNext</c>/<c>MoveNextAsync</c> and read via <see cref="Current"/>.</summary>
    /// <typeparam name="TModel">The ref-struct-capable row model type to bind each row to.</typeparam>
    /// <typeparam name="TEnumerator">The concrete row enumerator type this instance pulls rows from.</typeparam>
    public sealed class NamedRefRowEnumerator<TModel, TEnumerator> : IDisposable, IAsyncDisposable
        where TModel : allows ref struct
        where TEnumerator : class, IExcelRowEnumerator
    {
        [SuppressMessage("Performance", "HLQ011:ReadOnlyEnumeratorField",
            Justification = "TEnumerator is constrained to `class` here, so it is always a reference type — no copy-on-mutate risk from a readonly field.")]
        private readonly TEnumerator _rows;
        private readonly ExcelRowContext _context;
        private readonly TypeMapInfo<TModel> _typeInfo;
        private readonly StringComparer _comparer;
        private readonly HeaderNormalization _normalization;
        private readonly int _headerRow;
        private readonly bool _throwOnParseFailure;
        private ColumnBinding<TModel>[]? _bindings;
        private bool[] _seen;
        private int _requireValueCount;
        private int _rowNumber;

        internal NamedRefRowEnumerator(
            TEnumerator rows,
            ExcelRowContext context,
            TypeMapInfo<TModel> typeInfo,
            StringComparer comparer,
            HeaderNormalization normalization,
            int headerRow,
            bool throwOnParseFailure = false)
        {
            _rows = rows;
            _context = context;
            _typeInfo = typeInfo;
            _comparer = comparer;
            _normalization = normalization;
            _headerRow = headerRow;
            _throwOnParseFailure = throwOnParseFailure;
            _seen = [];
        }

        /// <summary>Gets the row at the enumerator's current position, freshly parsed into a new model instance on every access.</summary>
        // Recomputed on every access (see class remarks) — never cached in a field.
        public TModel Current
        {
            get
            {
                TModel model = default!;
                ParseCurrentRow(_rows.Current, ref model);
                return model;
            }
        }

        /// <inheritdoc cref="System.Collections.IEnumerator.MoveNext"/>
        public bool MoveNext()
        {
            while (_rows.MoveNext())
            {
                ProjectionStep step = ProjectionRules.ClassifyRow(ref _rowNumber, _headerRow, _bindings is not null);
                switch (step)
                {
                    case ProjectionStep.Yield:
                        return true;
                    case ProjectionStep.BuildMap:
                        BuildColumnMap(_rows.Current);
                        break;
                    case ProjectionStep.Stop:
                        return false;
                    // Skip: loop again.
                }
            }
            return false;
        }

        // Async twin of MoveNext, mirroring AsyncRowEnumerator<T,TReader,TRows>.MoveNextAsync: a non-async
        // fast path that stays synchronous whenever the underlying row-enumerator resolves synchronously
        // (the common case — no second state machine on top of _rows' own), only falling to an awaiting
        // continuation on a genuine buffer miss. Every state mutation (ClassifyRow's ref _rowNumber,
        // BuildColumnMap) runs on the shared class instance, so it survives the await (see class remarks).
        /// <inheritdoc cref="IExcelRowEnumerator.MoveNextAsync"/>
        [SuppressMessage("VisualStudio.Threading", "VSTHRD103:Result synchronously blocks",
            Justification = "The .Result access is guarded by IsCompletedSuccessfully immediately above it — never blocks.")]
        public ValueTask<bool> MoveNextAsync()
        {
            while (true)
            {
                ValueTask<bool> moveTask = _rows.MoveNextAsync();
                if (!moveTask.IsCompletedSuccessfully)
                {
                    return AwaitThenContinueAsync(moveTask);
                }
                if (!moveTask.Result)
                {
                    return new ValueTask<bool>(false);
                }
                switch (ProjectionRules.ClassifyRow(ref _rowNumber, _headerRow, _bindings is not null))
                {
                    case ProjectionStep.Yield:
                        return new ValueTask<bool>(true);
                    case ProjectionStep.BuildMap:
                        BuildColumnMap(_rows.Current);
                        break;
                    case ProjectionStep.Stop:
                        return new ValueTask<bool>(false);
                        // Skip: loop again, still synchronous.
                }
            }
        }

        private async ValueTask<bool> AwaitThenContinueAsync(ValueTask<bool> pendingMoveNext)
        {
            if (!await pendingMoveNext.ConfigureAwait(false))
            {
                return false;
            }
            ProjectionStep step = ProjectionRules.ClassifyRow(ref _rowNumber, _headerRow, _bindings is not null);
            switch (step)
            {
                case ProjectionStep.Yield:
                    return true;
                case ProjectionStep.BuildMap:
                    BuildColumnMap(_rows.Current);
                    break; // map built at the header row — resume the fast path for the next row.
                case ProjectionStep.Stop:
                    return false;
                // Skip: resume the fast path.
            }
            return await MoveNextAsync().ConfigureAwait(false);
        }

        private void BuildColumnMap(Row row)
        {
            _bindings = SparseRowProjection.BuildColumnMap(in row, _typeInfo, _comparer, _normalization, out int requireValueCount);
            _requireValueCount = requireValueCount;
            _seen = requireValueCount > 0 ? new bool[_bindings.Length] : [];
        }

        private void ParseCurrentRow(Row row, ref TModel model)
        {
            bool track = _requireValueCount > 0;
            SparseRowProjection.ParseRow(
                in row, _bindings!, _seen, track, _context.IsDate1904, _context.FormatProvider, _throwOnParseFailure, _rowNumber, ref model);
        }

        /// <inheritdoc/>
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "_rows is created for this enumerator alone by NamedRefRowEnumerable.Get(Async)Enumerator() — owned here, not injected.")]
        public void Dispose()
        {
            _rows.Dispose();
        }

        /// <inheritdoc/>
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "_rows is created for this enumerator alone by NamedRefRowEnumerable.Get(Async)Enumerator() — owned here, not injected.")]
        public ValueTask DisposeAsync()
        {
            return _rows.DisposeAsync();
        }
    }
}
#endif
