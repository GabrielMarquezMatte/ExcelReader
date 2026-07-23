#if NET9_0_OR_GREATER
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Parser.Internal
{
    /// <summary>A zero-model-allocation sequence of rows bound by attribute to a ref struct model type, consumable via <c>foreach</c> or <c>await foreach</c>.</summary>
    /// <typeparam name="TModel">The ref-struct-capable row model type to bind each row to.</typeparam>
    /// <typeparam name="TReader">The concrete row reader type this instance pulls rows from.</typeparam>
    /// <typeparam name="TEnumerator">The concrete row enumerator type <typeparamref name="TReader"/> produces.</typeparam>
    /// <remarks>
    /// Mirrors <c>RefRowEnumerable&lt;TModel,TReader,TEnumerator&gt;</c>'s shape exactly: a struct
    /// <c>GetEnumerator()</c> gives a zero-allocation <c>foreach</c>, while <see cref="IEnumerable{TModel}"/>
    /// is implemented only for the familiar shape — its members throw, since a ref struct
    /// <typeparamref name="TModel"/> cannot be surfaced through the boxed <c>IEnumerator&lt;TModel&gt;</c>/<c>IEnumerator</c>.
    /// </remarks>
    public sealed class NamedRefRowEnumerable<TModel, TReader, TEnumerator> : IEnumerable<TModel>
        where TModel : allows ref struct
        where TReader : IExcelRowReader<TEnumerator>
        where TEnumerator : class, IExcelRowEnumerator
    {
        private readonly TReader _reader;
        private readonly ExcelRowContext _context;
        private readonly TypeMapInfo<TModel> _typeInfo;
        private readonly StringComparer _comparer;
        private readonly HeaderNormalization _normalization;
        private readonly int _headerRow;
        private readonly bool _throwOnParseFailure;

        internal NamedRefRowEnumerable(
            TReader reader,
            TypeMapInfo<TModel> typeInfo,
            StringComparer comparer,
            HeaderNormalization normalization,
            int headerRow,
            IFormatProvider? formatProvider,
            bool throwOnParseFailure = false)
        {
            _reader = reader;
            _typeInfo = typeInfo;
            _comparer = comparer;
            _normalization = normalization;
            _headerRow = headerRow;
            _throwOnParseFailure = throwOnParseFailure;
            _context = new ExcelRowContext(reader.IsDate1904, formatProvider ?? CultureInfo.InvariantCulture);
        }

        /// <summary>Gets the enumerator used to consume this sequence with a <c>foreach</c> loop.</summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Performance", "HLQ006:GetEnumerator should return a value type",
            Justification = "The enumerator is a class so the same type can also expose MoveNextAsync for the async path.")]
        public NamedRefRowEnumerator<TModel, TEnumerator> GetEnumerator()
        {
            return new(_reader.GetEnumerator(), _context, _typeInfo, _comparer, _normalization, _headerRow, _throwOnParseFailure);
        }

        /// <summary>Gets the enumerator used to consume this sequence with an <c>await foreach</c> loop, opening the underlying sheet synchronously.</summary>
        /// <remarks>
        /// The <c>await foreach</c> entry point: C#'s pattern-based async binding picks up this
        /// parameterless <c>GetAsyncEnumerator()</c> (the enumerator it returns has
        /// <c>MoveNextAsync</c>/<c>Current</c>/<c>DisposeAsync</c>). The async work happens per-row via
        /// <c>MoveNextAsync</c>. A ref-struct <typeparamref name="TModel"/> can't be surfaced through
        /// <c>IAsyncEnumerable&lt;TModel&gt;</c> (CS9267), so this stays a pattern match, never the interface.
        /// </remarks>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Performance", "HLQ006:GetAsyncEnumerator should return a value type",
            Justification = "The enumerator is a class so the same type can also expose MoveNextAsync for the async path.")]
        public NamedRefRowEnumerator<TModel, TEnumerator> GetAsyncEnumerator()
        {
            return new(_reader.GetAsyncEnumerator(), _context, _typeInfo, _comparer, _normalization, _headerRow, _throwOnParseFailure);
        }

        /// <summary>Asynchronously opens the underlying sheet and returns an enumerator to drive manually with <c>MoveNextAsync</c>, for callers who cannot use <c>await foreach</c>.</summary>
        /// <param name="ct">A token to cancel the open operation.</param>
        /// <returns>An enumerator positioned before the first row.</returns>
        // Manual-use alternative that opens the sheet asynchronously (awaits the reader's async open).
        // Not reachable by 'await foreach' — its shape (returning a ValueTask of the enumerator) doesn't
        // match the pattern. Await it, then drive the returned enumerator with MoveNextAsync in a loop.
        public async ValueTask<NamedRefRowEnumerator<TModel, TEnumerator>> GetAsyncEnumeratorAsync(CancellationToken ct = default)
        {
            var enumerator = await _reader.GetAsyncEnumeratorAsync(ct).ConfigureAwait(false);
            return new(enumerator, _context, _typeInfo, _comparer, _normalization, _headerRow, _throwOnParseFailure);
        }

        IEnumerator<TModel> IEnumerable<TModel>.GetEnumerator()
        {
            throw new NotSupportedException(
                "A ref struct row model cannot be enumerated through IEnumerable<T>. Use a direct 'foreach' over this type instead, which binds to its struct GetEnumerator().");
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            throw new NotSupportedException(
                "A ref struct row model cannot be enumerated through IEnumerable. Use a direct 'foreach' over this type instead.");
        }
    }
}
#endif
