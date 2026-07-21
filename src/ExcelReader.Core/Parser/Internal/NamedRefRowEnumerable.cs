#if NET9_0_OR_GREATER
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Parser.Internal
{
    // Zero-model-allocation, attribute-driven row sequence (see NamedRefRowEnumerator). Mirrors
    // RefRowEnumerable<TModel,TReader,TEnumerator>'s shape exactly (struct GetEnumerator for a
    // zero-allocation foreach; IEnumerable<TModel> implemented only for the familiar shape — its
    // members throw, since a ref struct TModel cannot be surfaced through the boxed
    // IEnumerator<TModel>/IEnumerator).
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

        internal NamedRefRowEnumerable(
            TReader reader,
            TypeMapInfo<TModel> typeInfo,
            StringComparer comparer,
            HeaderNormalization normalization,
            int headerRow,
            IFormatProvider? formatProvider)
        {
            _reader = reader;
            _typeInfo = typeInfo;
            _comparer = comparer;
            _normalization = normalization;
            _headerRow = headerRow;
            _context = new ExcelRowContext(reader.IsDate1904, formatProvider ?? CultureInfo.InvariantCulture);
        }

        // The supported way to consume this sequence via foreach.
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Performance", "HLQ006:GetEnumerator should return a value type",
            Justification = "The enumerator is a class so the same type can also expose MoveNextAsync for the async path.")]
        public NamedRefRowEnumerator<TModel, TEnumerator> GetEnumerator()
        {
            return new(_reader.GetEnumerator(), _context, _typeInfo, _comparer, _normalization, _headerRow);
        }

        // The 'await foreach' entry point: C#'s pattern-based async binding picks up this parameterless
        // GetAsyncEnumerator() (the enumerator it returns has MoveNextAsync/Current/DisposeAsync). Opens
        // the sheet synchronously — the reader's GetAsyncEnumerator() is a sync open (no I/O await), the
        // async work is per-row via MoveNextAsync. A ref-struct TModel can't be surfaced through
        // IAsyncEnumerable<TModel> (CS9267), so this stays a pattern match, never the interface.
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Performance", "HLQ006:GetAsyncEnumerator should return a value type",
            Justification = "The enumerator is a class so the same type can also expose MoveNextAsync for the async path.")]
        public NamedRefRowEnumerator<TModel, TEnumerator> GetAsyncEnumerator()
        {
            return new(_reader.GetAsyncEnumerator(), _context, _typeInfo, _comparer, _normalization, _headerRow);
        }

        // Manual-use alternative that opens the sheet asynchronously (awaits the reader's async open).
        // Not reachable by 'await foreach' — its shape (returning a ValueTask of the enumerator) doesn't
        // match the pattern. Await it, then drive the returned enumerator with MoveNextAsync in a loop.
        public async ValueTask<NamedRefRowEnumerator<TModel, TEnumerator>> GetAsyncEnumeratorAsync(CancellationToken ct = default)
        {
            var enumerator = await _reader.GetAsyncEnumeratorAsync(ct).ConfigureAwait(false);
            return new(enumerator, _context, _typeInfo, _comparer, _normalization, _headerRow);
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
