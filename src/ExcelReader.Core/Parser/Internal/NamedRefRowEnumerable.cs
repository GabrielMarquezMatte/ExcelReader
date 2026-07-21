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

        // Zero-allocation struct enumerator — the supported way to consume this sequence.
        public NamedRefRowEnumerator<TModel, TEnumerator> GetEnumerator()
        {
            return new NamedRefRowEnumerator<TModel, TEnumerator>(
                _reader.GetEnumerator(), _context, _typeInfo, _comparer, _normalization, _headerRow);
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
