using System.Collections;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Core.Reader
{
    /// <summary>The sheets of a workbook, in workbook order. Indexing or enumerating the list opens nothing.</summary>
    /// <typeparam name="TSheet">The workbook's sheet type.</typeparam>
    public sealed class ExcelSheetList<TSheet> : IReadOnlyList<TSheet>
    {
        private readonly TSheet[] _sheets;

        internal ExcelSheetList(TSheet[] sheets)
        {
            _sheets = sheets;
        }

        /// <summary>Gets the number of sheets.</summary>
        public int Count
        {
            get
            {
                return _sheets.Length;
            }
        }

        /// <summary>Gets the sheet at the given zero-based index.</summary>
        /// <param name="index">The zero-based sheet index. Must be within <c>[0, Count)</c>.</param>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is outside <c>[0, Count)</c>.</exception>
        public TSheet this[int index]
        {
            get
            {
                WorkbookLookups.ValidateSheetIndex(index, _sheets.Length);
                return _sheets[index];
            }
        }

        /// <summary>Gets an enumerator over the sheets that allocates nothing.</summary>
        public ArraySegment<TSheet>.Enumerator GetEnumerator()
        {
            return new ArraySegment<TSheet>(_sheets).GetEnumerator();
        }

        IEnumerator<TSheet> IEnumerable<TSheet>.GetEnumerator()
        {
            return GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
