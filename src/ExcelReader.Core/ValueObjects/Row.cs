using ExcelReader.Core.Enums;

namespace ExcelReader.Core.ValueObjects
{
    /// <summary>
    /// A single worksheet row, exposed as a zero-allocation view over the reader's underlying buffers.
    /// Only valid for the lifetime of the enumeration that produced it — do not store it past that point.
    /// </summary>
    public readonly ref struct Row
    {
        private readonly ReadOnlySpan<CellDesc> _cells;
        private readonly ReadOnlySpan<byte> _rowValues;
        private readonly ReadOnlySpan<byte> _shared;
        private readonly ReadOnlySpan<byte> _rowBuffer;
        private readonly string?[]? _sharedStringCache;
        private readonly Utf8StringCache? _contentCache;

        internal Row(ReadOnlySpan<CellDesc> cells, ReadOnlySpan<byte> rowValues, ReadOnlySpan<byte> shared,
            ReadOnlySpan<byte> rowBuffer = default, string?[]? sharedStringCache = null, Utf8StringCache? contentCache = null)
        {
            _cells = cells;
            _rowValues = rowValues;
            _shared = shared;
            _rowBuffer = rowBuffer;
            _sharedStringCache = sharedStringCache;
            _contentCache = contentCache;
        }

        /// <summary>One past the highest populated column index, so callers can iterate 0..ColumnCount.</summary>
        public int ColumnCount => _cells.IsEmpty ? 0 : _cells[^1].Column + 1;

        /// <summary>Enumerates only the populated cells in this row, in ascending column order.</summary>
        public RowCellEnumerator Cells => new(_cells, _rowValues, _shared, _rowBuffer, _sharedStringCache, _contentCache);

        internal bool IsEmptyRecord =>
            _cells.IsEmpty || (_cells.Length == 1 && _cells[0].Column == 0 && _cells[0].Type == CellType.Empty);

        /// <summary>Gets the cell at the given column index, or an empty cell if the column has no value.</summary>
        public Cell this[int column]
        {
            get
            {
                int i = IndexOf(column);
                if (i < 0)
                {
                    return new Cell(CellType.Empty, default);
                }
                return _cells[i].ToCell(_rowValues, _shared, _rowBuffer, _sharedStringCache, _contentCache);
            }
        }

        private int IndexOf(int column)
        {
            if ((uint)column < (uint)_cells.Length && _cells[column].Column == column)
            {
                return column;
            }
            int lo = 0, hi = _cells.Length - 1;
            while (lo <= hi)
            {
                int mid = (lo + hi) >> 1;
                int c = _cells[mid].Column;
                if (c == column) { return mid; }
                if (c < column) { lo = mid + 1; }
                else { hi = mid - 1; }
            }
            return -1;
        }
    }
}
