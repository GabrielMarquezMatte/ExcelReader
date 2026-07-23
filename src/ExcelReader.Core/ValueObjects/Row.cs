using ExcelReader.Core.Enums;

namespace ExcelReader.Core.ValueObjects
{
    /// <summary>
    /// A single worksheet row, exposed as a zero-allocation view over the reader's underlying buffers.
    /// Only valid for the lifetime of the enumeration that produced it — do not store it past that point.
    /// </summary>
    public readonly ref struct Row
    {
        private readonly ReadOnlySpan<CellDesc> _cells; // ascending by Column, gaps allowed
        private readonly ReadOnlySpan<byte> _rowValues;
        private readonly ReadOnlySpan<byte> _shared;
        // Non-null only for readers with a true, cross-row-stable shared-string table (XLSX/XLSB); see
        // CellDesc.ToCell. Defaults to null so existing 3-arg call sites (CSV, XLS) are unaffected.
        private readonly Dictionary<int, string>? _sharedStringCache;

        internal Row(ReadOnlySpan<CellDesc> cells, ReadOnlySpan<byte> rowValues, ReadOnlySpan<byte> shared,
            Dictionary<int, string>? sharedStringCache = null)
        {
            _cells = cells;
            _rowValues = rowValues;
            _shared = shared;
            _sharedStringCache = sharedStringCache;
        }

        // One past the highest populated column, so callers can iterate 0..ColumnCount.
        /// <summary>One past the highest populated column index, so callers can iterate 0..ColumnCount.</summary>
        public int ColumnCount => _cells.IsEmpty ? 0 : _cells[^1].Column + 1;

        // Populated cells only, in ascending column order. Skips gaps instead of binary-searching them.
        /// <summary>Enumerates only the populated cells in this row, in ascending column order.</summary>
        public RowCellEnumerator Cells => new(_cells, _rowValues, _shared, _sharedStringCache);

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
                return _cells[i].ToCell(_rowValues, _shared, _sharedStringCache);
            }
        }

        // Binary search by Column (cells are sorted ascending) — keeps wide-row access O(log n).
        private int IndexOf(int column)
        {
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
