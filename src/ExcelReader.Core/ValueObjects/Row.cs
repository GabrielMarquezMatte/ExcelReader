using ExcelReader.Core.Enums;

namespace ExcelReader.Core.ValueObjects
{
    public readonly ref struct Row
    {
        private readonly ReadOnlySpan<CellDesc> _cells; // ascending by Column, gaps allowed
        private readonly ReadOnlySpan<byte> _rowValues;
        private readonly ReadOnlySpan<byte> _shared;

        internal Row(ReadOnlySpan<CellDesc> cells, ReadOnlySpan<byte> rowValues, ReadOnlySpan<byte> shared)
        {
            _cells = cells;
            _rowValues = rowValues;
            _shared = shared;
        }

        // One past the highest populated column, so callers can iterate 0..ColumnCount.
        public int ColumnCount => _cells.IsEmpty ? 0 : _cells[^1].Column + 1;

        // Populated cells only, in ascending column order. Skips gaps instead of binary-searching them.
        public RowCells Cells => new(_cells, _rowValues, _shared);

        public Cell this[int column]
        {
            get
            {
                int i = IndexOf(column);
                if (i < 0)
                {
                    return new Cell(CellType.Empty, default);
                }
                ref readonly var d = ref _cells[i];
                var buf = d.FromShared ? _shared : _rowValues;
                return new Cell(d.Type, buf.Slice(d.Start, d.Length), d.Number, d.HasNumber, d.Style);
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

    public readonly ref struct RowCell
    {
        public RowCell(int columnIndex, Cell value)
        {
            ColumnIndex = columnIndex;
            Value = value;
        }

        public int ColumnIndex { get; }
        public Cell Value { get; }
    }

    public readonly ref struct RowCells
    {
        private readonly ReadOnlySpan<CellDesc> _cells;
        private readonly ReadOnlySpan<byte> _rowValues;
        private readonly ReadOnlySpan<byte> _shared;

        internal RowCells(ReadOnlySpan<CellDesc> cells, ReadOnlySpan<byte> rowValues, ReadOnlySpan<byte> shared)
        {
            _cells = cells;
            _rowValues = rowValues;
            _shared = shared;
        }

        public RowCellEnumerator GetEnumerator()
        {
            return new RowCellEnumerator(_cells, _rowValues, _shared);
        }
    }

    public ref struct RowCellEnumerator
    {
        private readonly ReadOnlySpan<CellDesc> _cells;
        private readonly ReadOnlySpan<byte> _rowValues;
        private readonly ReadOnlySpan<byte> _shared;
        private int _index;

        internal RowCellEnumerator(ReadOnlySpan<CellDesc> cells, ReadOnlySpan<byte> rowValues, ReadOnlySpan<byte> shared)
        {
            _cells = cells;
            _rowValues = rowValues;
            _shared = shared;
            _index = -1;
        }

        public readonly RowCell Current
        {
            get
            {
                ref readonly var d = ref _cells[_index];
                var buf = d.FromShared ? _shared : _rowValues;
                return new RowCell(d.Column, new Cell(d.Type, buf.Slice(d.Start, d.Length), d.Number, d.HasNumber, d.Style));
            }
        }

        public bool MoveNext()
        {
            int next = _index + 1;
            if (next >= _cells.Length)
            {
                return false;
            }
            _index = next;
            return true;
        }
    }
}
