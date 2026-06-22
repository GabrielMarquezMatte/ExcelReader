using ExcelReader.Core.Enums;

namespace ExcelReader.Core.ValueObjects
{
    // One parsed cell, located in either the shared-strings flat buffer or the per-row value buffer.
    internal readonly struct CellDesc
    {
        public int Column { get; init; }
        public int Start { get; init; }
        public int Length { get; init; }
        public CellType Type { get; init; }
        public int Style { get; init; }
        public bool FromShared { get; init; }
    }

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
                return new Cell(d.Type, buf.Slice(d.Start, d.Length), d.Style);
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
