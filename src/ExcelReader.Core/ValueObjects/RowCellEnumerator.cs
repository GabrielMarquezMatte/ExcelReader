namespace ExcelReader.Core.ValueObjects
{
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
