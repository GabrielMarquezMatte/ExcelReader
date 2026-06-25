namespace ExcelReader.Core.ValueObjects
{
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
}
