namespace ExcelReader.Core.ValueObjects
{
    public ref struct RowCellEnumerator
    {
        private readonly ReadOnlySpan<CellDesc> _cells;
        private readonly ReadOnlySpan<byte> _rowValues;
        private readonly ReadOnlySpan<byte> _shared;
        private readonly Dictionary<int, string>? _sharedStringCache;
        private int _index;
        internal RowCellEnumerator(ReadOnlySpan<CellDesc> cells, ReadOnlySpan<byte> rowValues, ReadOnlySpan<byte> shared,
            Dictionary<int, string>? sharedStringCache = null)
        {
            _cells = cells;
            _rowValues = rowValues;
            _shared = shared;
            _sharedStringCache = sharedStringCache;
            _index = -1;
        }
        public readonly RowCellEnumerator GetEnumerator()
        {
            return this;
        }
        public readonly RowCell Current
        {
            get
            {
                ref readonly var d = ref _cells[_index];
                return new RowCell(d.Column, d.ToCell(_rowValues, _shared, _sharedStringCache));
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
