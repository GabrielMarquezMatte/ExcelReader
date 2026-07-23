namespace ExcelReader.Core.ValueObjects
{
    /// <summary>
    /// Enumerates the populated cells of a <see cref="Row"/>, in ascending column order, skipping gaps.
    /// Supports <c>foreach</c> via the duck-typed enumerator pattern (ref structs cannot implement
    /// <see cref="System.Collections.Generic.IEnumerator{T}"/>).
    /// </summary>
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
        /// <summary>Returns this enumerator, enabling <c>foreach</c> over a <see cref="Row"/>'s cells.</summary>
        public readonly RowCellEnumerator GetEnumerator()
        {
            return this;
        }
        /// <summary>The cell at the enumerator's current position.</summary>
        public readonly RowCell Current
        {
            get
            {
                ref readonly var d = ref _cells[_index];
                return new RowCell(d.Column, d.ToCell(_rowValues, _shared, _sharedStringCache));
            }
        }
        /// <summary>Advances to the next populated cell; returns false when enumeration is exhausted.</summary>
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
