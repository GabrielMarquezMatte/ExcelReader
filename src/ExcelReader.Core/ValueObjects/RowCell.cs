namespace ExcelReader.Core.ValueObjects
{
    /// <summary>A cell paired with its column index, as produced by <see cref="RowCellEnumerator"/>.</summary>
    public readonly ref struct RowCell
    {
        /// <summary>Creates a row cell with the given column index and value.</summary>
        public RowCell(int columnIndex, Cell value)
        {
            ColumnIndex = columnIndex;
            Value = value;
        }
        /// <summary>The zero-based column index of this cell within its row.</summary>
        public int ColumnIndex { get; }
        /// <summary>The cell's value.</summary>
        public Cell Value { get; }
    }
}
