namespace ExcelReader.Core.ValueObjects
{
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
}
