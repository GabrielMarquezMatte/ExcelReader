namespace ExcelReader.NetFramework
{
    /// <summary>A raw row: the non-empty cells the native reader reported.</summary>
    public sealed class ExcelRow
    {
        internal ExcelRow(ExcelCell[] cells)
        {
            Cells = cells;
        }

        /// <summary>The cells of the row, in column order.</summary>
        public IReadOnlyList<ExcelCell> Cells { get; }
    }
}
