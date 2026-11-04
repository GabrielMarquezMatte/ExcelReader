namespace ExcelReader.Core.Writer.Internal
{
    /// <summary>Lets a styled record layout register its styles with the sheet's workbook and apply them per column.</summary>
    internal interface IStyleBinder
    {
        int RegisterStyle(CellStyle style);

        /// <summary>Overrides the column styles of the cells written next; entries below 0 keep the column's own style.</summary>
        void PushColumnStyleOverlay(ReadOnlySpan<int> styleIds, int firstColumn);

        void PopColumnStyleOverlay();

        void BeginStyledTable(string name, IReadOnlyList<string> columns, ExcelTableOptions? options, int headerStyleId)
        {
            throw new NotSupportedException("This sheet does not write tables.");
        }

        ValueTask BeginStyledTableAsync(string name, IReadOnlyList<string> columns, ExcelTableOptions? options, int headerStyleId, CancellationToken ct)
        {
            throw new NotSupportedException("This sheet does not write tables.");
        }
    }
}
