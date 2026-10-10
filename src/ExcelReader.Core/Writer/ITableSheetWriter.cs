using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Writer
{
    /// <summary>
    /// A sheet writer that can write Excel tables (ranges formatted with "Format as Table"). Implemented by the
    /// XLSX and XLSB sheet writers.
    /// </summary>
    /// <typeparam name="TRow">The concrete row writer type.</typeparam>
    public interface ITableSheetWriter<TRow> : ISheetWriter<TRow>
    {
        /// <summary>
        /// Writes a header row with <paramref name="columns"/> at the sheet's next row and opens a table there.
        /// The table covers every row written until <see cref="EndTable"/> or the end of the sheet.
        /// </summary>
        /// <param name="name">The table name: unique in the workbook, ignoring case, and valid as an Excel name.</param>
        /// <param name="columns">The column names, written as the header row.</param>
        /// <param name="options">Position and style. Defaults to <see cref="ExcelTableOptions.Default"/>.</param>
        /// <exception cref="ArgumentNullException"><paramref name="name"/> or <paramref name="columns"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException">The name, a column name or the style is not valid, or the name is already used in this workbook.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><see cref="ExcelTableOptions.FirstColumn"/> is negative.</exception>
        /// <exception cref="InvalidOperationException">A table is already open on this sheet, or a row is still active.</exception>
        /// <exception cref="ObjectDisposedException">The sheet has already been ended.</exception>
        /// <exception cref="ExcelLimitExceededException">The sheet has no room left for a header row and a data row.</exception>
        void BeginTable(string name, IReadOnlyList<string> columns, ExcelTableOptions? options = null);

        /// <inheritdoc cref="BeginTable"/>
        /// <param name="name">The table name: unique in the workbook, ignoring case, and valid as an Excel name.</param>
        /// <param name="columns">The column names, written as the header row.</param>
        /// <param name="options">Position and style. Defaults to <see cref="ExcelTableOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel the header write.</param>
        ValueTask BeginTableAsync(string name, IReadOnlyList<string> columns, ExcelTableOptions? options = null, CancellationToken ct = default);

        /// <summary>Closes the open table at the last row written. Writes nothing; the table is stored when the sheet ends.</summary>
        /// <exception cref="InvalidOperationException">No table is open on this sheet.</exception>
        /// <exception cref="ObjectDisposedException">The sheet has already been ended.</exception>
        void EndTable();
    }
}
