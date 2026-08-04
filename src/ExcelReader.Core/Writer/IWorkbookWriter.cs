namespace ExcelReader.Core.Writer
{
    /// <summary>
    /// Writes a whole workbook (XLSX, XLSB, XLS or CSV) to a destination stream, one sheet at a time.
    /// Implementations open the underlying stream/archive on construction, so a caller only needs to
    /// call <see cref="StartAsync"/>, add sheets with <see cref="AddSheet"/>, and dispose the workbook
    /// (or call <see cref="EndAsync"/> then dispose) when finished.
    /// </summary>
    /// <typeparam name="TSheet">The concrete <see cref="ISheetWriter{TRow}"/> this workbook produces.</typeparam>
    /// <remarks>
    /// <b>Thread safety:</b> no implementation is thread-safe, nor are the <see cref="ISheetWriter{TRow}"/>/
    /// <see cref="IRowWriter"/> instances it hands out. All of them carry mutable state (current
    /// sheet/row, buffered output) with no synchronization. Use one workbook writer per thread; do not
    /// call into a writer, its current sheet, or its current row from more than one thread at a time.
    /// </remarks>
    public interface IWorkbookWriter<out TSheet> : IAsyncDisposable
    {
        /// <summary>
        /// Writes the workbook's leading structure (e.g. archive/package headers) and moves the writer
        /// into the started state. Must be called exactly once, before <see cref="AddSheet"/>.
        /// </summary>
        /// <param name="ct">A token to cancel the operation.</param>
        ValueTask StartAsync(CancellationToken ct = default);

        /// <summary>
        /// Begins a new sheet named <paramref name="name"/> and returns its writer. Only one sheet may be
        /// open (started and not yet ended) at a time; the previous sheet's writer must be ended and
        /// disposed before starting the next one.
        /// </summary>
        /// <param name="name">The sheet's name, shown to Excel; must be unique within the workbook and
        /// meet Excel's sheet-name restrictions (1-31 characters, no <c>: \ / ? * [ ]</c>).</param>
        /// <returns>The writer for the new sheet.</returns>
        TSheet AddSheet(string name);

        /// <summary>
        /// Finalizes the workbook, writing any trailing structure (e.g. the workbook manifest/index)
        /// so the destination stream contains a complete, readable file. No further sheets may be added
        /// afterward. Idempotent with disposal: disposing an already-ended workbook is a no-op beyond
        /// releasing resources.
        /// </summary>
        /// <param name="ct">A token to cancel the operation.</param>
        ValueTask EndAsync(CancellationToken ct = default);

        /// <summary>
        /// Flushes any buffered output for the workbook and its current sheet to the destination stream.
        /// </summary>
        /// <param name="ct">A token to cancel the operation.</param>
        ValueTask FlushAsync(CancellationToken ct = default);

        /// <summary>
        /// Registers <paramref name="style"/> and returns its index, reusing the index of an
        /// already-registered style with the same value. Index 0 is always the general/default
        /// style; index 1 is always the builtin date style every date cell already uses.
        /// </summary>
        /// <param name="style">The number format/bold/italic combination to register.</param>
        /// <returns>The style's index, for use with <see cref="ISheetWriter{TRow}.SetColumnStyle"/> and <see cref="ISheetWriter{TRow}.StartRowAsync(int, CancellationToken)"/>.</returns>
        int AddStyle(CellStyle style);
    }

    /// <summary>
    /// Writes one sheet's rows to the workbook. Obtained from <see cref="IWorkbookWriter{TSheet}.AddSheet"/>;
    /// a caller starts the sheet, writes rows in order via <see cref="StartRowAsync(CancellationToken)"/>, then ends the sheet.
    /// </summary>
    /// <typeparam name="TRow">The concrete <see cref="IRowWriter"/> this sheet produces.</typeparam>
    public interface ISheetWriter<TRow> : IAsyncDisposable
    {
        /// <summary>
        /// Writes the sheet's leading structure and moves it into the started state. Must be called
        /// exactly once, before <see cref="StartRowAsync(CancellationToken)"/>.
        /// </summary>
        /// <param name="ct">A token to cancel the operation.</param>
        ValueTask StartAsync(CancellationToken ct = default);

        /// <summary>
        /// Begins the next row, in order starting from the first, and returns its writer. Only one row
        /// may be open (started and not yet disposed) at a time; the previous row's writer must be
        /// disposed before starting the next one.
        /// </summary>
        /// <param name="ct">A token to cancel the operation.</param>
        /// <returns>The writer for the new row.</returns>
        ValueTask<TRow> StartRowAsync(CancellationToken ct = default);

        /// <summary>
        /// Begins the next row, applying <paramref name="styleId"/> (from <see cref="IWorkbookWriter{TSheet}.AddStyle"/>)
        /// to its cells.
        /// </summary>
        /// <param name="styleId">The style to apply to every cell of this row.</param>
        /// <param name="ct">A token to cancel the operation.</param>
        /// <returns>The writer for the new row.</returns>
        ValueTask<TRow> StartRowAsync(int styleId, CancellationToken ct = default);

        /// <summary>
        /// Applies <paramref name="styleId"/> to every cell of column <paramref name="columnIndex"/>
        /// that does not carry its own row style. Must be called before <see cref="StartAsync"/>.
        /// </summary>
        /// <param name="columnIndex">The 0-based column index.</param>
        /// <param name="styleId">The style to apply, from <see cref="IWorkbookWriter{TSheet}.AddStyle"/>.</param>
        /// <exception cref="InvalidOperationException">The sheet has already been started.</exception>
        void SetColumnStyle(int columnIndex, int styleId);

        /// <summary>
        /// Sets the display width, in characters, of column <paramref name="columnIndex"/>. Must be
        /// called before <see cref="StartAsync"/>.
        /// </summary>
        /// <param name="columnIndex">The 0-based column index.</param>
        /// <param name="width">The column width, in characters.</param>
        /// <exception cref="InvalidOperationException">The sheet has already been started.</exception>
        void SetColumnWidth(int columnIndex, double width);

        /// <summary>
        /// Finalizes the sheet, writing any trailing structure so the sheet is complete within the
        /// workbook. The active row writer, if any, must already be disposed. No further rows may be
        /// started afterward.
        /// </summary>
        /// <param name="ct">A token to cancel the operation.</param>
        ValueTask EndAsync(CancellationToken ct = default);
    }

    /// <summary>
    /// Writes one row's cells in column order, starting at the first column. Obtained from
    /// <see cref="ISheetWriter{TRow}.StartRowAsync(CancellationToken)"/>; each <c>Write</c>/<see cref="Skip"/> call advances
    /// to the next column, so cells must be written left-to-right with no way to revisit an earlier
    /// column. Disposing the row writer (sync or async) finalizes the row and returns it to the sheet.
    /// Calling any member after disposal throws <see cref="ObjectDisposedException"/>.
    /// </summary>
    /// <remarks>
    /// On a <see langword="null"/> value, every nullable <c>Write</c> overload (including the nullable
    /// generic overload) still advances the column position by one, but whether it emits an explicit
    /// blank cell or omits a cell record entirely for that column is implementation-defined — either
    /// way, reading the cell back yields an empty/blank value.
    /// </remarks>
    public interface IRowWriter : IAsyncDisposable
    {
        /// <summary>Writes a text cell, or a blank/empty cell if <paramref name="value"/> is <see langword="null"/>.</summary>
        /// <param name="value">The text to write.</param>
        /// <exception cref="ArgumentException"><paramref name="value"/> exceeds Excel's per-cell text limit (32,767 characters).</exception>
        void Write(string? value);

        /// <summary>Writes a boolean cell.</summary>
        /// <param name="value">The value to write.</param>
        void Write(bool value);

        /// <summary>Writes a boolean cell, or a blank/empty cell if <paramref name="value"/> is <see langword="null"/>.</summary>
        /// <param name="value">The value to write.</param>
        void Write(bool? value);

        /// <summary>Writes a date/time cell.</summary>
        /// <param name="value">The value to write.</param>
        void Write(DateTime value);

        /// <summary>Writes a date/time cell, or a blank/empty cell if <paramref name="value"/> is <see langword="null"/>.</summary>
        /// <param name="value">The value to write.</param>
        void Write(DateTime? value);

        /// <summary>Writes a date-only cell (no time component).</summary>
        /// <param name="value">The value to write.</param>
        void Write(DateOnly value);

        /// <summary>Writes a date-only cell, or a blank/empty cell if <paramref name="value"/> is <see langword="null"/>.</summary>
        /// <param name="value">The value to write.</param>
        void Write(DateOnly? value);

        /// <summary>Writes a time-of-day cell (no date component).</summary>
        /// <param name="value">The value to write.</param>
        void Write(TimeOnly value);

        /// <summary>Writes a time-of-day cell, or a blank/empty cell if <paramref name="value"/> is <see langword="null"/>.</summary>
        /// <param name="value">The value to write.</param>
        void Write(TimeOnly? value);

        /// <summary>Writes a numeric cell for any UTF-8-formattable value type (e.g. <see cref="int"/>, <see cref="double"/>, <see cref="decimal"/>).</summary>
        /// <typeparam name="T">The numeric value type to write.</typeparam>
        /// <param name="value">The value to write.</param>
        void Write<T>(T value) where T : IUtf8SpanFormattable;

        /// <summary>Writes a numeric cell, or a blank/empty cell if <paramref name="value"/> is <see langword="null"/>.</summary>
        /// <typeparam name="T">The underlying numeric value type to write.</typeparam>
        /// <param name="value">The value to write.</param>
        void Write<T>(T? value) where T : struct, IUtf8SpanFormattable;

        /// <summary>
        /// Advances past <paramref name="count"/> columns without writing any cell for them, leaving
        /// them blank when the row is read back.
        /// </summary>
        /// <param name="count">The number of columns to skip; must not be negative.</param>
        void Skip(int count = 1);
    }
}
