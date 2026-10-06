namespace ExcelReader.Core.Reader
{
    /// <summary>
    /// A workbook reader for the current sheet, typed to the concrete enumerator it hands out
    /// (<typeparamref name="TEnumerator"/>) so callers get zero-copy, format-specific row access
    /// without boxing to the format-agnostic <see cref="IExcelRowEnumerator"/>.
    /// </summary>
    /// <typeparam name="TEnumerator">The concrete row enumerator type this reader produces.</typeparam>
    public interface IExcelRowReader<TEnumerator>
        where TEnumerator : IExcelRowEnumerator
    {
        /// <summary>Gets a value indicating whether the workbook's date system is 1904-based rather than the default 1900-based system.</summary>
        bool IsDate1904 { get; }

        /// <summary>Gets an enumerator that reads the current sheet's rows synchronously from the start.</summary>
        TEnumerator GetEnumerator();

        /// <summary>Gets an enumerator that reads the current sheet's rows asynchronously from the start.</summary>
        /// <remarks>Setup that needs I/O, such as opening the sheet part or loading shared strings, runs on the first <see cref="IExcelRowEnumerator.MoveNextAsync"/>, so this call does not block.</remarks>
        /// <param name="ct">A token observed by that deferred setup and by every <see cref="IExcelRowEnumerator.MoveNextAsync"/> call.</param>
        TEnumerator GetAsyncEnumerator(CancellationToken ct = default);
    }

    /// <summary>
    /// Format-agnostic workbook reader implemented by every concrete reader (XLSX, XLSB, XLS, CSV),
    /// exposing row enumeration plus sheet navigation without requiring callers to downcast to the
    /// concrete reader type.
    /// </summary>
    /// <remarks>
    /// This is the generic <see cref="IExcelRowReader{TEnumerator}"/> specialized to <see cref="IExcelRowEnumerator"/>,
    /// plus a sheet-navigation surface and dispose; unifying them lets the typed parser drive a format-agnostic
    /// reader (<see cref="Excel.Open(string, ExcelReaderOptions?)"/>) and lets callers walk every sheet without
    /// downcasting to the concrete <c>XlsxWorkbook</c>/<c>XlsbWorkbook</c>/<c>XlsWorkbook</c> type.
    /// <para>
    /// <b>Thread safety:</b> enumerators obtained from one reader may be consumed on different threads,
    /// each by one thread at a time, including two enumerators over the same sheet. The reader's own
    /// members are not synchronized: select sheets and create enumerators from one thread at a time. A
    /// CSV reader over a non-seekable stream serves a single enumerator. Disposing the reader closes it
    /// to new enumerators; its file or stream is released when the last enumerator is disposed.
    /// </para>
    /// </remarks>
    public interface IExcelRowReader : IExcelRowReader<IExcelRowEnumerator>, IDisposable, IAsyncDisposable
    {
        /// <summary>Gets the name of the currently selected sheet.</summary>
        string SheetName { get; }

        /// <summary>Gets the number of sheets in the workbook.</summary>
        int SheetCount { get; }

        /// <summary>Gets the name of the sheet at the given zero-based index, without changing the current sheet.</summary>
        /// <param name="index">The zero-based sheet index. Must be within <c>[0, SheetCount)</c>.</param>
        string SheetNameAt(int index);

        /// <summary>Gets whether the currently selected sheet is shown in the workbook's tab bar.</summary>
        /// <remarks>Reported, never enforced: a hidden sheet enumerates its rows like any other, so a caller
        /// that wants to skip one — a converter, say — filters on this itself.</remarks>
        ExcelSheetVisibility SheetVisibility { get; }

        /// <summary>Gets the visibility of the sheet at the given zero-based index, without changing the current sheet.</summary>
        /// <param name="index">The zero-based sheet index. Must be within <c>[0, SheetCount)</c>.</param>
        ExcelSheetVisibility SheetVisibilityAt(int index);

        /// <summary>Attempts to select the sheet with the given name (case-insensitive) as the current sheet.</summary>
        /// <param name="name">The sheet name to look for.</param>
        /// <returns><see langword="true"/> if a matching sheet was found and selected; otherwise <see langword="false"/>.</returns>
        bool TryMoveToSheet(ReadOnlySpan<char> name);

        /// <summary>Selects the sheet at the given zero-based index as the current sheet.</summary>
        /// <param name="index">The zero-based sheet index. Must be within <c>[0, SheetCount)</c>.</param>
        void MoveToSheet(int index);
    }

    /// <summary>
    /// A forward-only cursor over a sheet's rows, implemented by every concrete format's row
    /// enumerator and driven either synchronously (<see cref="MoveNext"/>) or asynchronously
    /// (<see cref="MoveNextAsync"/>).
    /// </summary>
    public interface IExcelRowEnumerator : IDisposable, IAsyncDisposable
    {
        /// <summary>Gets the row at the enumerator's current position. Only valid after a call to <see cref="MoveNext"/> or <see cref="MoveNextAsync"/> has returned <see langword="true"/>.</summary>
        Row Current { get; }

        /// <summary>Advances the enumerator to the next row, reading synchronously.</summary>
        /// <returns><see langword="true"/> if a row was read; <see langword="false"/> if the sheet is exhausted.</returns>
        bool MoveNext();

        /// <summary>Advances the enumerator to the next row, reading asynchronously.</summary>
        /// <returns><see langword="true"/> if a row was read; <see langword="false"/> if the sheet is exhausted.</returns>
        ValueTask<bool> MoveNextAsync();
    }
}
