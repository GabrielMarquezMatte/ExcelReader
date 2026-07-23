using ExcelReader.Core.ValueObjects;

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
        TEnumerator GetAsyncEnumerator();

        /// <summary>Asynchronously creates an enumerator that reads the current sheet's rows from the start, performing any setup that requires I/O before the first row is fetched.</summary>
        /// <param name="ct">A token to cancel the setup operation.</param>
        ValueTask<TEnumerator> GetAsyncEnumeratorAsync(CancellationToken ct = default);
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
    /// downcasting to the concrete <c>XlsxReader</c>/<c>XlsbReader</c>/<c>XlsReader</c> type.
    /// </remarks>
    public interface IExcelRowReader : IExcelRowReader<IExcelRowEnumerator>, IDisposable, IAsyncDisposable
    {
        /// <summary>Gets the name of the currently selected sheet.</summary>
        string SheetName { get; }

        /// <summary>Gets the number of sheets in the workbook.</summary>
        int SheetCount { get; }

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