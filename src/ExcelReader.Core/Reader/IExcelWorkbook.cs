using System.Diagnostics.CodeAnalysis;

namespace ExcelReader.Core.Reader
{
    /// <summary>
    /// An open workbook of any supported format (XLSX, XLSB, XLS, CSV): its sheets and the state they
    /// share. Sheets are obtained from it and read independently.
    /// </summary>
    /// <remarks>
    /// <b>Thread safety:</b> once opened, a workbook is safe to use from several threads: reading its
    /// metadata, obtaining sheets and creating enumerators. Each enumerator is used by one thread at a
    /// time; several enumerators of one workbook, including two over the same sheet, may run on
    /// different threads. The exception is a CSV over a non-seekable stream, which serves a single
    /// enumerator: a second one throws <see cref="InvalidOperationException"/>. Using a single
    /// enumerator from two threads at once is undefined behaviour and is not checked. A stream or
    /// buffer passed to a workbook must not be used, modified or disposed by the caller until the
    /// workbook and every enumerator obtained from it are disposed.
    /// <para>
    /// Disposing the workbook closes it to new sheets and enumerators. Its file, stream and pooled
    /// buffers are released when the last enumerator obtained from it is disposed.
    /// </para>
    /// </remarks>
    public interface IExcelWorkbook : IDisposable, IAsyncDisposable
    {
        /// <summary>Gets a value indicating whether the workbook's date system is 1904-based rather than the default 1900-based system.</summary>
        bool IsDate1904 { get; }

        /// <summary>Gets the number of sheets in the workbook.</summary>
        int SheetCount { get; }

        /// <summary>Gets the workbook's first sheet: the same sheet as <c>SheetAt(0)</c>. Opens nothing.</summary>
        /// <exception cref="ObjectDisposedException">The workbook was disposed.</exception>
        IExcelSheet FirstSheet { get; }

        /// <summary>Gets the sheet at the given zero-based index. Opens nothing.</summary>
        /// <param name="index">The zero-based sheet index. Must be within <c>[0, SheetCount)</c>.</param>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is outside <c>[0, SheetCount)</c>.</exception>
        /// <exception cref="ObjectDisposedException">The workbook was disposed.</exception>
        IExcelSheet SheetAt(int index);

        /// <summary>Finds a sheet by name, ignoring case. Opens nothing.</summary>
        /// <param name="name">The sheet name to look for.</param>
        /// <param name="sheet">The matching sheet, when one is found.</param>
        /// <returns><see langword="true"/> if a sheet with that name exists; otherwise <see langword="false"/>.</returns>
        /// <exception cref="ObjectDisposedException">The workbook was disposed.</exception>
        bool TryGetSheet(ReadOnlySpan<char> name, [MaybeNullWhen(false)] out IExcelSheet sheet);

        /// <summary>
        /// Gets the workbook's tables (ranges formatted with "Format as Table"), ordered by sheet and then by
        /// the order they were created in. Always empty for XLS and CSV. Opens nothing.
        /// </summary>
        /// <exception cref="ObjectDisposedException">The workbook was disposed.</exception>
        IReadOnlyList<ExcelTable> Tables { get; }

        /// <summary>Finds a table by name, ignoring case. Opens nothing.</summary>
        /// <param name="name">The table name to look for.</param>
        /// <param name="table">The matching table, when one is found.</param>
        /// <returns><see langword="true"/> if a table with that name exists; otherwise <see langword="false"/>.</returns>
        /// <exception cref="ObjectDisposedException">The workbook was disposed.</exception>
        bool TryGetTable(ReadOnlySpan<char> name, [MaybeNullWhen(false)] out ExcelTable table);
    }

    /// <summary>
    /// One sheet's rows, typed to the concrete enumerator it hands out
    /// (<typeparamref name="TEnumerator"/>) so callers read without boxing to
    /// <see cref="IExcelRowEnumerator"/>.
    /// </summary>
    /// <typeparam name="TEnumerator">The concrete row enumerator type this sheet produces.</typeparam>
    public interface IExcelSheet<TEnumerator>
        where TEnumerator : IExcelRowEnumerator
    {
        /// <summary>Gets a value indicating whether the sheet's workbook uses the 1904 date system.</summary>
        bool IsDate1904 { get; }

        /// <summary>Gets a new enumerator that reads the sheet's rows synchronously from the start.</summary>
        /// <exception cref="ObjectDisposedException">The sheet's workbook was disposed.</exception>
        /// <exception cref="InvalidOperationException">The source is a CSV over a non-seekable stream that was already enumerated.</exception>
        TEnumerator GetEnumerator();

        /// <summary>Gets a new enumerator that reads the sheet's rows asynchronously from the start.</summary>
        /// <remarks>Setup that needs I/O, such as opening the sheet part or loading shared strings, runs on the first <see cref="IExcelRowEnumerator.MoveNextAsync"/>, so this call does not block.</remarks>
        /// <param name="ct">A token observed by that deferred setup and by every <see cref="IExcelRowEnumerator.MoveNextAsync"/> call.</param>
        /// <exception cref="ObjectDisposedException">The sheet's workbook was disposed.</exception>
        /// <exception cref="InvalidOperationException">The source is a CSV over a non-seekable stream that was already enumerated.</exception>
        TEnumerator GetAsyncEnumerator(CancellationToken ct = default);
    }

    /// <summary>A sheet of an <see cref="IExcelWorkbook"/>: its position, name and visibility, and its rows.</summary>
    /// <remarks>A sheet holds no resources. Every call to <see cref="IExcelSheet{TEnumerator}.GetEnumerator"/> starts an independent read.</remarks>
    public interface IExcelSheet : IExcelSheet<IExcelRowEnumerator>
    {
        /// <summary>Gets the sheet's zero-based position in its workbook.</summary>
        int Index { get; }

        /// <summary>Gets the sheet's name. Empty for the single sheet of a CSV source.</summary>
        string Name { get; }

        /// <summary>Gets whether the sheet is shown in the workbook's tab bar.</summary>
        /// <remarks>Reported, never enforced: a hidden sheet enumerates its rows like any other.</remarks>
        ExcelSheetVisibility Visibility { get; }
    }
}
