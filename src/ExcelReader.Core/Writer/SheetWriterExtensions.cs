using System.Diagnostics.CodeAnalysis;

namespace ExcelReader.Core.Writer
{
    /// <summary>
    /// Convenience methods layered on <see cref="ISheetWriter{TRow}"/>: write a whole collection of
    /// records to a sheet in one call, driving the start-row/write/dispose lifecycle per record.
    /// </summary>
    /// <remarks>
    /// Generic over <c>TRow</c> so callers keep the concrete row writer's full <c>Write</c> overloads
    /// (<see cref="int"/>, <see cref="double"/>, <see cref="DateTime"/>, ...).
    /// </remarks>
    public static class SheetWriterExtensions
    {
        /// <summary>
        /// Writes one row per item in <paramref name="records"/>, calling <paramref name="writeRow"/>
        /// for each to populate that row's cells.
        /// </summary>
        /// <typeparam name="T">The record type being written.</typeparam>
        /// <typeparam name="TRow">The concrete row writer type.</typeparam>
        /// <param name="sheet">The sheet to write rows to.</param>
        /// <param name="records">The records to write, one row each, in enumeration order.</param>
        /// <param name="writeRow">Populates a row's cells from a single record.</param>
        /// <param name="ct">A token to cancel the operation between rows.</param>
        public static async ValueTask WriteRecordsAsync<T, TRow>(this ISheetWriter<TRow> sheet, IEnumerable<T> records,
                                                                 Action<TRow, T> writeRow,
                                                                 CancellationToken ct = default)
                                                                 where TRow : IRowWriter
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(records);
            ArgumentNullException.ThrowIfNull(writeRow);
            foreach (T record in records)
            {
                ct.ThrowIfCancellationRequested();
                TRow row = await sheet.StartRowAsync(ct).ConfigureAwait(false);
                writeRow(row, record);
                await row.DisposeAsync().ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Writes one row per item produced by <paramref name="records"/>, calling <paramref name="writeRow"/>
        /// for each to populate that row's cells.
        /// </summary>
        /// <typeparam name="T">The record type being written.</typeparam>
        /// <typeparam name="TRow">The concrete row writer type.</typeparam>
        /// <param name="sheet">The sheet to write rows to.</param>
        /// <param name="records">The records to write, one row each, in enumeration order.</param>
        /// <param name="writeRow">Populates a row's cells from a single record.</param>
        /// <param name="ct">A token to cancel the operation between rows, and passed to the source enumerable.</param>
        public static async ValueTask WriteRecordsAsync<T, TRow>(this ISheetWriter<TRow> sheet,
                                                                 IAsyncEnumerable<T> records, Action<TRow, T> writeRow,
                                                                 CancellationToken ct = default)
                                                                 where TRow : IRowWriter
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(records);
            ArgumentNullException.ThrowIfNull(writeRow);
            await foreach (T record in records.WithCancellation(ct).ConfigureAwait(false))
            {
                TRow row = await sheet.StartRowAsync(ct).ConfigureAwait(false);
                writeRow(row, record);
                await row.DisposeAsync().ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Writes one row per item in <paramref name="records"/> to an <see cref="XlsxSheetWriter"/> using
        /// its synchronous fast path, calling <paramref name="writeRow"/> for each to populate that row's
        /// cells. Behaviorally equivalent to the generic <see cref="ISheetWriter{TRow}"/> overload, but
        /// resolved in preference to it when the caller's static sheet type is <see cref="XlsxSheetWriter"/>
        /// — e.g. plain <see cref="IEnumerable{T}"/> sources, which are the overwhelming majority.
        /// </summary>
        /// <remarks>
        /// Safe because <see cref="XlsxSheetWriter"/>/<see cref="XlsxRowWriter"/> row buffering only ever
        /// touches the destination stream on the rare buffer-threshold flush (<c>EndBufferedRow</c>,
        /// already synchronous), so this overload skips the per-row <see cref="ValueTask"/>/async-disposable
        /// machinery entirely.
        /// </remarks>
        /// <typeparam name="T">The record type being written.</typeparam>
        /// <param name="sheet">The sheet to write rows to.</param>
        /// <param name="records">The records to write, one row each, in enumeration order.</param>
        /// <param name="writeRow">Populates a row's cells from a single record.</param>
        /// <param name="ct">A token to cancel the operation between rows.</param>
        [SuppressMessage("Reliability", "CA1849:Call async methods when in an async method",
            Justification = "Deliberately using XlsxSheetWriter/XlsxRowWriter's synchronous fast path — see the <remarks> above.")]
        [SuppressMessage("Sonar", "S6966:Await StartRowAsync instead",
            Justification = "Deliberately using XlsxSheetWriter's synchronous row fast path — see the <remarks> above.")]
        [SuppressMessage("VisualStudio.Threading", "VSTHRD103:StartRow synchronously blocks",
            Justification = "Deliberately using XlsxSheetWriter's synchronous row fast path — see the <remarks> above.")]
        [SuppressMessage("Sonar", "S6966:Await DisposeAsync instead",
            Justification = "Deliberately using XlsxRowWriter's synchronous Dispose fast path — see the <remarks> above.")]
        [SuppressMessage("VisualStudio.Threading", "VSTHRD103:Dispose synchronously blocks",
            Justification = "Deliberately using XlsxRowWriter's synchronous Dispose fast path — see the <remarks> above.")]
        public static ValueTask WriteRecordsAsync<T>(this XlsxSheetWriter sheet, IEnumerable<T> records,
                                                      Action<XlsxRowWriter, T> writeRow, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(records);
            ArgumentNullException.ThrowIfNull(writeRow);
            foreach (T record in records)
            {
                ct.ThrowIfCancellationRequested();
                XlsxRowWriter row = sheet.StartRow(ct);
                writeRow(row, record);
                row.Dispose();
            }
            return ValueTask.CompletedTask;
        }
    }
}
