using System.Diagnostics.CodeAnalysis;

namespace ExcelReader.Core.Writer
{
    // High-level wrapper over the low-level ISheetWriter<TRow>/IRowWriter pair: maps each record to a
    // row via the caller-supplied delegate and drives the StartRow/dispose lifecycle. Generic over TRow
    // so callers keep the concrete row writer's full Write overloads (int, double, DateTime, ...).
    public static class SheetWriterExtensions
    {
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

        // SheetWriter/RowWriter-specific overload: row buffering only ever touches the destination
        // stream on the rare buffer-threshold flush (EndBufferedRow, already synchronous), so this
        // resolves in preference to the generic ISheetWriter<TRow> overload above whenever the caller's
        // static type is the concrete SheetWriter — e.g. plain IEnumerable<T> sources, which are the
        // overwhelming majority — skipping the per-row ValueTask/async-disposable machinery entirely.
        [SuppressMessage("Reliability", "CA1849:Call async methods when in an async method",
            Justification = "Deliberately using SheetWriter/RowWriter's synchronous fast path — see the comment above.")]
        [SuppressMessage("Sonar", "S6966:Await StartRowAsync instead",
            Justification = "Deliberately using SheetWriter's synchronous row fast path — see the comment above.")]
        [SuppressMessage("VisualStudio.Threading", "VSTHRD103:StartRow synchronously blocks",
            Justification = "Deliberately using SheetWriter's synchronous row fast path — see the comment above.")]
        [SuppressMessage("Sonar", "S6966:Await DisposeAsync instead",
            Justification = "Deliberately using RowWriter's synchronous Dispose fast path — see the comment above.")]
        [SuppressMessage("VisualStudio.Threading", "VSTHRD103:Dispose synchronously blocks",
            Justification = "Deliberately using RowWriter's synchronous Dispose fast path — see the comment above.")]
        public static ValueTask WriteRecordsAsync<T>(this SheetWriter sheet, IEnumerable<T> records,
                                                      Action<RowWriter, T> writeRow, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(records);
            ArgumentNullException.ThrowIfNull(writeRow);
            foreach (T record in records)
            {
                ct.ThrowIfCancellationRequested();
                RowWriter row = sheet.StartRow(ct);
                writeRow(row, record);
                row.Dispose();
            }
            return ValueTask.CompletedTask;
        }
    }
}
