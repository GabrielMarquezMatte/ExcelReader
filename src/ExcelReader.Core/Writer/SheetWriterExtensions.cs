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
    }
}
