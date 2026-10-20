using System.Collections.Concurrent;

namespace ExcelReader.Core.Writer.Parallel
{
    /// <summary>Drives a parallel record write for any sheet writer that implements <see cref="IBatchSheet{TSelf, TRow}"/>.</summary>
    internal static class ParallelRecordWriter
    {
        internal static async ValueTask WriteRecordsAsync<TSheet, TRow, T>(TSheet sheet, RecordSource<T> records,
            ExcelRecordLayout<T> layout, ParallelWriteOptions? options, CancellationToken ct)
            where TSheet : class, ISheetWriter<TRow>, IBatchSheet<TSheet, TRow>
            where TRow : IRowWriter
        {
            ParallelWriteOptions effective = Resolve(options);
            (string[] headers, Action<TRow, T> writeRow) = layout.Columns<TRow>();
            await SheetWriterExtensions.WriteHeaderAsync(sheet, headers, ct).ConfigureAwait(false);
            await RunAsync(sheet, records, writeRow, effective, ct).ConfigureAwait(false);
        }

        internal static ValueTask WriteRowsAsync<TSheet, TRow, T>(TSheet sheet, RecordSource<T> records,
            Action<TRow, T> writeRow, ParallelWriteOptions? options, CancellationToken ct)
            where TSheet : class, ISheetWriter<TRow>, IBatchSheet<TSheet, TRow>
            where TRow : IRowWriter
        {
            return RunAsync(sheet, records, writeRow, Resolve(options), ct);
        }

        internal static async ValueTask WriteTableAsync<TSheet, TRow, T>(TSheet sheet, RecordSource<T> records, string name,
            ExcelRecordLayout<T> layout, ExcelTableOptions? tableOptions, ParallelWriteOptions? options, CancellationToken ct)
            where TSheet : class, ITableSheetWriter<TRow>, IBatchSheet<TSheet, TRow>
            where TRow : IRowWriter
        {
            ParallelWriteOptions effective = Resolve(options);
            (string[] headers, Action<TRow, T> writeRow) = layout.Columns<TRow>();
            await sheet.BeginTableAsync(name, headers, tableOptions, ct).ConfigureAwait(false);
            await RunAsync(sheet, records, SheetWriterExtensions.ShiftToTable(writeRow, tableOptions), effective, ct).ConfigureAwait(false);
            sheet.EndTable();
        }

        private static ParallelWriteOptions Resolve(ParallelWriteOptions? options)
        {
            ParallelWriteOptions effective = options ?? ParallelWriteOptions.Default;
            effective.Validate();
            return effective;
        }

        private static async ValueTask RunAsync<TSheet, TRow, T>(TSheet sheet, RecordSource<T> records,
            Action<TRow, T> writeRow, ParallelWriteOptions options, CancellationToken ct)
            where TSheet : class, ISheetWriter<TRow>, IBatchSheet<TSheet, TRow>
            where TRow : IRowWriter
        {
            if (options.EffectiveDegree == 1)
            {
                await RunSeriallyAsync(sheet, records, writeRow, ct).ConfigureAwait(false);
                return;
            }
            int rowsBefore = await sheet.BeginSegmentsAsync(ct).ConfigureAwait(false);
            var workers = new ConcurrentBag<TSheet>();
            try
            {
                var runner = new BatchRunner<T>(options, rowsBefore,
                    (batch, count, before) => FormatBatch(sheet, workers, writeRow, batch, count, before),
                    sheet.AppendSegmentAsync, ct);
                await runner.RunAsync(records).ConfigureAwait(false);
            }
            catch
            {
                await sheet.FailSegmentsAsync().ConfigureAwait(false);
                throw;
            }
            finally
            {
                foreach (TSheet worker in workers)
                {
                    worker.ReleaseDetached();
                }
            }
            sheet.EndSegments();
        }

        // Batching only pays when batches overlap; one worker would add the cutting and sync flushes for nothing.
        private static async ValueTask RunSeriallyAsync<TSheet, TRow, T>(TSheet sheet, RecordSource<T> records,
            Action<TRow, T> writeRow, CancellationToken ct)
            where TSheet : class, ISheetWriter<TRow>, IBatchSheet<TSheet, TRow>
            where TRow : IRowWriter
        {
            try
            {
                if (records.Sync is not null)
                {
                    await sheet.WriteRecordsAsync(records.Sync, writeRow, ct).ConfigureAwait(false);
                    return;
                }
                await sheet.WriteRecordsAsync(records.Async!, writeRow, ct).ConfigureAwait(false);
            }
            catch
            {
                await sheet.FailSegmentsAsync().ConfigureAwait(false);
                throw;
            }
        }

        private static BatchSegment FormatBatch<TSheet, TRow, T>(TSheet sheet, ConcurrentBag<TSheet> workers,
            Action<TRow, T> writeRow, T[] batch, int count, int rowsBefore)
            where TSheet : class, IBatchSheet<TSheet, TRow>
            where TRow : IRowWriter
        {
            if (!workers.TryTake(out TSheet? worker))
            {
                worker = sheet.CreateDetached(count);
            }
            try
            {
                worker.BeginBatch(rowsBefore);
                for (int i = 0; i < count; i++)
                {
                    TRow row = worker.StartBatchRow();
                    writeRow(row, batch[i]);
                    row.Dispose();
                }
                return Compress<TSheet, TRow>(sheet, worker, count);
            }
            finally
            {
                workers.Add(worker);
            }
        }

        private static BatchSegment Compress<TSheet, TRow>(TSheet sheet, TSheet worker, int rows)
            where TSheet : class, IBatchSheet<TSheet, TRow>
            where TRow : IRowWriter
        {
            var segment = new BatchSegment(rows);
            try
            {
                segment.Fill(worker.EndBatch(segment), sheet.BatchCompression);
                return segment;
            }
            catch
            {
                segment.Dispose();
                throw;
            }
        }
    }
}
