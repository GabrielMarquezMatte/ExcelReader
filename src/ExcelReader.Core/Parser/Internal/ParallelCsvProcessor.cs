using System.Runtime.ExceptionServices;
using ExcelReader.Core.Reader;
using Microsoft.Win32.SafeHandles;

namespace ExcelReader.Core.Parser.Internal
{
    internal static class ParallelCsvProcessor
    {
        private const int ChunksPerWorker = 4;
        private const int MinChunkBytes = 1024 * 1024;
        private const int MaxChunkBytes = 64 * 1024 * 1024;
        private const int RowsPerCancellationCheck = 4096;

        private sealed class ChunkOutcome<TState>(long actualStart, TState state)
        {
            public readonly long ActualStart = actualStart;
            public long ResolvedNextStart = long.MaxValue;
            public TState State = state;
            public long Delivered;
            public ExceptionDispatchInfo? Failure;
        }

        internal static async Task<TState> RunAsync<TState>(
            string path, CsvAggregation<TState> aggregation, CsvAccumulateFactory<TState>? factory, CsvParallelOptions options, CancellationToken ct)
        {
            options = WithResolvedDialect(options, CsvDialectResolver.Resolve(path, options.Reader));
            int dop = ParallelCsvFactory.Normalize(options.DegreeOfParallelism);
            long length = new FileInfo(path).Length;
            if (!ParallelCsvFactory.CanPartition(dop, length, options.Reader))
            {
                CsvReader reader = await Excel.FromCsvFileAsync(path, options.Reader, ct).ConfigureAwait(false);
                return await SequentialAsync(reader, aggregation, factory, options.HeaderRow, ct).ConfigureAwait(false);
            }
            using SafeFileHandle handle = File.OpenHandle(path, FileMode.Open, FileAccess.Read, FileShare.Read, FileOptions.Asynchronous | FileOptions.RandomAccess);
            return await PartitionedAsync(new CsvChunkSource(handle, length), aggregation, factory, options, dop, chunkSizeOverride: 0, ct).ConfigureAwait(false);
        }

        internal static Task<TState> RunAsync<TState>(
            ReadOnlyMemory<byte> data, CsvAggregation<TState> aggregation, CsvAccumulateFactory<TState>? factory, CsvParallelOptions options, CancellationToken ct)
        {
            options = WithResolvedDialect(options, CsvDialectResolver.Resolve(data, options.Reader));
            int dop = ParallelCsvFactory.Normalize(options.DegreeOfParallelism);
            if (!ParallelCsvFactory.CanPartition(dop, data.Length, options.Reader))
            {
                return SequentialAsync(Excel.FromCsv(data, options.Reader), aggregation, factory, options.HeaderRow, ct);
            }
            return PartitionedAsync(new CsvChunkSource(data), aggregation, factory, options, dop, chunkSizeOverride: 0, ct);
        }

        internal static async Task<TState> RunAsync<TState>(
            Stream stream, CsvAggregation<TState> aggregation, CsvAccumulateFactory<TState>? factory, CsvParallelOptions options, CancellationToken ct)
        {
            options = WithResolvedDialect(options, await CsvDialectResolver.ResolveAsync(stream, options.Reader, ct).ConfigureAwait(false));
            int dop = ParallelCsvFactory.Normalize(options.DegreeOfParallelism);
            if (!CsvSourceResolver.TryResolve(stream, out CsvChunkSource source)
                || !ParallelCsvFactory.CanPartition(dop, source.Length, options.Reader))
            {
                CsvReader reader = await Excel.FromCsvAsync(stream, leaveOpen: true, options.Reader, ct).ConfigureAwait(false);
                return await SequentialAsync(reader, aggregation, factory, options.HeaderRow, ct).ConfigureAwait(false);
            }
            return await PartitionedAsync(source, aggregation, factory, options, dop, chunkSizeOverride: 0, ct).ConfigureAwait(false);
        }

        internal static Task<TState> RunWithChunkSizeAsync<TState>(
            ReadOnlyMemory<byte> data, CsvAggregation<TState> aggregation, CsvAccumulateFactory<TState>? factory, CsvParallelOptions options, int chunkSize, CancellationToken ct)
        {
            options = WithResolvedDialect(options, CsvDialectResolver.Resolve(data, options.Reader));
            int dop = ParallelCsvFactory.Normalize(options.DegreeOfParallelism);
            return PartitionedAsync(new CsvChunkSource(data), aggregation, factory, options, dop, chunkSize, ct);
        }

        private static CsvParallelOptions WithResolvedDialect(CsvParallelOptions options, CsvReaderOptions reader)
        {
            if (ReferenceEquals(reader, options.Reader))
            {
                return options;
            }
            return new CsvParallelOptions
            {
                DegreeOfParallelism = options.DegreeOfParallelism,
                Reader = reader,
                HeaderRow = options.HeaderRow,
            };
        }

        private static CsvAggregation<TState> WithAccumulate<TState>(CsvAggregation<TState> aggregation, CsvRowAction<TState> accumulate)
        {
            return new CsvAggregation<TState> { Seed = aggregation.Seed, Accumulate = accumulate, Combine = aggregation.Combine };
        }

        private static async Task<TState> SequentialAsync<TState>(
            CsvReader reader, CsvAggregation<TState> aggregation, CsvAccumulateFactory<TState>? factory, int headerRow, CancellationToken ct)
        {
            await using (reader.ConfigureAwait(false))
            {
                CsvReader.Enumerator rows = reader.GetAsyncEnumerator(ct);
                await using (rows.ConfigureAwait(false))
                {
                    TState state = aggregation.Seed();
                    if (factory is not null && headerRow == 0)
                    {
                        aggregation = WithAccumulate(aggregation, factory(default, sequential: true));
                    }
                    for (int skipped = 0; skipped < headerRow; skipped++)
                    {
                        if (!await rows.MoveNextAsync().ConfigureAwait(false))
                        {
                            return state;
                        }
                        if (factory is not null && skipped == headerRow - 1)
                        {
                            aggregation = WithAccumulate(aggregation, factory(rows.Current, sequential: true));
                        }
                    }
                    long delivered = 0;
                    while (await rows.MoveNextAsync().ConfigureAwait(false))
                    {
                        if (++delivered % RowsPerCancellationCheck == 0)
                        {
                            ct.ThrowIfCancellationRequested();
                        }
                        Deliver(aggregation.Accumulate, ref state, rows);
                    }
                    return state;
                }
            }
        }

        private static async Task<TState> PartitionedAsync<TState>(
            CsvChunkSource source, CsvAggregation<TState> aggregation, CsvAccumulateFactory<TState>? factory, CsvParallelOptions options, int dop, int chunkSizeOverride,
            CancellationToken ct)
        {
            CsvReaderOptions reader = options.Reader;
            long dataStart = 0;
            if (options.HeaderRow > 0)
            {
                (dataStart, CsvRowAction<TState>? bound) = await RecordStartAfterAsync(source, reader, options.HeaderRow, factory, ct).ConfigureAwait(false);
                if (bound is not null)
                {
                    aggregation = WithAccumulate(aggregation, bound);
                }
            }
            else if (factory is not null)
            {
                aggregation = WithAccumulate(aggregation, factory(default, sequential: false));
            }
            if (dataStart >= source.Length)
            {
                return aggregation.Seed();
            }

            long dataLength = source.Length - dataStart;
            int chunkSize = chunkSizeOverride > 0
                ? chunkSizeOverride
                : (int)Math.Clamp(dataLength / ((long)dop * ChunksPerWorker), MinChunkBytes, MaxChunkBytes);
            CsvChunkPlan plan = CsvChunkPlan.Create(dataStart, dataLength, dop, chunkSize);

            var outcomes = new ChunkOutcome<TState>[plan.Count];
            var workers = new Task[Math.Min(dop, plan.Count)];
            for (int w = 0; w < workers.Length; w++)
            {
                workers[w] = Task.Run(() => DrainPlanAsync(source, plan, dataStart, reader, aggregation, outcomes, ct), ct);
            }
            await Task.WhenAll(workers).ConfigureAwait(false);

            TState merged = default!;
            long confirmedStart = dataStart;
            for (int i = 0; i < plan.Count && confirmedStart < long.MaxValue; i++)
            {
                ChunkOutcome<TState> outcome = outcomes[i];
                if (outcome.ActualStart != confirmedStart)
                {
                    outcome = await ParseChunkAsync(source, plan[i], confirmedStart, reader, aggregation, ct).ConfigureAwait(false);
                }
                outcome.Failure?.Throw();
                merged = i == 0 ? outcome.State : aggregation.Combine(merged, outcome.State);
                confirmedStart = outcome.ResolvedNextStart;
            }
            return merged;
        }

        private static async Task DrainPlanAsync<TState>(
            CsvChunkSource source, CsvChunkPlan plan, long dataStart, CsvReaderOptions reader,
            CsvAggregation<TState> aggregation, ChunkOutcome<TState>[] outcomes, CancellationToken ct)
        {
            while (plan.TryTakeNext(out CsvChunk chunk))
            {
                ct.ThrowIfCancellationRequested();
                long start = chunk.Index == 0 ? dataStart : CsvChunkWorker.GuessStart(source, chunk, reader.Quote);
                outcomes[chunk.Index] = await ParseChunkAsync(source, chunk, start, reader, aggregation, ct).ConfigureAwait(false);
            }
        }

        private static ValueTask<ChunkOutcome<TState>> ParseChunkAsync<TState>(
            CsvChunkSource source, CsvChunk chunk, long start, CsvReaderOptions reader,
            CsvAggregation<TState> aggregation, CancellationToken ct)
        {
            var outcome = new ChunkOutcome<TState>(start, aggregation.Seed());
            if (start >= source.Length)
            {
                return new ValueTask<ChunkOutcome<TState>>(outcome);
            }
            CsvReaderOptions chunkOptions = reader with { DetectEncodingFromByteOrderMark = start == 0 };
            if (!source.IsMemory)
            {
                return ParseFileChunkAsync(source, chunk, outcome, chunkOptions, aggregation.Accumulate, ct);
            }
            ParseMemoryChunk(source, chunk, outcome, chunkOptions, aggregation.Accumulate, ct);
            return new ValueTask<ChunkOutcome<TState>>(outcome);
        }

        private static void ParseMemoryChunk<TState>(
            CsvChunkSource source, CsvChunk chunk, ChunkOutcome<TState> outcome, CsvReaderOptions chunkOptions,
            CsvRowAction<TState> accumulate, CancellationToken ct)
        {
            var rows = new CsvReader.Enumerator(source.SliceAt(outcome.ActualStart), chunkOptions, ct);
            try
            {
                while (rows.MoveNext())
                {
                    if (!DeliverWithinChunk(rows, chunk, outcome, accumulate, ct))
                    {
                        break;
                    }
                }
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                outcome.Failure = ExceptionDispatchInfo.Capture(ex);
            }
            finally
            {
                rows.Dispose();
            }
        }

        private static async ValueTask<ChunkOutcome<TState>> ParseFileChunkAsync<TState>(
            CsvChunkSource source, CsvChunk chunk, ChunkOutcome<TState> outcome, CsvReaderOptions chunkOptions,
            CsvRowAction<TState> accumulate, CancellationToken ct)
        {
            Stream partition = source.OpenAt(outcome.ActualStart);
            await using (partition.ConfigureAwait(false))
            {
                var rows = new CsvReader.Enumerator(partition, chunkOptions, ct);
                await using (rows.ConfigureAwait(false))
                {
                    try
                    {
                        while (await rows.MoveNextAsync().ConfigureAwait(false))
                        {
                            if (!DeliverWithinChunk(rows, chunk, outcome, accumulate, ct))
                            {
                                break;
                            }
                        }
                    }
                    catch (Exception ex) when (ex is not OperationCanceledException)
                    {
                        outcome.Failure = ExceptionDispatchInfo.Capture(ex);
                    }
                }
            }
            return outcome;
        }

        private static bool DeliverWithinChunk<TState>(
            CsvReader.Enumerator rows, CsvChunk chunk, ChunkOutcome<TState> outcome, CsvRowAction<TState> accumulate, CancellationToken ct)
        {
            long absolute = outcome.ActualStart + rows.CurrentRecordStart;
            if (absolute >= chunk.End)
            {
                outcome.ResolvedNextStart = absolute;
                return false;
            }
            if (++outcome.Delivered % RowsPerCancellationCheck == 0)
            {
                ct.ThrowIfCancellationRequested();
            }
            accumulate(ref outcome.State, rows.Current);
            return true;
        }

        private static void Deliver<TState>(CsvRowAction<TState> accumulate, ref TState state, CsvReader.Enumerator rows)
        {
            accumulate(ref state, rows.Current);
        }

        private static async Task<(long Start, CsvRowAction<TState>? Accumulate)> RecordStartAfterAsync<TState>(
            CsvChunkSource source, CsvReaderOptions reader, int records, CsvAccumulateFactory<TState>? factory, CancellationToken ct)
        {
            CsvRowAction<TState>? bound = null;
            CsvReader csv = source.OpenReader(reader);
            await using (csv.ConfigureAwait(false))
            {
                CsvReader.Enumerator rows = csv.GetAsyncEnumerator(ct);
                await using (rows.ConfigureAwait(false))
                {
                    for (int i = 0; i <= records; i++)
                    {
                        if (!await rows.MoveNextAsync().ConfigureAwait(false))
                        {
                            return (long.MaxValue, bound);
                        }
                        if (factory is not null && i == records - 1)
                        {
                            bound = factory(rows.Current, sequential: false);
                        }
                    }
                    return (rows.CurrentRecordStart, bound);
                }
            }
        }
    }
}
