using System.Runtime.ExceptionServices;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;

namespace ExcelReader.Core.Parser.ParallelCsv
{
    internal static class ParallelCsvProcessor
    {
        private const int ChunksPerWorker = 4;
        private const int MinChunkBytes = 1024 * 1024;
        private const int MaxChunkBytes = 64 * 1024 * 1024;
        private const int RowsPerCancellationCheck = 4096;

        private sealed class ChunkOutcome<TState>(long actualStart, TState state, FailureGate? gate)
        {
            public readonly long ActualStart = actualStart;
            public readonly FailureGate? Gate = gate;
            public long ResolvedNextStart = long.MaxValue;
            public TState State = state;
            public long Delivered;
            public bool Abandoned;
            public ExceptionDispatchInfo? Failure;
        }

        // The lowest index of a chunk that failed. Chunks past it stop: the merge throws at that failure before
        // reaching them, unless the failing chunk was misguessed and is re-read clean, and then it reads them itself.
        private sealed class FailureGate
        {
            private int _first = int.MaxValue;

            internal void Fail(int index)
            {
                int seen = Volatile.Read(ref _first);
                while (index < seen)
                {
                    int prior = Interlocked.CompareExchange(ref _first, index, seen);
                    if (prior == seen)
                    {
                        return;
                    }
                    seen = prior;
                }
            }

            internal bool Past(int index)
            {
                return index > Volatile.Read(ref _first);
            }
        }

        internal static Task<TState> RunWithChunkSizeAsync<TState>(
            ReadOnlyMemory<byte> data, RowAggregation<TState> aggregation, RowActionFactory<TState>? factory,
            CsvReaderOptions reader, int headerRow, int degreeOfParallelism, int chunkSize, CancellationToken ct)
        {
            CsvReaderOptions resolved = CsvDialectResolver.Resolve(data, reader);
            return PartitionedAsync(new CsvChunkSource(data), aggregation, factory, resolved, headerRow, ParallelCsvFactory.Normalize(degreeOfParallelism), chunkSize, ct);
        }

        internal static Task<TState> RunPartitionedAsync<TState>(
            CsvChunkSource source, RowAggregation<TState> aggregation, CsvReaderOptions reader, int headerRow, int degreeOfParallelism,
            int chunkSizeOverride, CancellationToken ct)
        {
            return PartitionedAsync(source, aggregation, null, reader, headerRow, ParallelCsvFactory.Normalize(degreeOfParallelism), chunkSizeOverride, ct);
        }

        internal static async Task<TState> RunAsync<TState>(
            CsvReader reader, RowAggregation<TState> aggregation, RowActionFactory<TState>? factory, SheetParallelOptions options, CancellationToken ct)
        {
            reader.Lifetime.Acquire(reader);
            try
            {
                int dop = ParallelCsvFactory.Normalize(options.DegreeOfParallelism);
                if (!reader.TryGetChunkSource(out CsvChunkSource source) || !ParallelCsvFactory.CanPartition(dop, source.Length, reader.Options))
                {
                    return await SequentialAsync(reader, aggregation, factory, options.HeaderRow, ct).ConfigureAwait(false);
                }
                return await PartitionedAsync(source, aggregation, factory, reader.Options, options.HeaderRow, dop, chunkSizeOverride: 0, ct).ConfigureAwait(false);
            }
            finally
            {
                reader.Lifetime.Release();
            }
        }

        private static RowAggregation<TState> WithAccumulate<TState>(RowAggregation<TState> aggregation, RowAction<TState> accumulate)
        {
            return new RowAggregation<TState> { Seed = aggregation.Seed, Accumulate = accumulate, Combine = aggregation.Combine };
        }

        private static async Task<TState> SequentialAsync<TState>(
            CsvReader reader, RowAggregation<TState> aggregation, RowActionFactory<TState>? factory, int headerRow, CancellationToken ct)
        {
            CsvReader.Enumerator rows = reader.OpenSheet(ct);
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

        private static async Task<TState> PartitionedAsync<TState>(
            CsvChunkSource source, RowAggregation<TState> aggregation, RowActionFactory<TState>? factory, CsvReaderOptions reader, int headerRow, int dop,
            int chunkSizeOverride, CancellationToken ct)
        {
            long dataStart = 0;
            if (headerRow > 0)
            {
                (dataStart, RowAction<TState>? bound) = await RecordStartAfterAsync(source, reader, headerRow, factory, ct).ConfigureAwait(false);
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

            var outcomes = new ChunkOutcome<TState>?[plan.Count];
            var gate = new FailureGate();
            var workers = new Task[Math.Min(dop, plan.Count)];
            for (int w = 0; w < workers.Length; w++)
            {
                workers[w] = Task.Run(() => DrainPlanAsync(source, plan, dataStart, reader, aggregation, outcomes, gate, ct), ct);
            }
            await Task.WhenAll(workers).ConfigureAwait(false);

            TState merged = default!;
            long confirmedStart = dataStart;
            long recordsBefore = headerRow;
            for (int i = 0; i < plan.Count && confirmedStart < long.MaxValue; i++)
            {
                ChunkOutcome<TState>? outcome = outcomes[i];
                if (outcome is null || outcome.Abandoned || outcome.ActualStart != confirmedStart)
                {
                    outcome = await ParseChunkAsync(source, plan[i], confirmedStart, reader, aggregation, gate: null, ct).ConfigureAwait(false);
                }
                if (outcome.Failure is not null)
                {
                    // A partitioned run parses without row numbers; the failing record is the chunk's last delivered one.
                    if (outcome.Failure.SourceException is ExcelParseException { Row: 0 } parse)
                    {
                        throw ProjectionRules.AtRow(parse, recordsBefore + outcome.Delivered);
                    }
                    outcome.Failure.Throw();
                }
                merged = i == 0 ? outcome.State : aggregation.Combine(merged, outcome.State);
                confirmedStart = outcome.ResolvedNextStart;
                recordsBefore += outcome.Delivered;
            }
            return merged;
        }

        private static async Task DrainPlanAsync<TState>(
            CsvChunkSource source, CsvChunkPlan plan, long dataStart, CsvReaderOptions reader,
            RowAggregation<TState> aggregation, ChunkOutcome<TState>?[] outcomes, FailureGate gate, CancellationToken ct)
        {
            while (plan.TryTakeNext(out CsvChunk chunk))
            {
                ct.ThrowIfCancellationRequested();
                if (gate.Past(chunk.Index))
                {
                    continue;
                }
                long start = chunk.Index == 0 ? dataStart : CsvChunkWorker.GuessStart(source, chunk, reader.Quote, reader.Delimiter);
                outcomes[chunk.Index] = await ParseChunkAsync(source, chunk, start, reader, aggregation, gate, ct).ConfigureAwait(false);
            }
        }

        private static ValueTask<ChunkOutcome<TState>> ParseChunkAsync<TState>(
            CsvChunkSource source, CsvChunk chunk, long start, CsvReaderOptions reader,
            RowAggregation<TState> aggregation, FailureGate? gate, CancellationToken ct)
        {
            var outcome = new ChunkOutcome<TState>(start, aggregation.Seed(), gate);
            if (start >= source.Length)
            {
                return new ValueTask<ChunkOutcome<TState>>(outcome);
            }
            CsvReaderOptions chunkOptions = reader with { DetectEncodingFromByteOrderMark = reader.DetectEncodingFromByteOrderMark && start == 0 };
            if (!source.IsMemory)
            {
                return ParseFileChunkAsync(source, chunk, outcome, chunkOptions, aggregation.Accumulate, ct);
            }
            ParseMemoryChunk(source, chunk, outcome, chunkOptions, aggregation.Accumulate, ct);
            return new ValueTask<ChunkOutcome<TState>>(outcome);
        }

        private static void ParseMemoryChunk<TState>(
            CsvChunkSource source, CsvChunk chunk, ChunkOutcome<TState> outcome, CsvReaderOptions chunkOptions,
            RowAction<TState> accumulate, CancellationToken ct)
        {
            var rows = new CsvReader.Enumerator(source.SliceAt(outcome.ActualStart), chunkOptions, contentCache: null, ct);
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
                Fail(outcome, chunk, ex);
            }
            finally
            {
                rows.Dispose();
            }
        }

        private static async ValueTask<ChunkOutcome<TState>> ParseFileChunkAsync<TState>(
            CsvChunkSource source, CsvChunk chunk, ChunkOutcome<TState> outcome, CsvReaderOptions chunkOptions,
            RowAction<TState> accumulate, CancellationToken ct)
        {
            Stream partition = source.OpenAt(outcome.ActualStart);
            await using (partition.ConfigureAwait(false))
            {
                var rows = new CsvReader.Enumerator(partition, chunkOptions, contentCache: null, ct);
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
                        Fail(outcome, chunk, ex);
                    }
                }
            }
            return outcome;
        }

        private static void Fail<TState>(ChunkOutcome<TState> outcome, CsvChunk chunk, Exception ex)
        {
            outcome.Failure = ExceptionDispatchInfo.Capture(ex);
            outcome.Gate?.Fail(chunk.Index);
        }

        private static bool DeliverWithinChunk<TState>(
            CsvReader.Enumerator rows, CsvChunk chunk, ChunkOutcome<TState> outcome, RowAction<TState> accumulate, CancellationToken ct)
        {
            long absolute = outcome.ActualStart + rows.CurrentRecordStart;
            if (absolute >= chunk.End)
            {
                outcome.ResolvedNextStart = absolute;
                return false;
            }
            if (outcome.Gate is not null && outcome.Gate.Past(chunk.Index))
            {
                outcome.Abandoned = true;
                return false;
            }
            if (++outcome.Delivered % RowsPerCancellationCheck == 0)
            {
                ct.ThrowIfCancellationRequested();
            }
            accumulate(ref outcome.State, rows.Current);
            return true;
        }

        private static void Deliver<TState>(RowAction<TState> accumulate, ref TState state, CsvReader.Enumerator rows)
        {
            accumulate(ref state, rows.Current);
        }

        private static async Task<(long Start, RowAction<TState>? Accumulate)> RecordStartAfterAsync<TState>(
            CsvChunkSource source, CsvReaderOptions reader, int records, RowActionFactory<TState>? factory, CancellationToken ct)
        {
            RowAction<TState>? bound = null;
            CsvReader csv = source.OpenReader(reader);
            await using (csv.ConfigureAwait(false))
            {
                CsvReader.Enumerator rows = csv.OpenSheet(ct);
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
