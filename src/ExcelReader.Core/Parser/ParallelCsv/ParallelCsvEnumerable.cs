using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader.Csv;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Core.Parser.ParallelCsv
{
    internal sealed class ParallelCsvEnumerable<T> : IAsyncEnumerable<T>
    {
        private readonly CsvChunkSource _source;
        private readonly CsvChunkPlan _plan;
        private readonly long _firstDataRecordOffset;
        private readonly CsvBoundColumnMap<T> _map;
        private readonly TypeMapInfo<T> _info;
        private readonly CsvReaderOptions _readerOptions;
        private readonly ExcelParserOptions _config;
        private readonly int _dop;

        private readonly ReaderLifetime? _lease;
        private readonly object? _owner;

        internal ParallelCsvEnumerable(
            CsvChunkSource source,
            CsvChunkPlan plan,
            long firstDataRecordOffset,
            CsvBoundColumnMap<T> map,
            TypeMapInfo<T> info,
            CsvReaderOptions readerOptions,
            ExcelParserOptions config,
            int degreeOfParallelism,
            ReaderLifetime? lease = null,
            object? owner = null)
        {
            _source = source;
            _plan = plan;
            _firstDataRecordOffset = firstDataRecordOffset;
            _map = map;
            _info = info;
            _readerOptions = readerOptions;
            _config = config;
            _dop = degreeOfParallelism;
            _lease = lease;
            _owner = owner;
        }

        [SuppressMessage("Reliability", "CA2025:Do not pass 'IDisposable' instances into unawaited tasks",
            Justification = "The CTS and semaphore handed to the workers outlive them by construction: the finally block cancels and then awaits every worker task to completion, and only the enclosing `using` declarations' finally — which runs after it — disposes them.")]
        public async IAsyncEnumerator<T> GetAsyncEnumerator(CancellationToken cancellationToken = default)
        {
            _lease?.Acquire(_owner!);
            using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            CancellationToken ct = cts.Token;

            int inFlight = _dop;
            using var slots = new SemaphoreSlim(inFlight, inFlight);

            int ring = inFlight + 1;

            var lists = new ConcurrentBag<List<T>>();
            MergeState state = StartWorkers(slots, lists, ring, ct);
            Task allWorkers = Task.WhenAll(state.Workers);
            long recordsBefore = 0;

            try
            {
                long confirmedNextStart = _firstDataRecordOffset;
                for (int i = 0; i < _plan.Count; i++)
                {
                    ct.ThrowIfCancellationRequested();
                    CsvChunkResult<T> result = await TakeChunkAsync(
                        state, slots, i, ring, confirmedNextStart, allWorkers, ct).ConfigureAwait(false);

                    List<T> models = result.Models;
                    int count = models.Count;
                    for (int m = 0; m < count; m++)
                    {
                        yield return models[m];
                    }

                    models.Clear();
                    lists.Add(models);

                    confirmedNextStart = AdvanceOrThrow(result, recordsBefore);
                    recordsBefore += result.Records;
                    if (confirmedNextStart >= long.MaxValue)
                    {
                        break;
                    }
                }
            }
            finally
            {
                await cts.CancelAsync().ConfigureAwait(false);
                await allWorkers.ConfigureAwait(ConfigureAwaitOptions.SuppressThrowing);
                _lease?.Release();
            }
        }

        private long AdvanceOrThrow(CsvChunkResult<T> result, long recordsBefore)
        {
            if (result.Failure is null)
            {
                return result.ResolvedNextStart;
            }
            if (result.Failure.SourceException is ExcelParseException parse)
            {
                int headerRows = _info.IsIndexBased ? 0 : _config.HeaderRow;
                throw ProjectionRules.AtRow(parse, headerRows + recordsBefore + result.Records + 1);
            }
            result.Failure.Throw();
            return long.MaxValue;
        }

        private readonly record struct MergeState(
            CsvChunkResult<T>?[] Results,
            TaskCompletionSource<bool>[] Ready,
            Task[] Workers,
            CsvChunkPlan Plan);

        private MergeState StartWorkers(SemaphoreSlim slots, ConcurrentBag<List<T>> lists, int ring, CancellationToken ct)
        {
            var results = new CsvChunkResult<T>?[ring];
            var ready = new TaskCompletionSource<bool>[ring];
            foreach (ref TaskCompletionSource<bool> slot in ready.AsSpan())
            {
                slot = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            }

            // Each enumeration hands out the chunks again; a plan already drained would leave the merge waiting forever.
            CsvChunkPlan plan = _plan.Restart();
            Task[] workers = new Task[Math.Min(_dop, plan.Count)];
            foreach (ref Task worker in workers.AsSpan())
            {
                worker = Task.Run(body, ct);
            }
            return new MergeState(results, ready, workers, plan);
            Task body()
            {
                return RunWorkerAsync(plan, results, ready, slots, lists, ring, ct);
            }
        }

        private async Task<CsvChunkResult<T>> TakeChunkAsync(
            MergeState state,
            SemaphoreSlim slots,
            int i,
            int ring,
            long confirmedNextStart,
            Task allWorkers,
            CancellationToken ct)
        {
            int slot = i % ring;
            await WaitForChunkAsync(state.Ready[slot].Task, allWorkers).ConfigureAwait(false);
            CsvChunkResult<T> result = state.Results[slot]!;

            if (confirmedNextStart < long.MaxValue && result.ActualStart != confirmedNextStart)
            {
                List<T> reuse = result.Models;
                reuse.Clear();
                result = await CsvChunkWorker.ParseAsync(
                    _source, state.Plan[i], confirmedNextStart, _map, _info, _readerOptions, _config, reuse, contentCache: null, ct).ConfigureAwait(false);
            }

            state.Results[slot] = null;
            state.Ready[slot] = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            slots.Release();
            return result;
        }

        private async Task RunWorkerAsync(
            CsvChunkPlan plan,
            CsvChunkResult<T>?[] results,
            TaskCompletionSource<bool>[] ready,
            SemaphoreSlim slots,
            ConcurrentBag<List<T>> lists,
            int ring,
            CancellationToken ct)
        {
            int claimed = -1;
            // One per worker: a cache per chunk would be allocated, and start cold, every 64 KB of input.
            Utf8StringCache? contentCache = _readerOptions.InternStrings ? new Utf8StringCache() : null;
            try
            {
                while (true)
                {
                    ct.ThrowIfCancellationRequested();
                    await slots.WaitAsync(ct).ConfigureAwait(false);
                    if (!plan.TryTakeNext(out CsvChunk chunk))
                    {
                        slots.Release();
                        return;
                    }
                    claimed = chunk.Index;
                    long? confirmed = chunk.Index == 0 ? _firstDataRecordOffset : null;
                    CsvChunkResult<T> result = await CsvChunkWorker.ParseAsync(
                        _source, chunk, confirmed, _map, _info, _readerOptions, _config, lists.TryTake(out List<T>? pooled) ? pooled : [], contentCache, ct).ConfigureAwait(false);
                    results[chunk.Index % ring] = result;
                    ready[chunk.Index % ring].TrySetResult(true);
                }
            }
            catch (OperationCanceledException) when (claimed >= 0)
            {
                ready[claimed % ring].TrySetCanceled(ct);
                throw;
            }
            catch (Exception ex) when (claimed >= 0)
            {
                ready[claimed % ring].TrySetException(ex);
                throw;
            }
        }

        private static async Task WaitForChunkAsync(Task<bool> chunkReady, Task allWorkers)
        {
            Task finished = await Task.WhenAny(chunkReady, allWorkers).ConfigureAwait(false);
            if (finished == allWorkers)
            {
                await allWorkers.ConfigureAwait(false);
            }
            await chunkReady.ConfigureAwait(false);
        }
    }
}
