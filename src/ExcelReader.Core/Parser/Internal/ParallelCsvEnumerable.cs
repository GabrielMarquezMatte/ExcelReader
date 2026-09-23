using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Reader;
using Microsoft.Win32.SafeHandles;

namespace ExcelReader.Core.Parser.Internal
{
    internal sealed class ParallelCsvEnumerable<T> : IAsyncEnumerable<T>
    {
        private readonly CsvChunkSource _source;
        private readonly CsvChunkPlan _plan;
        private readonly long _firstDataRecordOffset;
        private readonly CsvBoundColumnMap<T> _map;
        private readonly TypeMapInfo<T> _info;
        private readonly CsvReaderOptions _readerOptions;
        private readonly ExcelParserConfig _config;
        private readonly int _dop;

        private readonly SafeFileHandle? _ownedHandle;

        internal ParallelCsvEnumerable(
            CsvChunkSource source,
            CsvChunkPlan plan,
            long firstDataRecordOffset,
            CsvBoundColumnMap<T> map,
            TypeMapInfo<T> info,
            CsvReaderOptions readerOptions,
            ExcelParserConfig config,
            int degreeOfParallelism,
            SafeFileHandle? ownedHandle)
        {
            _source = source;
            _plan = plan;
            _firstDataRecordOffset = firstDataRecordOffset;
            _map = map;
            _info = info;
            _readerOptions = readerOptions;
            _config = config;
            _dop = degreeOfParallelism;
            _ownedHandle = ownedHandle;
        }

        [SuppressMessage("Reliability", "CA2025:Do not pass 'IDisposable' instances into unawaited tasks",
            Justification = "The CTS and semaphore handed to the workers outlive them by construction: the finally block cancels and then awaits every worker task to completion, and only the enclosing `using` declarations' finally — which runs after it — disposes them.")]
        public async IAsyncEnumerator<T> GetAsyncEnumerator(CancellationToken cancellationToken = default)
        {
            using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            CancellationToken ct = cts.Token;

            int inFlight = _dop;
            using var slots = new SemaphoreSlim(inFlight, inFlight);

            int ring = inFlight + 1;

            var lists = new ConcurrentBag<List<T>>();
            MergeState state = StartWorkers(slots, lists, ring, ct);
            Task allWorkers = Task.WhenAll(state.Workers);
            long rowsEmitted = 0;

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
                        rowsEmitted++;
                        yield return models[m];
                    }

                    models.Clear();
                    lists.Add(models);

                    confirmedNextStart = AdvanceOrThrow(result, rowsEmitted);
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
                _ownedHandle?.Dispose();
            }
        }

        private long AdvanceOrThrow(CsvChunkResult<T> result, long rowsEmitted)
        {
            if (result.Failure is not null)
            {
                throw RenumberFailure(result.Failure, _config.HeaderRow + rowsEmitted + 1);
            }
            return result.ResolvedNextStart;
        }

        private readonly record struct MergeState(
            CsvChunkResult<T>?[] Results,
            TaskCompletionSource<bool>[] Ready,
            Task[] Workers);

        private MergeState StartWorkers(SemaphoreSlim slots, ConcurrentBag<List<T>> lists, int ring, CancellationToken ct)
        {
            var results = new CsvChunkResult<T>?[ring];
            var ready = new TaskCompletionSource<bool>[ring];
            foreach (ref TaskCompletionSource<bool> slot in ready.AsSpan())
            {
                slot = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            }

            Task[] workers = new Task[Math.Min(_dop, _plan.Count)];
            foreach (ref Task worker in workers.AsSpan())
            {
                worker = Task.Run(body, ct);
            }
            return new MergeState(results, ready, workers);
            Task body()
            {
                return RunWorkerAsync(results, ready, slots, lists, ring, ct);
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
                    _source, _plan[i], confirmedNextStart, _map, _info, _readerOptions, _config, reuse, ct).ConfigureAwait(false);
            }

            state.Results[slot] = null;
            state.Ready[slot] = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            slots.Release();
            return result;
        }

        private async Task RunWorkerAsync(
            CsvChunkResult<T>?[] results,
            TaskCompletionSource<bool>[] ready,
            SemaphoreSlim slots,
            ConcurrentBag<List<T>> lists,
            int ring,
            CancellationToken ct)
        {
            int claimed = -1;
            try
            {
                while (true)
                {
                    ct.ThrowIfCancellationRequested();
                    await slots.WaitAsync(ct).ConfigureAwait(false);
                    if (!_plan.TryTakeNext(out CsvChunk chunk))
                    {
                        slots.Release();
                        return;
                    }
                    claimed = chunk.Index;
                    long? confirmed = chunk.Index == 0 ? _firstDataRecordOffset : null;
                    CsvChunkResult<T> result = await CsvChunkWorker.ParseAsync(
                        _source, chunk, confirmed, _map, _info, _readerOptions, _config, lists.TryTake(out List<T>? pooled) ? pooled : [], ct).ConfigureAwait(false);
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

        private static ExcelParseException RenumberFailure(ExcelParseException original, long globalRow)
        {
            if (original.RawValue.Length == 0)
            {
                return ProjectionRules.MissingRequiredValue(original.ColumnName, (int)globalRow);
            }
            return new ExcelParseException((int)globalRow, original.ColumnName, original.RawValue);
        }
    }
}
