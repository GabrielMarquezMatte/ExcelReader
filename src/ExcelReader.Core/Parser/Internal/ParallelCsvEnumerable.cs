using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Reader;
using Microsoft.Win32.SafeHandles;

namespace ExcelReader.Core.Parser.Internal
{
    // Runs a pool of chunk workers and merges their output in file order.
    //
    // Ordering is a correctness requirement, not a preference: a chunk's rows are only valid once its
    // predecessor confirms which boundary hypothesis was right.
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

        // The file handle this enumeration owns, or null for a memory/borrowed-stream source.
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

        // No [EnumeratorCancellation] here: that attribute only wires a token through on an iterator
        // returning IAsyncEnumerable<T>. This is the enumerator factory itself, so the parameter *is*
        // the token and is used directly (CS8424 fires if the attribute is applied anyway).
        [SuppressMessage("Usage", "VSTHRD003:Avoid awaiting foreign Tasks",
            Justification = "The awaited tasks are the worker tasks and per-chunk TaskCompletionSources created in this very method; every await is ConfigureAwait(false), so there is no captured context to deadlock against.")]
        [SuppressMessage("Reliability", "CA2025:Do not pass 'IDisposable' instances into unawaited tasks",
            Justification = "The CTS and semaphore handed to the workers outlive them by construction: the finally block cancels and then awaits every worker task to completion, and only the enclosing `using` declarations' finally — which runs after it — disposes them.")]
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP013:Await in using",
            Justification = "Task.Run is deliberately not awaited at creation: the workers must run concurrently with the merge loop. They are joined in the finally block before the `using` declarations dispose anything they touch.")]
        public async IAsyncEnumerator<T> GetAsyncEnumerator(CancellationToken cancellationToken = default)
        {
            using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            CancellationToken ct = cts.Token;

            // Chunks that may be claimed but not yet consumed by the merge. One per worker — deeper
            // buffering keeps more parsed models alive, pushing them out of Gen0.
            int inFlight = _dop;
            using var slots = new SemaphoreSlim(inFlight, inFlight);

            // At most `inFlight` chunks are claimed at once, handed out in index order, so the live
            // window is `inFlight` consecutive indices; one extra ring slot makes `index % ring`
            // collision-free across it, keeping these arrays O(dop) instead of O(chunk count).
            int ring = inFlight + 1;

            // Recycles chunk model lists instead of reallocating: a chunk's list is the largest
            // allocation the parallel path adds, and the merge drops it right after yielding its rows.
            var lists = new ListPool<T>(ring);
            var results = new CsvChunkResult<T>?[ring];
            var ready = new TaskCompletionSource<bool>[ring];
            foreach (ref TaskCompletionSource<bool> slot in ready.AsSpan())
            {
                slot = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            }

            // Typed as Func<Task>: `Task.Run(() => RunWorkerAsync(...))` would otherwise bind to the
            // Action overload, whose Task completes at the worker's first await rather than at its end.
            Func<Task> body = () => RunWorkerAsync(results, ready, slots, lists, ring, ct);
            Task[] workers = new Task[Math.Min(_dop, _plan.Count)];
            foreach (ref Task worker in workers.AsSpan())
            {
                worker = Task.Run(body, ct);
            }

            Task allWorkers = Task.WhenAll(workers);
            long rowsEmitted = 0;

            try
            {
                long confirmedNextStart = _firstDataRecordOffset;
                for (int i = 0; i < _plan.Count; i++)
                {
                    ct.ThrowIfCancellationRequested();
                    int slot = i % ring;
                    await WaitForChunkAsync(ready[slot].Task, allWorkers).ConfigureAwait(false);
                    CsvChunkResult<T> result = results[slot]!;

                    // The predecessor's ResolvedNextStart is ground truth. When this chunk guessed a
                    // different start, its rows are wrong and it is reparsed from the proven offset;
                    // a cascading correction into the next chunk falls out naturally since
                    // confirmedNextStart is recomputed from whatever result is actually emitted.
                    if (confirmedNextStart < long.MaxValue && result.ActualStart != confirmedNextStart)
                    {
                        // Reused verbatim: rows are wrong, but capacity was sized against this chunk.
                        List<T> reuse = result.Models;
                        reuse.Clear();
                        result = await CsvChunkWorker.ParseAsync(
                            _source, _plan[i], confirmedNextStart, _map, _info, _readerOptions, _config, reuse, ct).ConfigureAwait(false);
                    }

                    // Rearm this ring slot for chunk i + ring, then release — in that order, so the
                    // rearm happens before any worker can claim i + ring on the permit this releases.
                    results[slot] = null;
                    ready[slot] = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
                    slots.Release();

                    // Indexed, not a CollectionsMarshal.AsSpan foreach: a Span cannot live across yield.
                    List<T> models = result.Models;
                    int count = models.Count;
                    for (int m = 0; m < count; m++)
                    {
                        rowsEmitted++;
                        yield return models[m];
                    }

                    models.Clear();
                    lists.Return(models);

                    if (result.Failure is not null)
                    {
                        throw RenumberFailure(result.Failure, _config.HeaderRow + rowsEmitted + 1);
                    }

                    confirmedNextStart = result.ResolvedNextStart;
                    if (confirmedNextStart >= long.MaxValue)
                    {
                        break;
                    }
                }
            }
            finally
            {
                // Cancel, then join — a worker still writing into a pooled buffer when it's returned
                // to ArrayPool corrupts an unrelated consumer, so waiting for real completion matters.
                await cts.CancelAsync().ConfigureAwait(false);
                await allWorkers.ConfigureAwait(ConfigureAwaitOptions.SuppressThrowing);
                _ownedHandle?.Dispose();
            }
        }

        // One worker: pulls chunks until the plan is empty and publishes each result to the merge.
        private async Task RunWorkerAsync(
            CsvChunkResult<T>?[] results,
            TaskCompletionSource<bool>[] ready,
            SemaphoreSlim slots,
            ListPool<T> lists,
            int ring,
            CancellationToken ct)
        {
            int claimed = -1;
            try
            {
                // Slot acquired before the chunk is claimed, never the reverse: claiming first could
                // leave a worker holding the earliest unpublished chunk with no slot, descheduled while
                // later chunks drain the semaphore — a permanent hang on a chunk nobody will publish.
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
                        _source, chunk, confirmed, _map, _info, _readerOptions, _config, lists.Rent(), ct).ConfigureAwait(false);
                    results[chunk.Index % ring] = result;
                    ready[chunk.Index % ring].TrySetResult(true);
                }
            }
            // A dying worker owes the merge an answer for the chunk it held, or the merge waits
            // forever on a TaskCompletionSource nobody completes.
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

        // Backstop for every worker dying before claiming the chunk the merge is waiting on.
        [SuppressMessage("Usage", "VSTHRD003:Avoid awaiting foreign Tasks",
            Justification = "Both tasks are created by GetAsyncEnumerator, the sole caller, and passed in only to keep the waiting logic out of the iterator body.")]
        private static async Task WaitForChunkAsync(Task<bool> chunkReady, Task allWorkers)
        {
            Task finished = await Task.WhenAny(chunkReady, allWorkers).ConfigureAwait(false);
            if (finished == allWorkers)
            {
                await allWorkers.ConfigureAwait(false);
            }
            await chunkReady.ConfigureAwait(false);
        }

        // Only the merge knows a chunk's global row offset. The worker numbers rows from 1 within its
        // own chunk, so the exception is rebuilt here with the real row number (ExcelParseException is
        // immutable). An empty RawValue means missing-required-value rather than a parse failure.
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
