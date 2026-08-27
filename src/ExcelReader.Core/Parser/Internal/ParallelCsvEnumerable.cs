using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Parser.Internal
{
    // Runs a pool of chunk workers and merges their output in file order.
    //
    // Ordering is not a preference here: a chunk's rows are only *valid* once its predecessor has
    // parsed up to the chunk's start and confirmed which boundary hypothesis was right. That
    // sequencing is a correctness requirement, and the ordered merge is where it happens.
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

        internal ParallelCsvEnumerable(
            CsvChunkSource source,
            CsvChunkPlan plan,
            long firstDataRecordOffset,
            CsvBoundColumnMap<T> map,
            TypeMapInfo<T> info,
            CsvReaderOptions readerOptions,
            ExcelParserConfig config,
            int degreeOfParallelism)
        {
            _source = source;
            _plan = plan;
            _firstDataRecordOffset = firstDataRecordOffset;
            _map = map;
            _info = info;
            _readerOptions = readerOptions;
            _config = config;
            _dop = degreeOfParallelism;
        }

        // No [EnumeratorCancellation] here: that attribute only wires a token through on an iterator
        // returning IAsyncEnumerable<T>. This is the enumerator factory itself, so the parameter *is*
        // the token and is used directly (CS8424 fires if the attribute is applied anyway).
        [SuppressMessage("Performance", "HLQ006:GetAsyncEnumerator should return a value type",
            Justification = "IAsyncEnumerable<T> mandates the interface return type, and the compiler-generated async iterator is a reference type by construction.")]
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

            // Published-but-unconsumed results. Strictly below the chunk count (4 * dop) so it
            // actually binds, and at least dop so workers are not throttled while the consumer keeps
            // up. Without it, workers run ahead and the buffer grows to the whole file.
            using var slots = new SemaphoreSlim(2 * _dop, 2 * _dop);
            var results = new CsvChunkResult<T>?[_plan.Count];
            var ready = new TaskCompletionSource<bool>[_plan.Count];
            foreach (ref TaskCompletionSource<bool> slot in ready.AsSpan())
            {
                slot = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            }

            // Typed as Func<Task> deliberately: `Task.Run(() => RunWorkerAsync(...))` would also bind
            // to the Action overload, which returns a Task that completes at the worker's first await
            // instead of at its end — and the whole join-before-release guarantee below rests on
            // `allWorkers` meaning the workers have actually finished.
            Func<Task> body = () => RunWorkerAsync(results, ready, slots, ct);
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
                    await WaitForChunkAsync(ready[i].Task, allWorkers).ConfigureAwait(false);
                    CsvChunkResult<T> result = results[i]!;

                    // The predecessor's ResolvedNextStart is ground truth. When this chunk guessed a
                    // different start, its rows are wrong and it is reparsed from the proven offset.
                    // A reparse changes THIS chunk's ResolvedNextStart too, so the correction can
                    // cascade into the next chunk — which the loop handles naturally, because
                    // confirmedNextStart is recomputed from whatever result we end up emitting.
                    if (confirmedNextStart < long.MaxValue && result.ActualStart != confirmedNextStart)
                    {
                        result = await CsvChunkWorker.ParseAsync(
                            _source, _plan[i], confirmedNextStart, _map, _info, _readerOptions, _config, ct).ConfigureAwait(false);
                    }

                    // Drop the array's reference before yielding: the semaphore bounds how many chunks
                    // are *in flight*, but only this keeps already-merged chunks from pinning every
                    // parsed row in the array for the whole enumeration.
                    results[i] = null;
                    slots.Release();

                    // Indexed rather than foreach: HLQ012 would have this iterate a
                    // CollectionsMarshal.AsSpan over the list, and a Span cannot live across the
                    // `yield return` an ordered merge is built on.
                    List<T> models = result.Models;
                    for (int m = 0; m < models.Count; m++)
                    {
                        rowsEmitted++;
                        yield return models[m];
                    }

                    if (result.Failure is not null)
                    {
                        // The failing record is the one after this chunk's last emitted model (the
                        // worker stops at the failure, so Models holds exactly FailureRowInChunk
                        // items), shifted past the rows 1..HeaderRow the sequential path counts
                        // before the first data record.
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
                // Cancel, then JOIN. A pooled buffer handed back to ArrayPool while a worker is still
                // writing into it corrupts an unrelated consumer elsewhere in the process, so waiting
                // for real worker completion is not optional here. SuppressThrowing keeps the join
                // from replacing the merge's own exception (or from throwing out of DisposeAsync when
                // the consumer simply broke early) while still marking worker faults observed.
                //
                // `cts` and `slots` are disposed by their `using` declarations, whose generated
                // finally block encloses this one — so nothing a worker touches is released until
                // after this await has returned.
                await cts.CancelAsync().ConfigureAwait(false);
                await allWorkers.ConfigureAwait(ConfigureAwaitOptions.SuppressThrowing);
            }
        }

        // One worker: pulls chunks until the plan is empty and publishes each result to the merge.
        private async Task RunWorkerAsync(
            CsvChunkResult<T>?[] results,
            TaskCompletionSource<bool>[] ready,
            SemaphoreSlim slots,
            CancellationToken ct)
        {
            int inFlight = -1;
            try
            {
                // Acquire the slot BEFORE claiming the chunk, never the other way round. Claiming
                // first opens a window in which a worker owns the earliest unpublished chunk while
                // holding no slot: if it is descheduled there and later-indexed chunks drain the
                // semaphore, it blocks in WaitAsync forever and the merge waits forever on a chunk
                // nobody will publish — a silent, permanent hang.
                //
                // With this order the invariant the merge depends on actually holds. If the merge is
                // blocked on chunk i, then either i is claimed — and its owner already holds a slot,
                // so it is parsing and will publish — or i is unclaimed, which (chunks being handed
                // out in index order) means no chunk >= i is claimed either, so every permit is held
                // by a chunk < i, all of which the merge has already released. A worker therefore
                // acquires immediately and claims i.
                while (true)
                {
                    ct.ThrowIfCancellationRequested();
                    await slots.WaitAsync(ct).ConfigureAwait(false);
                    if (!_plan.TryTakeNext(out CsvChunk chunk))
                    {
                        // No chunk to spend this permit on, and no merge iteration will ever release
                        // it on this worker's behalf, so hand it back before leaving.
                        slots.Release();
                        return;
                    }
                    inFlight = chunk.Index;
                    long? confirmed = chunk.Index == 0 ? _firstDataRecordOffset : null;
                    CsvChunkResult<T> result = await CsvChunkWorker.ParseAsync(
                        _source, chunk, confirmed, _map, _info, _readerOptions, _config, ct).ConfigureAwait(false);
                    results[chunk.Index] = result;
                    ready[chunk.Index].TrySetResult(true);
                }
            }
            // A worker that dies owes the merge an answer for the chunk it was holding. Without this
            // the merge would wait forever on a TaskCompletionSource nobody will ever complete: the
            // only other signal, Task.WhenAll, cannot complete while the surviving workers are still
            // running. TrySetCanceled rather than TrySetException for cancellation keeps the
            // early-break path from parking an exception nobody will observe.
            catch (OperationCanceledException) when (inFlight >= 0)
            {
                ready[inFlight].TrySetCanceled(ct);
                throw;
            }
            catch (Exception ex) when (inFlight >= 0)
            {
                ready[inFlight].TrySetException(ex);
                throw;
            }
        }

        // Surfaces a worker's infrastructure failure (I/O, OOM) instead of deadlocking on a chunk
        // whose TaskCompletionSource will never be completed. This is the backstop for the case where
        // every worker died before claiming the chunk the merge is waiting for; the common case is the
        // per-chunk propagation in RunWorkerAsync.
        [SuppressMessage("Usage", "VSTHRD003:Avoid awaiting foreign Tasks",
            Justification = "Both tasks are created by GetAsyncEnumerator, the sole caller, and passed in only to keep the waiting logic out of the iterator body.")]
        private static async Task WaitForChunkAsync(Task<bool> chunkReady, Task allWorkers)
        {
            Task finished = await Task.WhenAny(chunkReady, allWorkers).ConfigureAwait(false);
            if (finished == allWorkers)
            {
                // Rethrows the first worker exception as-is. No AggregateException wrapping: callers
                // already have catch blocks written against the sequential path's exception types.
                await allWorkers.ConfigureAwait(false);
            }
            await chunkReady.ConfigureAwait(false);
        }

        // Only the merge knows a chunk's global row offset, so only the merge can raise a parse
        // failure with the row number the sequential path would have reported. The chunk worker's
        // projector numbers rows from 1 within its own chunk, so the original number is discarded and
        // the exception rebuilt — ExcelParseException is immutable.
        //
        // An empty RawValue distinguishes the two failure shapes: a parse failure is only raised for a
        // non-empty cell (see ExcelParseException's own remarks), so a blank raw value is the
        // missing-required-value case, which carries a different message.
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
