using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Reader;
using Microsoft.Win32.SafeHandles;

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

        // The file handle this enumeration owns, or null for a memory/borrowed-stream source. Held
        // here rather than in a wrapping IAsyncEnumerable<T> so that closing it costs one `finally`
        // at the end of the enumeration instead of re-yielding every row through a second async
        // iterator — the merge is single-threaded, so per-row cost there is not something more
        // workers can absorb.
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
        [SuppressMessage("Performance", "HLQ006:GetAsyncEnumerator should return a value type",
            Justification = "IAsyncEnumerable<T> mandates the interface return type, and the compiler-generated async iterator is a reference type by construction.")]
        [SuppressMessage("Usage", "VSTHRD003:Avoid awaiting foreign Tasks",
            Justification = "The awaited tasks are the worker tasks and per-chunk TaskCompletionSources created in this very method; every await is ConfigureAwait(false), so there is no captured context to deadlock against.")]
        [SuppressMessage("Reliability", "CA2025:Do not pass 'IDisposable' instances into unawaited tasks",
            Justification = "The CTS and semaphore handed to the workers outlive them by construction: the finally block cancels and then awaits every worker task to completion, and only the enclosing `using` declarations' finally — which runs after it — disposes them.")]
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP013:Await in using",
            Justification = "Task.Run is deliberately not awaited at creation: the workers must run concurrently with the merge loop. They are joined in the finally block before the `using` declarations dispose anything they touch.")]
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "_ownedHandle is not borrowed: it is the SafeFileHandle ParallelCsvFactory.Create<T>(string, ...) opened for this one enumeration and handed over with it, and closing it when the consumer stops — including an early break — is exactly this enumerable's obligation. Every other source passes null here, since those handles belong to the caller.")]
        public async IAsyncEnumerator<T> GetAsyncEnumerator(CancellationToken cancellationToken = default)
        {
            using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            CancellationToken ct = cts.Token;

            // Chunks that may be claimed but not yet consumed by the merge. One per worker: with 64 KB
            // chunks (CsvChunkPlan) every worker still has a chunk to parse the moment the merge frees
            // a slot, and deeper buffering measurably *hurt* — 2*dop was the original value, and at
            // dop=16 it cost 10-45% wall clock on both corpora by keeping enough parsed models alive
            // to push them out of Gen0.
            int inFlight = _dop;
            using var slots = new SemaphoreSlim(inFlight, inFlight);

            // Ring size: at most `inFlight` chunks are claimed at once, and they are handed out in
            // index order, so the live window is [merge position, merge position + inFlight - 1] —
            // inFlight consecutive indices. One more slot than that makes index % Ring collision-free
            // across the whole window, so all three per-chunk arrays are O(dop) rather than O(chunks).
            // At 64 KB a 10 GB file has ~160k chunks; arrays indexed by chunk would be megabytes of
            // live bookkeeping, and 160k TaskCompletionSources held to the end of the enumeration.
            int ring = inFlight + 1;

            // Chunk model lists are recycled rather than reallocated. A chunk's list is the single
            // largest allocation the parallel path adds over the sequential one, and the merge drops
            // it the moment it has yielded its rows, so without recycling every chunk buys and burns
            // a fresh array.
            var lists = new ListPool<T>(ring);
            var results = new CsvChunkResult<T>?[ring];
            var ready = new TaskCompletionSource<bool>[ring];
            foreach (ref TaskCompletionSource<bool> slot in ready.AsSpan())
            {
                slot = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            }

            // Typed as Func<Task> deliberately: `Task.Run(() => RunWorkerAsync(...))` would also bind
            // to the Action overload, which returns a Task that completes at the worker's first await
            // instead of at its end — and the whole join-before-release guarantee below rests on
            // `allWorkers` meaning the workers have actually finished.
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
                    // different start, its rows are wrong and it is reparsed from the proven offset.
                    // A reparse changes THIS chunk's ResolvedNextStart too, so the correction can
                    // cascade into the next chunk — which the loop handles naturally, because
                    // confirmedNextStart is recomputed from whatever result we end up emitting.
                    if (confirmedNextStart < long.MaxValue && result.ActualStart != confirmedNextStart)
                    {
                        // The discarded result's list is reused verbatim for the reparse: its rows are
                        // wrong, but its capacity was sized against this very chunk.
                        List<T> reuse = result.Models;
                        reuse.Clear();
                        result = await CsvChunkWorker.ParseAsync(
                            _source, _plan[i], confirmedNextStart, _map, _info, _readerOptions, _config, reuse, ct).ConfigureAwait(false);
                    }

                    // Rearm this ring slot for chunk i + ring, then release. Strictly in that order:
                    // chunk i + ring can only be claimed on a permit released at chunk i + 1 or later,
                    // so rearming before this release puts it safely ahead of any worker that will
                    // read the slot, with the semaphore providing the ordering edge. Dropping the
                    // result reference here also matters on its own: the semaphore bounds how many
                    // chunks are in flight, but only this keeps a merged chunk's rows from staying
                    // reachable until the slot is overwritten a full ring later.
                    results[slot] = null;
                    ready[slot] = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
                    slots.Release();

                    // Indexed rather than foreach: HLQ012 would have this iterate a
                    // CollectionsMarshal.AsSpan over the list, and a Span cannot live across the
                    // `yield return` an ordered merge is built on.
                    List<T> models = result.Models;
                    int count = models.Count;
                    for (int m = 0; m < count; m++)
                    {
                        rowsEmitted++;
                        yield return models[m];
                    }

                    // Recycled only once its rows are out the door. Clear() drops this chunk's model
                    // references immediately, so recycling never keeps a parsed row alive longer than
                    // not recycling would.
                    models.Clear();
                    lists.Return(models);

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

                // After the join, never before: a worker still running would be reading through this
                // very handle. Closing here (rather than at finalization) is what makes an early
                // `break` by the consumer release the file promptly.
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
                    claimed = chunk.Index;
                    long? confirmed = chunk.Index == 0 ? _firstDataRecordOffset : null;
                    CsvChunkResult<T> result = await CsvChunkWorker.ParseAsync(
                        _source, chunk, confirmed, _map, _info, _readerOptions, _config, lists.Rent(), ct).ConfigureAwait(false);
                    results[chunk.Index % ring] = result;
                    ready[chunk.Index % ring].TrySetResult(true);
                }
            }
            // A worker that dies owes the merge an answer for the chunk it was holding. Without this
            // the merge would wait forever on a TaskCompletionSource nobody will ever complete: the
            // only other signal, Task.WhenAll, cannot complete while the surviving workers are still
            // running. TrySetCanceled rather than TrySetException for cancellation keeps the
            // early-break path from parking an exception nobody will observe.
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
