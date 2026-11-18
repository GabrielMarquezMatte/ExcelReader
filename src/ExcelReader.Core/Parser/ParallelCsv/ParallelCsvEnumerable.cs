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
        private readonly CancellationToken _ct;

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
            object? owner = null,
            CancellationToken ct = default)
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
            _ct = ct;
        }

        public IAsyncEnumerator<T> GetAsyncEnumerator(CancellationToken cancellationToken = default)
        {
            return new Enumerator(this, cancellationToken);
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
            ColumnStringCaches? contentCache = _readerOptions.InternStrings ? new ColumnStringCaches() : null;
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

        private sealed class Enumerator : IAsyncEnumerator<T>
        {
            private readonly ParallelCsvEnumerable<T> _parent;
            private readonly CancellationToken _cancellationToken;
            private readonly ConcurrentBag<List<T>> _lists = new();
            private CancellationTokenSource? _cts;
            private SemaphoreSlim? _slots;
            private CancellationToken _ct;
            private MergeState _state;
            private Task _allWorkers = Task.CompletedTask;
            private int _ring;
            private int _next;
            private long _confirmedNextStart;
            private long _recordsBefore;
            private CsvChunkResult<T>? _chunk;
            private List<T>? _list;
            private int _index;
            private int _count;
            private bool _done;

            internal Enumerator(ParallelCsvEnumerable<T> parent, CancellationToken cancellationToken)
            {
                _parent = parent;
                _cancellationToken = cancellationToken;
                _confirmedNextStart = parent._firstDataRecordOffset;
            }

            public T Current { get; private set; } = default!;

            public ValueTask<bool> MoveNextAsync()
            {
                if (_list is not null && _index < _count)
                {
                    Current = _list[_index++];
                    return new ValueTask<bool>(true);
                }
                return NextChunkAsync();
            }

            private async ValueTask<bool> NextChunkAsync()
            {
                if (_done)
                {
                    return false;
                }
                try
                {
                    if (_cts is null)
                    {
                        Start();
                    }
                    while (true)
                    {
                        if ((_chunk is not null && !FinishChunk()) || _next >= _parent._plan.Count)
                        {
                            await DisposeAsync().ConfigureAwait(false);
                            return false;
                        }
                        _ct.ThrowIfCancellationRequested();
                        CsvChunkResult<T> result = await _parent.TakeChunkAsync(
                            _state, _slots!, _next, _ring, _confirmedNextStart, _allWorkers, _ct).ConfigureAwait(false);
                        _next++;
                        _chunk = result;
                        _list = result.Models;
                        _count = _list.Count;
                        _index = 0;
                        if (_count > 0)
                        {
                            Current = _list[_index++];
                            return true;
                        }
                    }
                }
                catch
                {
                    await DisposeAsync().ConfigureAwait(false);
                    throw;
                }
            }

            [SuppressMessage("Reliability", "CA2025:Do not pass 'IDisposable' instances into unawaited tasks",
                Justification = "The CTS and semaphore handed to the workers outlive them by construction: DisposeAsync cancels and then awaits every worker task to completion before disposing them.")]
            private void Start()
            {
                _parent._lease?.Acquire(_parent._owner!);
                _cts = CancellationTokenSource.CreateLinkedTokenSource(_cancellationToken, _parent._ct);
                _ct = _cts.Token;
                _slots = new SemaphoreSlim(_parent._dop, _parent._dop);
                _ring = _parent._dop + 1;
                _state = _parent.StartWorkers(_slots, _lists, _ring, _ct);
                _allWorkers = Task.WhenAll(_state.Workers);
            }

            // The chunk's list goes back to the workers' pool only once the consumer has moved past its last row.
            private bool FinishChunk()
            {
                CsvChunkResult<T> chunk = _chunk!;
                _chunk = null;
                _list = null;
                List<T> models = chunk.Models;
                models.Clear();
                _lists.Add(models);
                _confirmedNextStart = _parent.AdvanceOrThrow(chunk, _recordsBefore);
                _recordsBefore += chunk.Records;
                return _confirmedNextStart < long.MaxValue;
            }

            public async ValueTask DisposeAsync()
            {
                if (_done)
                {
                    return;
                }
                _done = true;
                _list = null;
                _chunk = null;
                if (_cts is null)
                {
                    return;
                }
                await _cts.CancelAsync().ConfigureAwait(false);
                await _allWorkers.ConfigureAwait(ConfigureAwaitOptions.SuppressThrowing);
                _parent._lease?.Release();
                _slots?.Dispose();
                _cts.Dispose();
            }
        }
    }
}
