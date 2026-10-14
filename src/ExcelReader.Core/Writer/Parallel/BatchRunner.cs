using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Core.Writer.Parallel
{
    /// <summary>
    /// Cuts a record source into batches, formats at most a fixed number of them at once and hands the
    /// results to the caller in order. The source is enumerated, and results are emitted, only on the
    /// calling flow.
    /// </summary>
    internal sealed class BatchRunner<T>
    {
        private readonly Queue<Task<BatchSegment>> _inFlight = new();
        private readonly Func<T[], int, int, BatchSegment> _format;
        private readonly Func<BatchSegment, CancellationToken, ValueTask> _emit;
        private readonly CancellationToken _ct;
        private readonly int _degree;
        private readonly int _batchRows;
        private T[] _batch;
        private int _count;
        private int _rowsBefore;

        internal BatchRunner(ParallelWriteOptions options, int rowsBefore, Func<T[], int, int, BatchSegment> format,
            Func<BatchSegment, CancellationToken, ValueTask> emit, CancellationToken ct)
        {
            _degree = options.DegreeOfParallelism == 0 ? Environment.ProcessorCount : options.DegreeOfParallelism;
            _batchRows = Math.Min(options.BatchRows, ExcelLimits.MaxRows);
            _rowsBefore = rowsBefore;
            _format = format;
            _emit = emit;
            _ct = ct;
            _batch = ArrayPool<T>.Shared.Rent(_batchRows);
        }

        internal async ValueTask RunAsync(RecordSource<T> source)
        {
            try
            {
                if (source.Sync is not null)
                {
                    foreach (T record in source.Sync)
                    {
                        await AddAsync(record).ConfigureAwait(false);
                    }
                }
                else
                {
                    await foreach (T record in source.Async!.WithCancellation(_ct).ConfigureAwait(false))
                    {
                        await AddAsync(record).ConfigureAwait(false);
                    }
                }
                await CompleteAsync().ConfigureAwait(false);
            }
            catch
            {
                await AbandonAsync().ConfigureAwait(false);
                throw;
            }
        }

        private ValueTask AddAsync(T record)
        {
            _batch[_count++] = record;
            if (_count == _batchRows)
            {
                return DispatchAsync();
            }
            return ValueTask.CompletedTask;
        }

        private async ValueTask CompleteAsync()
        {
            if (_count > 0)
            {
                await DispatchAsync().ConfigureAwait(false);
            }
            while (_inFlight.Count > 0)
            {
                await EmitOldestAsync().ConfigureAwait(false);
            }
            ReturnBatch(_batch);
            _batch = [];
        }

        private async ValueTask DispatchAsync()
        {
            _ct.ThrowIfCancellationRequested();
            if (_rowsBefore + (long)_count > ExcelLimits.MaxRows)
            {
                ExcelLimits.ThrowRowLimit(_rowsBefore + (long)_count);
            }
            if (_inFlight.Count == _degree)
            {
                await EmitOldestAsync().ConfigureAwait(false);
            }
            T[] batch = _batch;
            int count = _count;
            int rowsBefore = _rowsBefore;
            _batch = ArrayPool<T>.Shared.Rent(_batchRows);
            _count = 0;
            _rowsBefore += count;
            if (_degree == 1)
            {
                _inFlight.Enqueue(Task.FromResult(Format(batch, count, rowsBefore)));
            }
            else
            {
                _inFlight.Enqueue(Task.Run(() => Format(batch, count, rowsBefore), CancellationToken.None));
            }
        }

        private BatchSegment Format(T[] batch, int count, int rowsBefore)
        {
            try
            {
                return _format(batch, count, rowsBefore);
            }
            finally
            {
                ReturnBatch(batch);
            }
        }

        private async ValueTask EmitOldestAsync()
        {
            using BatchSegment segment = await _inFlight.Dequeue().ConfigureAwait(false);
            await _emit(segment, _ct).ConfigureAwait(false);
        }

        // After a failure: no worker may outlive the call, and every finished batch gives its buffer back.
        [SuppressMessage("Design", "CA1031:Do not catch general exception types",
            Justification = "The caller is already rethrowing the first failure; a later batch's failure must not replace it.")]
        private async ValueTask AbandonAsync()
        {
            while (_inFlight.Count > 0)
            {
                try
                {
                    (await _inFlight.Dequeue().ConfigureAwait(false)).Dispose();
                }
                catch (Exception)
                {
                }
            }
            ReturnBatch(_batch);
            _batch = [];
        }

        private static void ReturnBatch(T[] batch)
        {
            if (batch.Length > 0)
            {
                ArrayPool<T>.Shared.Return(batch, RuntimeHelpers.IsReferenceOrContainsReferences<T>());
            }
        }
    }
}
