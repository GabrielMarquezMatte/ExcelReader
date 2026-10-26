using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;
using System.Threading.Channels;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Parser.ParallelCsv;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Parser.SheetPipeline
{
    internal static class SheetPipelineRunner
    {
        private readonly record struct Work<TState>(SheetChunk Chunk, bool First, TaskCompletionSource<ChunkOutcome<TState>> Done);

        internal static async Task<TState> AggregateAsync<TState>(
            SheetPipelineSource source, RowAggregation<TState> aggregation, RowActionFactory<TState>? factory,
            int headerRow, int dop, int chunkBytes, CancellationToken ct)
        {
            TState merged = default!;
            bool any = false;
            await foreach (TState state in RunAsync(source, aggregation, factory, headerRow, dop, chunkBytes, ct).ConfigureAwait(false))
            {
                merged = any ? aggregation.Combine(merged, state) : state;
                any = true;
            }
            return any ? merged : aggregation.Seed();
        }

        internal static async IAsyncEnumerable<TState> RunAsync<TState>(
            SheetPipelineSource source, RowAggregation<TState> aggregation, RowActionFactory<TState>? factory,
            int headerRow, int dop, int chunkBytes, [EnumeratorCancellation] CancellationToken cancellationToken)
        {
            source.Lifetime.Acquire(source.Owner);
            try
            {
                RowAction<TState>? bound = factory is not null && headerRow == 0 ? factory(default, sequential: false) : null;
                PipelineStart? start = source.Begin(headerRow, factory is null || headerRow == 0 ? null : header => bound = factory(header, sequential: false));
                if (start is null)
                {
                    yield break;
                }
                if (bound is not null)
                {
                    aggregation = new RowAggregation<TState> { Seed = aggregation.Seed, Accumulate = bound, Combine = aggregation.Combine };
                }

                using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                CancellationToken ct = cts.Token;
                int workers = Math.Max(1, dop - 1);
                var work = Channel.CreateBounded<Work<TState>>(workers * 2);
                var ordered = Channel.CreateBounded<Task<ChunkOutcome<TState>>>(
                    new BoundedChannelOptions(workers * 2) { SingleReader = true, SingleWriter = true });

                var tasks = new Task[workers + 1];
                tasks[0] = Task.Run(() => ProduceAsync(start, chunkBytes, source.MaxCellBytes, work.Writer, ordered.Writer, ct), CancellationToken.None);
                for (int w = 1; w <= workers; w++)
                {
                    tasks[w] = Task.Run(() => WorkAsync(source, aggregation, work.Reader, ct), CancellationToken.None);
                }

                long rows = 0;
                bool recover = false;
                try
                {
                    await foreach (Task<ChunkOutcome<TState>> pending in ordered.Reader.ReadAllAsync(ct).ConfigureAwait(false))
                    {
                        ChunkOutcome<TState> outcome = await pending.WaitAsync(ct).ConfigureAwait(false);
                        if (outcome.Failure is not null)
                        {
                            ThrowRenumbered(outcome.Failure, headerRow + rows, outcome.Rows);
                        }
                        rows += outcome.Rows;
                        yield return outcome.State;
                        if (!outcome.Clean)
                        {
                            recover = true;
                            break;
                        }
                        if (outcome.ReachedEnd)
                        {
                            break;
                        }
                    }
                    cancellationToken.ThrowIfCancellationRequested();
                }
                finally
                {
                    await cts.CancelAsync().ConfigureAwait(false);
                    await Task.WhenAll(tasks).ConfigureAwait(ConfigureAwaitOptions.SuppressThrowing);
                    while (work.Reader.TryRead(out Work<TState> unparsed))
                    {
                        ArrayPool<byte>.Shared.Return(unparsed.Chunk.Buffer);
                    }
                }

                if (recover)
                {
                    ChunkOutcome<TState> tail = source.ParseTail(headerRow, rows, aggregation, cancellationToken);
                    if (tail.Failure is not null)
                    {
                        ThrowRenumbered(tail.Failure, headerRow + rows, tail.Rows);
                    }
                    yield return tail.State;
                }
            }
            finally
            {
                source.Lifetime.Release();
            }
        }

        [SuppressMessage("Design", "CA1031:Do not catch general exception types",
            Justification = "The exception is recorded and delivered to the merge through the ordered channel's completion.")]
        private static async Task ProduceAsync<TState>(
            PipelineStart start, int chunkBytes, int maxCellBytes, ChannelWriter<Work<TState>> work,
            ChannelWriter<Task<ChunkOutcome<TState>>> ordered, CancellationToken ct)
        {
            Exception? failure = null;
            try
            {
                bool first = true;
                await ChunkProducer.RunAsync(
                    start.Stream, start.Cutter, chunkBytes, maxCellBytes,
                    async chunk =>
                    {
                        var done = new TaskCompletionSource<ChunkOutcome<TState>>(TaskCreationOptions.RunContinuationsAsynchronously);
                        await ordered.WriteAsync(done.Task, ct).ConfigureAwait(false);
                        await work.WriteAsync(new Work<TState>(chunk, first, done), ct).ConfigureAwait(false);
                        first = false;
                    },
                    ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (ct.IsCancellationRequested)
            {
            }
            catch (Exception ex)
            {
                failure = ex;
            }
            finally
            {
                await start.Stream.DisposeAsync().ConfigureAwait(false);
                work.Complete();
                ordered.Complete(failure);
            }
        }

        private static async Task WorkAsync<TState>(
            SheetPipelineSource source, RowAggregation<TState> aggregation, ChannelReader<Work<TState>> work, CancellationToken ct)
        {
            await foreach (Work<TState> item in work.ReadAllAsync(ct).ConfigureAwait(false))
            {
                try
                {
                    item.Done.SetResult(source.ParseChunk(item.Chunk.Buffer.AsMemory(0, item.Chunk.Length), item.First, aggregation, ct));
                }
                catch (Exception ex)
                {
                    item.Done.TrySetException(ex);
                    _ = item.Done.Task.Exception;
                    throw;
                }
                finally
                {
                    ArrayPool<byte>.Shared.Return(item.Chunk.Buffer);
                }
            }
        }

        // A chunk numbers its rows from 1 when it parses into models; a callback or accumulator run leaves Row
        // at 0, and then the chunk's count of rows read, which includes the failing one, places it.
        private static void ThrowRenumbered(ExceptionDispatchInfo failure, long rowsBefore, long rowsInChunk)
        {
            if (failure.SourceException is ExcelParseException parse)
            {
                throw ProjectionRules.AtRow(parse, rowsBefore + (parse.Row > 0 ? parse.Row : rowsInChunk));
            }
            failure.Throw();
        }
    }
}
