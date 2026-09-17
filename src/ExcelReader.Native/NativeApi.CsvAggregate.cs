using System.Buffers;
using System.Collections.Concurrent;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Native
{
    internal static unsafe partial class NativeApi
    {
        private sealed class CsvAggregateContext(NativeCsvAggregationRaw raw) : IDisposable
        {
            private int _failure;

            internal readonly delegate* unmanaged<void**, void*, int> Seed =
                (delegate* unmanaged<void**, void*, int>)raw.Seed;
            internal readonly delegate* unmanaged<void*, NativeRow*, void*, int> Accumulate =
                (delegate* unmanaged<void*, NativeRow*, void*, int>)raw.Accumulate;
            internal readonly delegate* unmanaged<void*, void*, void*, int> Combine =
                (delegate* unmanaged<void*, void*, void*, int>)raw.Combine;
            internal readonly delegate* unmanaged<void*, void*, void> FreeState =
                (delegate* unmanaged<void*, void*, void>)raw.FreeState;
            internal readonly void* UserData = (void*)raw.UserData;
            internal readonly ConcurrentBag<CsvAggregateState> Seeded = [];
            internal readonly CancellationTokenSource Cancellation = new();

            internal int Failure => Volatile.Read(ref _failure);

            internal void Abort(int code)
            {
                Interlocked.CompareExchange(ref _failure, code, 0);
                Cancellation.Cancel();
                throw new OperationCanceledException(Cancellation.Token);
            }

            public void Dispose()
            {
                Cancellation.Dispose();
            }
        }

        private sealed class PointerMemoryManager(byte* pointer, int length) : MemoryManager<byte>
        {
            public override Span<byte> GetSpan()
            {
                return new(pointer, length);
            }

            public override MemoryHandle Pin(int elementIndex = 0)
            {
                return new(pointer + elementIndex);
            }

            public override void Unpin()
            {
            }

            protected override void Dispose(bool disposing)
            {
            }
        }

        internal static int AggregateCsvFile(
            ReadOnlySpan<byte> utf8Path, NativeCsvAggregationRaw aggregation,
            NativeCsvParallelOptionsRaw? rawOptions, out nint result)
        {
            result = 0;
            int status = NativeCsvAggregateOptions.Translate(rawOptions, out CsvParallelOptions options);
            if (status != NativeStatus.Ok)
            {
                return status;
            }

            string path = Encoding.UTF8.GetString(utf8Path);
            CsvAggregateContext context = new(aggregation);
            return Run(
                context,
                Task.Run(() => Excel.AggregateCsvParallelAsync(path, Build(context), options, context.Cancellation.Token)),
                out result);
        }

        internal static int AggregateCsvMemory(
            byte* data, int dataLength, NativeCsvAggregationRaw aggregation,
            NativeCsvParallelOptionsRaw? rawOptions, out nint result)
        {
            result = 0;
            int status = NativeCsvAggregateOptions.Translate(rawOptions, out CsvParallelOptions options);
            if (status != NativeStatus.Ok)
            {
                return status;
            }

            using PointerMemoryManager manager = new(data, dataLength);
            CsvAggregateContext context = new(aggregation);
            return Run(
                context,
                Task.Run(() => Excel.AggregateCsvParallelAsync(manager.Memory, Build(context), options, context.Cancellation.Token)),
                out result);
        }

        private static CsvAggregation<CsvAggregateState> Build(CsvAggregateContext context)
        {
            return new CsvAggregation<CsvAggregateState>
            {
                Seed = () =>
                {
                    void* native = null;
                    CsvAggregateState state = new();
                    context.Seeded.Add(state);
                    int status = context.Seed(&native, context.UserData);
                    state.Native = (nint)native;
                    if (status != NativeStatus.Ok)
                    {
                        context.Abort(status);
                    }
                    return state;
                },
                Accumulate = (ref CsvAggregateState state, Row row) =>
                {
                    NativeRow native = state.WriteRow(row);
                    int status = context.Accumulate((void*)state.Native, &native, context.UserData);
                    if (status != NativeStatus.Ok)
                    {
                        context.Abort(status);
                    }
                },
                Combine = (accumulator, next) =>
                {
                    int status = context.Combine(
                        (void*)accumulator.Native, (void*)next.Native, context.UserData);
                    if (status != NativeStatus.Ok)
                    {
                        context.Abort(status);
                    }
                    return accumulator;
                },
            };
        }

        private static int Run(CsvAggregateContext context, Task<CsvAggregateState> pending, out nint result)
        {
            result = 0;
            ClearLastError();
            CsvAggregateState? winner = null;
            try
            {
                winner = pending.GetAwaiter().GetResult();
                result = winner.Native;
                return NativeStatus.Ok;
            }
            catch (OperationCanceledException)
            {
                int failure = context.Failure;
                return failure != 0 ? failure : NativeStatus.Error;
            }
            catch (Exception exception)
            {
                int failure = context.Failure;
                if (failure != 0)
                {
                    return failure;
                }
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
            finally
            {
                Sweep(context, winner);
            }
        }

        private static void Sweep(CsvAggregateContext context, CsvAggregateState? winner)
        {
            foreach (CsvAggregateState state in context.Seeded)
            {
                if (!ReferenceEquals(state, winner) && state.Native != 0)
                {
                    context.FreeState((void*)state.Native, context.UserData);
                }
                state.ReleaseScratch();
            }
            context.Dispose();
        }
    }
}
