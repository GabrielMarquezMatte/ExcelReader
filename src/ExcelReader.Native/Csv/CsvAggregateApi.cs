using System.Buffers;
using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using ExcelReader.Native.Reading;

namespace ExcelReader.Native.Csv
{
    internal static unsafe partial class CsvAggregateApi
    {
        [SuppressMessage("Design", "CA1032:Implement standard exception constructors",
            Justification = "A caller's abort always carries a code; a message-only abort has no meaning.")]
        [SuppressMessage("Design", "CA1064:Exceptions should be public",
            Justification = "Signals an abort between Build and Run; it never escapes to a caller.")]
        private sealed class CsvAggregateAbortException(int code) : Exception
        {
            internal int Code { get; } = code;
        }

        private sealed class CsvAggregateContext(NativeCsvAggregationRaw raw)
        {
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

            internal static void Abort(int code)
            {
                throw new CsvAggregateAbortException(code);
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
            NativeApi.ClearLastError();
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
                Task.Run(() => CsvParallel.AggregateAsync(path, Build(context), options, CancellationToken.None)),
                out result);
        }

        internal static int AggregateCsvMemory(
            byte* data, int dataLength, NativeCsvAggregationRaw aggregation,
            NativeCsvParallelOptionsRaw? rawOptions, out nint result)
        {
            NativeApi.ClearLastError();
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
                Task.Run(() => CsvParallel.AggregateAsync(manager.Memory, Build(context), options, CancellationToken.None)),
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
                        CsvAggregateContext.Abort(status);
                    }
                    return state;
                },
                Accumulate = (ref CsvAggregateState state, Row row) =>
                {
                    NativeRow native = state.WriteRow(row);
                    int status = context.Accumulate((void*)state.Native, &native, context.UserData);
                    if (status != NativeStatus.Ok)
                    {
                        CsvAggregateContext.Abort(status);
                    }
                },
                Combine = (accumulator, next) =>
                {
                    int status = context.Combine(
                        (void*)accumulator.Native, (void*)next.Native, context.UserData);
                    if (status != NativeStatus.Ok)
                    {
                        CsvAggregateContext.Abort(status);
                    }
                    return accumulator;
                },
            };
        }

        private static int Run(CsvAggregateContext context, Task<CsvAggregateState> pending, out nint result)
        {
            result = 0;
            CsvAggregateState? winner = null;
            try
            {
                winner = pending.GetAwaiter().GetResult();
                result = winner.Native;
                return NativeStatus.Ok;
            }
            catch (CsvAggregateAbortException abort)
            {
                return abort.Code;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
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
        }
    }
}
