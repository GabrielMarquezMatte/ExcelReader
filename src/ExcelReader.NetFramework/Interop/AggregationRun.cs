using System.Diagnostics.CodeAnalysis;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using ExcelReader.NetFramework.Reader;

namespace ExcelReader.NetFramework.Interop
{
    internal abstract unsafe class AggregationRun
    {
        private const int CallbackFailed = int.MaxValue;

        private static readonly SeedFn SeedCallback = OnSeed;
        private static readonly AccumulateFn AccumulateCallback = OnAccumulate;
        private static readonly CombineFn CombineCallback = OnCombine;
        private static readonly FreeStateFn FreeStateCallback = OnFreeState;
        private static readonly IntPtr SeedPointer = Marshal.GetFunctionPointerForDelegate(SeedCallback);
        private static readonly IntPtr AccumulatePointer = Marshal.GetFunctionPointerForDelegate(AccumulateCallback);
        private static readonly IntPtr CombinePointer = Marshal.GetFunctionPointerForDelegate(CombineCallback);
        private static readonly IntPtr FreeStatePointer = Marshal.GetFunctionPointerForDelegate(FreeStateCallback);

        private ExceptionDispatchInfo? _error;

        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate int SeedFn(IntPtr* outState, IntPtr userData);

        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate int AccumulateFn(IntPtr state, XlCellView* cells, int cellCount, IntPtr userData);

        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate int CombineFn(IntPtr accumulator, IntPtr next, IntPtr userData);

        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate void FreeStateFn(IntPtr state, IntPtr userData);

        internal static TAccumulator Run<TAccumulator>(ExcelWorkbook workbook, int sheet, Func<TAccumulator> seed, int headerRow, int degreeOfParallelism)
            where TAccumulator : class, IExcelAccumulator<TAccumulator>
        {
            AggregationRun run = new AggregationRun<TAccumulator>(seed);
            GCHandle self = GCHandle.Alloc(run);
            XlAggregation aggregation = new XlAggregation
            {
                StructSize = sizeof(XlAggregation),
                Seed = SeedPointer,
                Accumulate = AccumulatePointer,
                Combine = CombinePointer,
                FreeState = FreeStatePointer,
                UserData = GCHandle.ToIntPtr(self),
            };
            WorkbookHandle handle = workbook.Handle;
            bool added = false;
            try
            {
                handle.DangerousAddRef(ref added);
                IntPtr state = IntPtr.Zero;
                int status = NativeMethods.xl_aggregate(handle.DangerousGetHandle(), sheet, &aggregation, headerRow, degreeOfParallelism, &state);
                object? result = Take(state);
                if (status == CallbackFailed)
                {
                    run._error?.Throw();
                }

                NativeError.ThrowIfFailed(status);
                if (result is null)
                {
                    throw new ExcelReaderException(ExcelStatus.Error, "The native library reported success but returned no aggregation state.");
                }

                return (TAccumulator)result;
            }
            finally
            {
                self.Free();
                if (added)
                {
                    handle.DangerousRelease();
                }
            }
        }

        protected abstract object Seed();

        protected abstract void Accumulate(object state, ExcelRowView row);

        protected abstract void Combine(object accumulator, object next);

        private static AggregationRun From(IntPtr userData)
        {
            return (AggregationRun)GCHandle.FromIntPtr(userData).Target!;
        }

        private static object Target(IntPtr state)
        {
            return GCHandle.FromIntPtr(state).Target!;
        }

        private static object? Take(IntPtr state)
        {
            if (state == IntPtr.Zero)
            {
                return null;
            }

            GCHandle handle = GCHandle.FromIntPtr(state);
            object? target = handle.Target;
            handle.Free();
            return target;
        }

        [SuppressMessage("Design", "CA1031:Do not catch general exception types",
            Justification = "An exception must not cross the native callback boundary; it is rethrown on return.")]
        private static int OnSeed(IntPtr* outState, IntPtr userData)
        {
            *outState = IntPtr.Zero;
            AggregationRun run = From(userData);
            try
            {
                *outState = GCHandle.ToIntPtr(GCHandle.Alloc(run.Seed()));
                return (int)ExcelStatus.Ok;
            }
            catch (Exception exception)
            {
                return run.Fail(exception);
            }
        }

        [SuppressMessage("Design", "CA1031:Do not catch general exception types",
            Justification = "An exception must not cross the native callback boundary; it is rethrown on return.")]
        private static int OnAccumulate(IntPtr state, XlCellView* cells, int cellCount, IntPtr userData)
        {
            AggregationRun run = From(userData);
            try
            {
                run.Accumulate(Target(state), new ExcelRowView(cells, cellCount));
                return (int)ExcelStatus.Ok;
            }
            catch (Exception exception)
            {
                return run.Fail(exception);
            }
        }

        [SuppressMessage("Design", "CA1031:Do not catch general exception types",
            Justification = "An exception must not cross the native callback boundary; it is rethrown on return.")]
        private static int OnCombine(IntPtr accumulator, IntPtr next, IntPtr userData)
        {
            AggregationRun run = From(userData);
            try
            {
                run.Combine(Target(accumulator), Target(next));
                return (int)ExcelStatus.Ok;
            }
            catch (Exception exception)
            {
                return run.Fail(exception);
            }
        }

        [SuppressMessage("Design", "CA1031:Do not catch general exception types",
            Justification = "An exception must not cross the native callback boundary.")]
        private static void OnFreeState(IntPtr state, IntPtr userData)
        {
            try
            {
                GCHandle.FromIntPtr(state).Free();
            }
            catch (Exception)
            {
            }
        }

        // The native side may re-read a failed chunk and still succeed, so the exception surfaces only if the run fails.
        private int Fail(Exception exception)
        {
            Interlocked.CompareExchange(ref _error, ExceptionDispatchInfo.Capture(exception), null);
            return CallbackFailed;
        }
    }

    internal sealed class AggregationRun<TAccumulator> : AggregationRun
        where TAccumulator : class, IExcelAccumulator<TAccumulator>
    {
        private readonly Func<TAccumulator> _seed;

        internal AggregationRun(Func<TAccumulator> seed)
        {
            _seed = seed;
        }

        protected override object Seed()
        {
            TAccumulator? accumulator = _seed();
            if (accumulator is null)
            {
                throw new InvalidOperationException($"The seed returned a null {typeof(TAccumulator).Name}.");
            }

            return accumulator;
        }

        protected override void Accumulate(object state, ExcelRowView row)
        {
            ((TAccumulator)state).Accumulate(row);
        }

        protected override void Combine(object accumulator, object next)
        {
            ((TAccumulator)accumulator).Combine((TAccumulator)next);
        }
    }
}
