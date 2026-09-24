using System.Runtime.InteropServices;
using ExcelReader.Native.Csv;

namespace ExcelReader.Native
{
    internal static unsafe partial class Exports
    {
        [UnmanagedCallersOnly(EntryPoint = "xl_csv_aggregate_file")]
        public static int CsvAggregateFile(
            byte* path, int pathLength, NativeCsvAggregationRaw* aggregation,
            NativeCsvParallelOptionsRaw* options, void** outState)
        {
            if (outState is null)
            {
                return NativeStatus.InvalidArgument;
            }
            *outState = null;
            if (path is null || pathLength <= 0 || !IsValidAggregation(aggregation))
            {
                return NativeStatus.InvalidArgument;
            }

            NativeCsvParallelOptionsRaw? rawOptions = options is null ? null : *options;
            int status = CsvAggregateApi.AggregateCsvFile(
                new ReadOnlySpan<byte>(path, pathLength), *aggregation, rawOptions, out nint result);
            if (status == NativeStatus.Ok)
            {
                *outState = (void*)result;
            }
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_csv_aggregate_memory")]
        public static int CsvAggregateMemory(
            byte* data, int dataLength, NativeCsvAggregationRaw* aggregation,
            NativeCsvParallelOptionsRaw* options, void** outState)
        {
            if (outState is null)
            {
                return NativeStatus.InvalidArgument;
            }
            *outState = null;
            if (dataLength < 0 || (data is null && dataLength > 0) || !IsValidAggregation(aggregation))
            {
                return NativeStatus.InvalidArgument;
            }

            NativeCsvParallelOptionsRaw? rawOptions = options is null ? null : *options;
            int status = CsvAggregateApi.AggregateCsvMemory(
                data, dataLength, *aggregation, rawOptions, out nint result);
            if (status == NativeStatus.Ok)
            {
                *outState = (void*)result;
            }
            return status;
        }

        private static bool IsValidAggregation(NativeCsvAggregationRaw* aggregation)
        {
            if (aggregation is null)
            {
                NativeApi.SetLastError("csv_aggregation must not be NULL.");
                return false;
            }
            if (aggregation->StructSize != sizeof(NativeCsvAggregationRaw))
            {
                NativeApi.SetLastError(
                    $"csv_aggregation.struct_size must be exactly {sizeof(NativeCsvAggregationRaw)}; got {aggregation->StructSize}.");
                return false;
            }
            if (aggregation->Seed == IntPtr.Zero || aggregation->Accumulate == IntPtr.Zero
                || aggregation->Combine == IntPtr.Zero || aggregation->FreeState == IntPtr.Zero)
            {
                List<string> missing = [];
                if (aggregation->Seed == IntPtr.Zero)
                {
                    missing.Add("csv_aggregation.seed");
                }
                if (aggregation->Accumulate == IntPtr.Zero)
                {
                    missing.Add("csv_aggregation.accumulate");
                }
                if (aggregation->Combine == IntPtr.Zero)
                {
                    missing.Add("csv_aggregation.combine");
                }
                if (aggregation->FreeState == IntPtr.Zero)
                {
                    missing.Add("csv_aggregation.free_state");
                }
                NativeApi.SetLastError($"{string.Join(", ", missing)} must not be NULL.");
                return false;
            }
            return true;
        }
    }
}
