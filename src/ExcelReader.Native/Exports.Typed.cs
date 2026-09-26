using System.Runtime.InteropServices;
using ExcelReader.Native.Typed;

namespace ExcelReader.Native
{
    internal static unsafe partial class Exports
    {
        [UnmanagedCallersOnly(EntryPoint = "xl_parse_typed")]
        public static int ParseTyped(nint handle, NativeColumnSpecRaw* specs, int specCount, int headerRow, NativeTable* outTable)
        {
            if (specs is null || outTable is null || !TypedApi.IsValidSpecCount(specCount))
            {
                return NativeStatus.InvalidArgument;
            }

            try
            {
                if (!TryDecodeColumnSpecs(specs, specCount, out NativeColumnSpec[] decoded))
                {
                    *outTable = default;
                    return NativeStatus.InvalidArgument;
                }

                int status = TypedApi.ParseTyped(Resolve(handle), decoded, headerRow, out NativeTable table);
                *outTable = table;
                return status;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                *outTable = default;
                return NativeStatus.Error;
            }
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_parse_typed_ex")]
        public static int ParseTypedEx(nint handle, NativeColumnSpecRaw* specs, int specCount, int headerRow, int degreeOfParallelism, NativeTable* outTable)
        {
            if (specs is null || outTable is null || !TypedApi.IsValidSpecCount(specCount))
            {
                return NativeStatus.InvalidArgument;
            }

            try
            {
                if (!TryDecodeColumnSpecs(specs, specCount, out NativeColumnSpec[] decoded))
                {
                    *outTable = default;
                    return NativeStatus.InvalidArgument;
                }

                int status = TypedApi.ParseTypedTable(Resolve(handle), decoded, headerRow, degreeOfParallelism, "xl_parse_typed_ex", out NativeTable table);
                *outTable = table;
                return status;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                *outTable = default;
                return NativeStatus.Error;
            }
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_free_table")]
        public static void FreeTable(NativeTable* table)
        {
            if (table is null)
            {
                return;
            }
            try
            {
                TypedApi.FreeTable(ref *table);
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
            }
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_typed_reader_open")]
        public static int TypedReaderOpen(nint handle, NativeColumnSpecRaw* specs, int specCount, int headerRow,
            long maxRows, nint* outReader)
        {
            if (outReader is null)
            {
                return NativeStatus.InvalidArgument;
            }
            *outReader = 0;
            if (specs is null || !TypedApi.IsValidSpecCount(specCount))
            {
                return NativeStatus.InvalidArgument;
            }

            try
            {
                if (!TryDecodeColumnSpecs(specs, specCount, out NativeColumnSpec[] decoded))
                {
                    return NativeStatus.InvalidArgument;
                }

                int status = TypedApi.OpenTypedReader(Resolve(handle), decoded, headerRow, maxRows, out nint reader);
                *outReader = reader;
                return status;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                *outReader = 0;
                return NativeStatus.Error;
            }
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_typed_reader_next")]
        public static int TypedReaderNext(nint reader, NativeTable* outTable)
        {
            if (outTable is null)
            {
                return NativeStatus.InvalidArgument;
            }
            *outTable = default;

            try
            {
                int status = TypedApi.NextTypedBatch(reader, out NativeTable table);
                *outTable = table;
                return status;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                *outTable = default;
                return NativeStatus.Error;
            }
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_typed_reader_close")]
        public static void TypedReaderClose(nint reader)
        {
            try
            {
                TypedApi.CloseTypedReader(reader);
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
            }
        }
    }
}
