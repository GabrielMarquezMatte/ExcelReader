using System.Runtime.InteropServices;
using ExcelReader.Native.Arrow;
using ExcelReader.Native.Typed;

namespace ExcelReader.Native
{
    internal static unsafe partial class Exports
    {
        [UnmanagedCallersOnly(EntryPoint = "xl_parse_arrow")]
        public static int ParseArrow(nint handle, NativeColumnSpecRaw* specs, int specCount, int headerRow, ArrowArray* outArray, ArrowSchema* outSchema)
        {
            if (specs is null || outArray is null || outSchema is null || !TypedApi.IsValidSpecCount(specCount))
            {
                return NativeStatus.InvalidArgument;
            }

            try
            {
                if (!TryDecodeColumnSpecs(specs, specCount, out NativeColumnSpec[] decoded))
                {
                    *outArray = default;
                    *outSchema = default;
                    return NativeStatus.InvalidArgument;
                }

                int status = ArrowApi.ParseArrow(Resolve(handle), decoded, headerRow, out ArrowArray array, out ArrowSchema schema);
                *outArray = array;
                *outSchema = schema;
                return status;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                *outArray = default;
                *outSchema = default;
                return NativeStatus.Error;
            }
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_parse_arrow_ex")]
        public static int ParseArrowEx(nint handle, NativeColumnSpecRaw* specs, int specCount, int headerRow, int degreeOfParallelism,
            ArrowArray* outArray, ArrowSchema* outSchema)
        {
            if (specs is null || outArray is null || outSchema is null || !TypedApi.IsValidSpecCount(specCount))
            {
                return NativeStatus.InvalidArgument;
            }

            try
            {
                if (!TryDecodeColumnSpecs(specs, specCount, out NativeColumnSpec[] decoded))
                {
                    *outArray = default;
                    *outSchema = default;
                    return NativeStatus.InvalidArgument;
                }

                int status = ArrowApi.ParseArrow(Resolve(handle), decoded, headerRow, degreeOfParallelism, out ArrowArray array, out ArrowSchema schema);
                *outArray = array;
                *outSchema = schema;
                return status;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                *outArray = default;
                *outSchema = default;
                return NativeStatus.Error;
            }
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_parse_arrow_stream")]
        public static int ParseArrowStream(nint handle, NativeColumnSpecRaw* specs, int specCount, int headerRow,
            long maxRows, ArrowArrayStream* outStream)
        {
            if (outStream is null)
            {
                return NativeStatus.InvalidArgument;
            }
            *outStream = default;
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

                int status = ArrowApi.OpenArrowStream(Resolve(handle), decoded, headerRow, maxRows,
                    out ArrowArrayStream stream);
                *outStream = stream;
                return status;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                *outStream = default;
                return NativeStatus.Error;
            }
        }

        [UnmanagedCallersOnly]
        public static void ReleaseArrowSchemaCallback(ArrowSchema* schema)
        {
            if (schema is null)
            {
                return;
            }
            try
            {
                ArrowApi.ReleaseArrowSchema((IntPtr)schema);
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
            }
        }

        [UnmanagedCallersOnly]
        public static void ReleaseArrowArrayCallback(ArrowArray* array)
        {
            if (array is null)
            {
                return;
            }
            try
            {
                ArrowApi.ReleaseArrowArray((IntPtr)array);
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
            }
        }

        [UnmanagedCallersOnly]
        internal static int ArrowStreamGetSchema(ArrowArrayStream* stream, ArrowSchema* outSchema)
        {
            try { return ArrowApi.ArrowStreamGetSchemaCore(stream, outSchema); }
            catch { return 5; }
        }

        [UnmanagedCallersOnly]
        internal static int ArrowStreamGetNext(ArrowArrayStream* stream, ArrowArray* outArray)
        {
            try { return ArrowApi.ArrowStreamGetNextCore(stream, outArray); }
            catch { return 5; }
        }

        [UnmanagedCallersOnly]
        internal static IntPtr ArrowStreamGetLastError(ArrowArrayStream* stream)
        {
            try { return ArrowApi.ArrowStreamGetLastErrorCore(stream); }
            catch { return IntPtr.Zero; }
        }

        [UnmanagedCallersOnly]
        internal static void ArrowStreamRelease(ArrowArrayStream* stream)
        {
            try { ArrowApi.ArrowStreamReleaseCore(stream); }
            catch { }
        }
    }
}
