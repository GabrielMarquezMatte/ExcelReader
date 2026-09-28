using System.Runtime.InteropServices;
using ExcelReader.Native.Reading;

namespace ExcelReader.Native
{
    internal static unsafe partial class Exports
    {
        [UnmanagedCallersOnly(EntryPoint = "xl_open_file")]
        public static int OpenFile(byte* path, int pathLength, int format, nint* outHandle)
        {
            if (!IsValidOpenRequest(path, pathLength, outHandle))
            {
                return NativeStatus.InvalidArgument;
            }

            int status = ReadApi.OpenFile(new ReadOnlySpan<byte>(path, pathLength), format, out NativeHandle? handle);
            return RegisterOpened(status, handle, outHandle);
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_open_file_ex")]
        public static int OpenFileEx(byte* path, int pathLength, int format, NativeOpenOptionsRaw* options, nint* outHandle)
        {
            if (!IsValidOpenRequest(path, pathLength, outHandle))
            {
                return NativeStatus.InvalidArgument;
            }

            if (!TryReadOpenOptions(options, out NativeOpenOptionsRaw? rawOptions))
            {
                return NativeStatus.InvalidArgument;
            }
            int status = ReadApi.OpenFileEx(new ReadOnlySpan<byte>(path, pathLength), format, rawOptions, out NativeHandle? handle);
            return RegisterOpened(status, handle, outHandle);
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_open_memory")]
        public static int OpenMemory(byte* data, int dataLength, int format, nint* outHandle)
        {
            if (!IsValidOpenRequest(data, dataLength, outHandle))
            {
                return NativeStatus.InvalidArgument;
            }

            int status = ReadApi.OpenMemory(new ReadOnlySpan<byte>(data, dataLength), format, out NativeHandle? handle);
            return RegisterOpened(status, handle, outHandle);
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_open_memory_ex")]
        public static int OpenMemoryEx(byte* data, int dataLength, int format, NativeOpenOptionsRaw* options, nint* outHandle)
        {
            if (!IsValidOpenRequest(data, dataLength, outHandle))
            {
                return NativeStatus.InvalidArgument;
            }

            if (!TryReadOpenOptions(options, out NativeOpenOptionsRaw? rawOptions))
            {
                return NativeStatus.InvalidArgument;
            }
            int status = ReadApi.OpenMemoryEx(new ReadOnlySpan<byte>(data, dataLength), format, rawOptions, out NativeHandle? handle);
            return RegisterOpened(status, handle, outHandle);
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_close")]
        public static int Close(nint handle)
        {
            if (handle == 0)
            {
                return NativeStatus.InvalidHandle;
            }

            if (!TryFree(handle, out NativeHandle? target))
            {
                return NativeStatus.InvalidHandle;
            }

            return ReadApi.Close(target);
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_sheet_count")]
        public static int SheetCount(nint handle, int* outCount)
        {
            if (outCount is null)
            {
                return NativeStatus.InvalidArgument;
            }

            int status = ReadApi.SheetCount(Resolve(handle), out int count);
            *outCount = count;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_sheet_name")]
        public static int SheetName(nint handle, byte* buffer, int capacity, int* outLength)
        {
            if (!IsValidOutBuffer(buffer, capacity, outLength))
            {
                return NativeStatus.InvalidArgument;
            }

            int status = ReadApi.SheetName(Resolve(handle), new Span<byte>(buffer, capacity), out int length);
            *outLength = length;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_sheet_name_at")]
        public static int SheetNameAt(nint handle, int index, byte* buffer, int capacity, int* outLength)
        {
            if (!IsValidOutBuffer(buffer, capacity, outLength))
            {
                return NativeStatus.InvalidArgument;
            }

            int status = ReadApi.SheetNameAt(Resolve(handle), index, new Span<byte>(buffer, capacity), out int length);
            *outLength = length;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_move_to_sheet")]
        public static int MoveToSheet(nint handle, int index)
        {
            return ReadApi.MoveToSheet(Resolve(handle), index);
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_is_date1904")]
        public static int IsDate1904(nint handle, int* outFlag)
        {
            if (outFlag is null)
            {
                return NativeStatus.InvalidArgument;
            }

            int status = ReadApi.IsDate1904(Resolve(handle), out int flag);
            *outFlag = flag;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_next_row")]
        public static int NextRow(nint handle, byte* buffer, int capacity, int* outWritten)
        {
            if (!IsValidOutBuffer(buffer, capacity, outWritten))
            {
                return NativeStatus.InvalidArgument;
            }

            int status = ReadApi.NextRow(Resolve(handle), new Span<byte>(buffer, capacity), out int written);
            *outWritten = written;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_next_row_view")]
        public static int NextRowView(nint handle, NativeRow* outRow)
        {
            if (outRow is null)
            {
                return NativeStatus.InvalidArgument;
            }

            int status = ReadApi.NextRowView(Resolve(handle), out NativeRow row);
            *outRow = row;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_read_all_blob")]
        public static int ReadAllBlob(nint handle, byte* buffer, int capacity, int* outWritten)
        {
            if (!IsValidOutBuffer(buffer, capacity, outWritten))
            {
                return NativeStatus.InvalidArgument;
            }

            int status = ReadApi.ReadAllBlob(Resolve(handle), new Span<byte>(buffer, capacity), out int written);
            *outWritten = written;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_read_all_decoded")]
        public static int ReadAllDecoded(nint handle, NativeRows* outRows)
        {
            if (outRows is null)
            {
                return NativeStatus.InvalidArgument;
            }

            int status = ReadApi.ReadAllDecoded(Resolve(handle), out NativeRows rows);
            *outRows = rows;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_free_rows")]
        public static void FreeRows(NativeRows* rows)
        {
            if (rows is null)
            {
                return;
            }
            try
            {
                ReadApi.FreeRows(ref *rows);
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
            }
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_infer_schema")]
        public static int InferSchema(nint handle, int headerRow, int sampleSize, NativeInferredSchema* outSchema)
        {
            if (outSchema is null)
            {
                return NativeStatus.InvalidArgument;
            }

            int status = ReadApi.InferSchema(Resolve(handle), headerRow, sampleSize, out NativeInferredSchema schema);
            *outSchema = schema;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_infer_schema_ex")]
        public static int InferSchemaEx(nint handle, int headerRow, int sampleSize, int flags, NativeInferredSchema* outSchema)
        {
            if (outSchema is null)
            {
                return NativeStatus.InvalidArgument;
            }

            int status = ReadApi.InferSchema(Resolve(handle), headerRow, sampleSize, flags, out NativeInferredSchema schema);
            *outSchema = schema;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_free_schema")]
        public static void FreeSchema(NativeInferredSchema* schema)
        {
            if (schema is null)
            {
                return;
            }
            try
            {
                ReadApi.FreeSchema(ref *schema);
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
            }
        }

        private static bool IsValidOpenRequest(byte* source, int sourceLength, nint* outHandle)
        {
            return source is not null && sourceLength >= 0 && outHandle is not null;
        }

        private static bool TryReadOpenOptions(NativeOpenOptionsRaw* options, out NativeOpenOptionsRaw? rawOptions)
        {
            if (options is null)
            {
                rawOptions = null;
                return true;
            }
            int expectedSize = sizeof(NativeOpenOptionsRaw);
            if (options->StructSize != expectedSize)
            {
                NativeApi.SetLastError($"xl_open_options.struct_size is {options->StructSize}, but this library expects {expectedSize}.");
                rawOptions = null;
                return false;
            }
            rawOptions = *options;
            return true;
        }

        private static int RegisterOpened(int status, NativeHandle? handle, nint* outHandle)
        {
            nint id = 0;
            if (handle is not null)
            {
                id = NativeHandleTable.Register(handle);
            }
            *outHandle = id;
            return status;
        }
    }
}
