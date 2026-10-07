using System.Runtime.InteropServices;
using ExcelReader.Native.Reading;

namespace ExcelReader.Native
{
    internal static unsafe partial class Exports
    {
        [UnmanagedCallersOnly(EntryPoint = "xl_open_file")]
        public static int OpenFile(byte* path, int pathLength, int format, NativeOpenOptionsRaw* options, nint* outHandle)
        {
            NativeApi.ClearLastError();
            if (outHandle is not null)
            {
                *outHandle = 0;
            }

            if (!IsValidOpenRequest(path, pathLength, outHandle))
            {
                return NativeStatus.InvalidArgument;
            }

            if (!TryReadOpenOptions(options, out NativeOpenOptionsRaw? rawOptions))
            {
                return NativeStatus.InvalidArgument;
            }
            int status = ReadApi.OpenFile(new ReadOnlySpan<byte>(path, pathLength), format, rawOptions, out NativeHandle? handle);
            return RegisterOpened(status, handle, outHandle);
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_open_memory")]
        public static int OpenMemory(byte* data, int dataLength, int format, NativeOpenOptionsRaw* options, nint* outHandle)
        {
            NativeApi.ClearLastError();
            if (outHandle is not null)
            {
                *outHandle = 0;
            }

            if (!IsValidOpenRequest(data, dataLength, outHandle))
            {
                return NativeStatus.InvalidArgument;
            }

            if (!TryReadOpenOptions(options, out NativeOpenOptionsRaw? rawOptions))
            {
                return NativeStatus.InvalidArgument;
            }
            int status = ReadApi.OpenMemory(new ReadOnlySpan<byte>(data, dataLength), format, rawOptions, out NativeHandle? handle);
            return RegisterOpened(status, handle, outHandle);
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_close")]
        public static int Close(nint handle)
        {
            NativeApi.ClearLastError();

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
            NativeApi.ClearLastError();

            if (outCount is null)
            {
                return NativeStatus.InvalidArgument;
            }

            int status = ReadApi.SheetCount(Resolve(handle), out int count);
            *outCount = count;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_sheet_name_at")]
        public static int SheetNameAt(nint handle, int index, byte* buffer, int capacity, int* outLength)
        {
            NativeApi.ClearLastError();

            if (!IsValidOutBuffer(buffer, capacity, outLength))
            {
                return NativeStatus.InvalidArgument;
            }

            int status = ReadApi.SheetNameAt(Resolve(handle), index, new Span<byte>(buffer, capacity), out int length);
            *outLength = length;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_is_date1904")]
        public static int IsDate1904(nint handle, int* outFlag)
        {
            NativeApi.ClearLastError();

            if (outFlag is null)
            {
                return NativeStatus.InvalidArgument;
            }

            int status = ReadApi.IsDate1904(Resolve(handle), out int flag);
            *outFlag = flag;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_sheet_visibility_at")]
        public static int SheetVisibilityAt(nint handle, int index, int* outVisibility)
        {
            NativeApi.ClearLastError();

            if (outVisibility is null)
            {
                return NativeStatus.InvalidArgument;
            }

            int status = ReadApi.SheetVisibilityAt(Resolve(handle), index, out int visibility);
            *outVisibility = visibility;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_sheet_index")]
        public static int SheetIndex(nint handle, byte* name, int nameLength, int* outIndex)
        {
            NativeApi.ClearLastError();

            if (outIndex is null)
            {
                return NativeStatus.InvalidArgument;
            }
            *outIndex = -1;
            if (nameLength < 0 || (name is null && nameLength != 0))
            {
                return NativeStatus.InvalidArgument;
            }

            int status = ReadApi.SheetIndex(Resolve(handle), new ReadOnlySpan<byte>(name, nameLength), out int index);
            *outIndex = index;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_rows_open")]
        public static int RowsOpen(nint handle, int sheet, nint* outRows)
        {
            NativeApi.ClearLastError();

            if (outRows is null)
            {
                return NativeStatus.InvalidArgument;
            }
            *outRows = 0;

            int status = ReadApi.OpenRows(Resolve(handle), sheet, out NativeRowCursor? cursor);
            if (cursor is not null)
            {
                *outRows = NativeHandleTable.Register(cursor);
            }
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_rows_close")]
        public static int RowsClose(nint rows)
        {
            NativeApi.ClearLastError();

            if (!NativeHandleTable.TryUnregister(rows, out NativeRowCursor? cursor))
            {
                return NativeStatus.InvalidHandle;
            }

            return ReadApi.CloseRows(cursor);
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_rows_next")]
        public static int RowsNext(nint rows, byte* buffer, int capacity, int* outWritten)
        {
            NativeApi.ClearLastError();

            if (!IsValidOutBuffer(buffer, capacity, outWritten))
            {
                return NativeStatus.InvalidArgument;
            }

            int status = ReadApi.NextRow(ResolveRows(rows), new Span<byte>(buffer, capacity), out int written);
            *outWritten = written;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_rows_next_view")]
        public static int RowsNextView(nint rows, NativeRow* outRow)
        {
            NativeApi.ClearLastError();

            if (outRow is null)
            {
                return NativeStatus.InvalidArgument;
            }

            int status = ReadApi.NextRowView(ResolveRows(rows), out NativeRow row);
            *outRow = row;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_rows_read_all_blob")]
        public static int RowsReadAllBlob(nint rows, byte* buffer, int capacity, int* outWritten)
        {
            NativeApi.ClearLastError();

            if (!IsValidOutBuffer(buffer, capacity, outWritten))
            {
                return NativeStatus.InvalidArgument;
            }

            int status = ReadApi.ReadAllBlob(ResolveRows(rows), new Span<byte>(buffer, capacity), out int written);
            *outWritten = written;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_rows_read_all_decoded")]
        public static int RowsReadAllDecoded(nint rows, NativeRows* outRows)
        {
            NativeApi.ClearLastError();

            if (outRows is null)
            {
                return NativeStatus.InvalidArgument;
            }

            int status = ReadApi.ReadAllDecoded(ResolveRows(rows), out NativeRows decoded);
            *outRows = decoded;
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
        public static int InferSchema(nint handle, int sheet, int headerRow, int sampleSize, int flags, NativeInferredSchema* outSchema)
        {
            NativeApi.ClearLastError();

            if (outSchema is null)
            {
                return NativeStatus.InvalidArgument;
            }

            int status = ReadApi.InferSchema(Resolve(handle), sheet, headerRow, sampleSize, flags, out NativeInferredSchema schema);
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
