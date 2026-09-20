using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Native.Writer;

namespace ExcelReader.Native
{
    [ExcludeFromCodeCoverage]
    internal static unsafe class Exports
    {
        [UnmanagedCallersOnly(EntryPoint = "xl_open_file")]
        public static int OpenFile(byte* path, int pathLength, int format, nint* outHandle)
        {
            if (!IsValidOpenRequest(path, pathLength, outHandle))
            {
                return NativeStatus.InvalidArgument;
            }

            int status = NativeApi.OpenFile(new ReadOnlySpan<byte>(path, pathLength), format, out NativeHandle? handle);
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
            int status = NativeApi.OpenFileEx(new ReadOnlySpan<byte>(path, pathLength), format, rawOptions, out NativeHandle? handle);
            return RegisterOpened(status, handle, outHandle);
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_open_memory")]
        public static int OpenMemory(byte* data, int dataLength, int format, nint* outHandle)
        {
            if (!IsValidOpenRequest(data, dataLength, outHandle))
            {
                return NativeStatus.InvalidArgument;
            }

            int status = NativeApi.OpenMemory(new ReadOnlySpan<byte>(data, dataLength), format, out NativeHandle? handle);
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
            int status = NativeApi.OpenMemoryEx(new ReadOnlySpan<byte>(data, dataLength), format, rawOptions, out NativeHandle? handle);
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

            return NativeApi.Close(target);
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_sheet_count")]
        public static int SheetCount(nint handle, int* outCount)
        {
            if (outCount is null)
            {
                return NativeStatus.InvalidArgument;
            }

            int status = NativeApi.SheetCount(Resolve(handle), out int count);
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

            int status = NativeApi.SheetName(Resolve(handle), new Span<byte>(buffer, capacity), out int length);
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

            int status = NativeApi.SheetNameAt(Resolve(handle), index, new Span<byte>(buffer, capacity), out int length);
            *outLength = length;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_move_to_sheet")]
        public static int MoveToSheet(nint handle, int index)
        {
            return NativeApi.MoveToSheet(Resolve(handle), index);
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_is_date1904")]
        public static int IsDate1904(nint handle, int* outFlag)
        {
            if (outFlag is null)
            {
                return NativeStatus.InvalidArgument;
            }

            int status = NativeApi.IsDate1904(Resolve(handle), out int flag);
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

            int status = NativeApi.NextRow(Resolve(handle), new Span<byte>(buffer, capacity), out int written);
            *outWritten = written;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_read_all_blob")]
        public static int ReadAllBlob(nint handle, byte* buffer, int capacity, int* outWritten)
        {
            if (!IsValidOutBuffer(buffer, capacity, outWritten))
            {
                return NativeStatus.InvalidArgument;
            }

            int status = NativeApi.ReadAllBlob(Resolve(handle), new Span<byte>(buffer, capacity), out int written);
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

            int status = NativeApi.ReadAllDecoded(Resolve(handle), out NativeRows rows);
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
                NativeApi.FreeRows(ref *rows);
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
            }
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_parse_typed")]
        public static int ParseTyped(nint handle, NativeColumnSpecRaw* specs, int specCount, int headerRow, NativeTable* outTable)
        {
            if (specs is null || outTable is null || !NativeApi.IsValidSpecCount(specCount))
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

                int status = NativeApi.ParseTyped(Resolve(handle), decoded, headerRow, out NativeTable table);
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
                NativeApi.FreeTable(ref *table);
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
            if (specs is null || !NativeApi.IsValidSpecCount(specCount))
            {
                return NativeStatus.InvalidArgument;
            }

            try
            {
                if (!TryDecodeColumnSpecs(specs, specCount, out NativeColumnSpec[] decoded))
                {
                    return NativeStatus.InvalidArgument;
                }

                int status = NativeApi.OpenTypedReader(Resolve(handle), decoded, headerRow, maxRows, out nint reader);
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
                int status = NativeApi.NextTypedBatch(reader, out NativeTable table);
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
                NativeApi.CloseTypedReader(reader);
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
            }
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_write_typed")]
        public static int WriteTyped(byte* path, int pathLength, int format, NativeColumnSpecRaw* specs, NativeTable* table, NativeWriteOptionsRaw* options)
        {
            if (path is null || pathLength <= 0 || specs is null || table is null
                || !NativeApi.IsValidSpecCount(table->ColumnCount))
            {
                return NativeStatus.InvalidArgument;
            }

            try
            {
                if (!TryDecodeColumnSpecs(specs, table->ColumnCount, out NativeColumnSpec[] decoded))
                {
                    return NativeStatus.InvalidArgument;
                }
                if (!TryDecodeWriteOptions(options, out NativeWriteOptions decodedOptions))
                {
                    return NativeStatus.InvalidArgument;
                }
                return NativeApi.WriteTyped(new ReadOnlySpan<byte>(path, pathLength), format, decoded, *table, decodedOptions);
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_write_typed_to_memory")]
        public static int WriteTypedToMemory(int format, NativeColumnSpecRaw* specs, NativeTable* table, NativeWriteOptionsRaw* options, NativeBuffer* outBuffer)
        {
            if (specs is null || table is null || outBuffer is null || !NativeApi.IsValidSpecCount(table->ColumnCount))
            {
                return NativeStatus.InvalidArgument;
            }
            *outBuffer = default;

            try
            {
                if (!TryDecodeColumnSpecs(specs, table->ColumnCount, out NativeColumnSpec[] decoded))
                {
                    return NativeStatus.InvalidArgument;
                }
                if (!TryDecodeWriteOptions(options, out NativeWriteOptions decodedOptions))
                {
                    return NativeStatus.InvalidArgument;
                }
                int status = NativeApi.WriteTypedToMemory(format, decoded, *table, decodedOptions, out byte[]? bytes);
                PublishBuffer(bytes, outBuffer);
                return status;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_encrypt_package")]
        public static int EncryptPackage(byte* packagePath, int packagePathLength, byte* destinationPath, int destinationPathLength, byte* password, int passwordLength)
        {
            if (packagePath is null || packagePathLength <= 0
                || destinationPath is null || destinationPathLength <= 0
                || password is null || passwordLength <= 0)
            {
                return NativeStatus.InvalidArgument;
            }

            return NativeApi.EncryptPackage(
                new ReadOnlySpan<byte>(packagePath, packagePathLength),
                new ReadOnlySpan<byte>(destinationPath, destinationPathLength),
                new ReadOnlySpan<byte>(password, passwordLength));
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_free_buffer")]
        public static void FreeBuffer(NativeBuffer* buffer)
        {
            if (buffer is null || buffer->Data == IntPtr.Zero)
            {
                return;
            }
            Marshal.FreeHGlobal(buffer->Data);
            *buffer = default;
        }

        private static void PublishBuffer(byte[]? bytes, NativeBuffer* outBuffer)
        {
            if (bytes is null || bytes.Length == 0)
            {
                return;
            }
            IntPtr data = Marshal.AllocHGlobal(bytes.Length);
            Marshal.Copy(bytes, 0, data, bytes.Length);
            outBuffer->Data = data;
            outBuffer->Length = bytes.Length;
        }

        private static bool TryDecodeWriteOptions(NativeWriteOptionsRaw* options, out NativeWriteOptions decoded)
        {
            decoded = default;
            int expectedSize = sizeof(NativeWriteOptionsRaw);
            if (options is not null && options->StructSize != expectedSize)
            {
                NativeApi.SetLastError($"xl_write_options.struct_size is {options->StructSize}, but this library expects {expectedSize}.");
                return false;
            }
            NativeWriteOptionsRaw raw = options is null
                ? new NativeWriteOptionsRaw { StructSize = expectedSize }
                : *options;
            string? sheetName = null;
            if (raw.SheetName is not null)
            {
                if (!NativeApi.IsValidNameLength(raw.SheetNameLen))
                {
                    NativeApi.SetLastError($"xl_write_options.sheet_name_len is out of range; got {raw.SheetNameLen}.");
                    return false;
                }
                sheetName = Encoding.UTF8.GetString(raw.SheetName, raw.SheetNameLen);
            }
            if (!NativeWriteOptions.TryDecode(raw, sheetName, out decoded, out string? error))
            {
                NativeApi.SetLastError(error!);
                return false;
            }
            return true;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_infer_schema")]
        public static int InferSchema(nint handle, int headerRow, int sampleSize, NativeInferredSchema* outSchema)
        {
            if (outSchema is null)
            {
                return NativeStatus.InvalidArgument;
            }

            int status = NativeApi.InferSchema(Resolve(handle), headerRow, sampleSize, out NativeInferredSchema schema);
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
                NativeApi.FreeSchema(ref *schema);
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
            }
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_parse_arrow")]
        public static int ParseArrow(nint handle, NativeColumnSpecRaw* specs, int specCount, int headerRow, ArrowArray* outArray, ArrowSchema* outSchema)
        {
            if (specs is null || outArray is null || outSchema is null || !NativeApi.IsValidSpecCount(specCount))
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

                int status = NativeApi.ParseArrow(Resolve(handle), decoded, headerRow, out ArrowArray array, out ArrowSchema schema);
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
            if (specs is null || !NativeApi.IsValidSpecCount(specCount))
            {
                return NativeStatus.InvalidArgument;
            }

            try
            {
                if (!TryDecodeColumnSpecs(specs, specCount, out NativeColumnSpec[] decoded))
                {
                    return NativeStatus.InvalidArgument;
                }

                int status = NativeApi.OpenArrowStream(Resolve(handle), decoded, headerRow, maxRows,
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

        private static bool IsValidOutBuffer(byte* buffer, int capacity, int* outLength)
        {
            return capacity >= 0 && outLength is not null && (buffer is not null || capacity == 0);
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

        private static bool TryDecodeColumnSpecs(NativeColumnSpecRaw* specs, int specCount, out NativeColumnSpec[] decoded)
        {
            decoded = new NativeColumnSpec[specCount];
            for (int i = 0; i < specCount; i++)
            {
                NativeColumnSpecRaw raw = specs[i];
                if (!NativeApi.IsValidNameCount(raw.NameCount))
                {
                    decoded = [];
                    return false;
                }
                if (raw.NameCount > 0 && (raw.Names is null || raw.NameLens is null))
                {
                    decoded = [];
                    return false;
                }
                string[] names = new string[raw.NameCount];
                for (int n = 0; n < raw.NameCount; n++)
                {
                    if (!NativeApi.IsValidNameLength(raw.NameLens[n]))
                    {
                        decoded = [];
                        return false;
                    }
                    names[n] = Encoding.UTF8.GetString(raw.Names[n], raw.NameLens[n]);
                }
                decoded[i] = new NativeColumnSpec
                {
                    Names = names,
                    Index = raw.Index,
                    Type = raw.Type,
                    Nullable = raw.Nullable != 0,
                };
            }
            return true;
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
                NativeApi.ReleaseArrowSchema((IntPtr)schema);
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
                NativeApi.ReleaseArrowArray((IntPtr)array);
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
            }
        }

        [UnmanagedCallersOnly]
        internal static int ArrowStreamGetSchema(ArrowArrayStream* stream, ArrowSchema* outSchema)
        {
            try { return NativeApi.ArrowStreamGetSchemaCore(stream, outSchema); }
            catch { return 5; }
        }

        [UnmanagedCallersOnly]
        internal static int ArrowStreamGetNext(ArrowArrayStream* stream, ArrowArray* outArray)
        {
            try { return NativeApi.ArrowStreamGetNextCore(stream, outArray); }
            catch { return 5; }
        }

        [UnmanagedCallersOnly]
        internal static IntPtr ArrowStreamGetLastError(ArrowArrayStream* stream)
        {
            try { return NativeApi.ArrowStreamGetLastErrorCore(stream); }
            catch { return IntPtr.Zero; }
        }

        [UnmanagedCallersOnly]
        internal static void ArrowStreamRelease(ArrowArrayStream* stream)
        {
            try { NativeApi.ArrowStreamReleaseCore(stream); }
            catch { }
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_open_write_handle")]
        public static int OpenWriteHandle(byte* path, int pathLength, int format, NativeWriteOptionsRaw* options, nint* outHandle)
        {
            if (!IsValidOpenRequest(path, pathLength, outHandle))
            {
                return NativeStatus.InvalidArgument;
            }
            *outHandle = 0;
            if (!TryDecodeWriteOptions(options, out NativeWriteOptions decodedOptions))
            {
                return NativeStatus.InvalidArgument;
            }
            int status = NativeApi.OpenWriteHandle(new ReadOnlySpan<byte>(path, pathLength), format, decodedOptions, out NativeWriterHandle? handle);
            if (handle is not null)
            {
                *outHandle = NativeHandleTable.Register(handle);
            }
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_open_write_handle_to_memory")]
        public static int OpenWriteHandleToMemory(int format, NativeWriteOptionsRaw* options, nint* outHandle)
        {
            if (outHandle is null)
            {
                return NativeStatus.InvalidArgument;
            }
            *outHandle = 0;
            if (!TryDecodeWriteOptions(options, out NativeWriteOptions decodedOptions))
            {
                return NativeStatus.InvalidArgument;
            }
            int status = NativeApi.OpenWriteHandleToMemory(format, decodedOptions, out NativeWriterHandle? handle);
            if (handle is not null)
            {
                *outHandle = NativeHandleTable.Register(handle);
            }
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_write_handle_bytes")]
        public static int WriteHandleBytes(nint handle, NativeBuffer* outBuffer)
        {
            if (outBuffer is null)
            {
                return NativeStatus.InvalidArgument;
            }
            *outBuffer = default;
            NativeWriterHandle? writerHandle = NativeHandleTable.Resolve<NativeWriterHandle>(handle);
            int status = NativeApi.GetWriteHandleBytes(writerHandle, out byte[]? bytes);
            PublishBuffer(bytes, outBuffer);
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_start_sheet")]
        public static int StartSheet(nint handle, byte* name, int nameLength)
        {
            if (name is null || !NativeApi.IsValidNameLength(nameLength))
            {
                NativeApi.SetLastError($"xl_start_sheet's name_len is out of range; got {nameLength}.");
                return NativeStatus.InvalidArgument;
            }
            if (!TryResolveWriter(handle, out NativeWriterHandle? writerHandle))
            {
                return NativeStatus.InvalidHandle;
            }
            try
            {
                string sheetName = Encoding.UTF8.GetString(name, nameLength);
                writerHandle.StartSheet(sheetName);
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_start_row")]
        public static int StartRow(nint handle)
        {
            return RunWriterOp(handle, static writerHandle => writerHandle.StartRow());
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_write_string")]
        public static int WriteString(nint handle, byte* value, int valueLength)
        {
            if (value is not null && !NativeApi.IsValidNameLength(valueLength))
            {
                NativeApi.SetLastError($"xl_write_string's value_len is out of range; got {valueLength}.");
                return NativeStatus.InvalidArgument;
            }
            if (!TryResolveWriter(handle, out NativeWriterHandle? writerHandle))
            {
                return NativeStatus.InvalidHandle;
            }
            try
            {
                string? text = value is null ? null : Encoding.UTF8.GetString(value, valueLength);
                writerHandle.WriteString(text);
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_write_int64")]
        public static int WriteInt64(nint handle, long value)
        {
            return RunWriterOp(handle, writerHandle => writerHandle.WriteInt64(value));
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_write_float64")]
        public static int WriteFloat64(nint handle, double value)
        {
            return RunWriterOp(handle, writerHandle => writerHandle.WriteFloat64(value));
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_write_bool")]
        public static int WriteBool(nint handle, int value)
        {
            return RunWriterOp(handle, writerHandle => writerHandle.WriteBool(value != 0));
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_write_date")]
        public static int WriteDate(nint handle, int daysSinceEpoch)
        {
            return RunWriterOp(handle, writerHandle => writerHandle.WriteDate(daysSinceEpoch));
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_write_time")]
        public static int WriteTime(nint handle, long microsecondsSinceMidnight)
        {
            return RunWriterOp(handle, writerHandle => writerHandle.WriteTime(microsecondsSinceMidnight));
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_write_timestamp")]
        public static int WriteTimestamp(nint handle, long microsecondsSinceEpoch)
        {
            return RunWriterOp(handle, writerHandle => writerHandle.WriteTimestamp(microsecondsSinceEpoch));
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_write_null")]
        public static int WriteNull(nint handle, int type)
        {
            return RunWriterOp(handle, writerHandle => writerHandle.WriteNull(type));
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_end_row")]
        public static int EndRow(nint handle)
        {
            return RunWriterOp(handle, static writerHandle => writerHandle.EndRow());
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_end_sheet")]
        public static int EndSheet(nint handle)
        {
            return RunWriterOp(handle, static writerHandle => writerHandle.EndSheet());
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_close_write_handle")]
        public static int CloseWriteHandle(nint handle)
        {
            if (handle == 0)
            {
                return NativeStatus.InvalidHandle;
            }
            if (!NativeHandleTable.TryUnregister(handle, out NativeWriterHandle? target))
            {
                return NativeStatus.InvalidHandle;
            }
            return NativeApi.CloseWriteHandle(target);
        }

        private static int RunWriterOp(nint handle, Action<NativeWriterHandle> operation)
        {
            if (!TryResolveWriter(handle, out NativeWriterHandle? writerHandle))
            {
                return NativeStatus.InvalidHandle;
            }
            try
            {
                operation(writerHandle);
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        private static bool TryResolveWriter(nint handle, [NotNullWhen(true)] out NativeWriterHandle? writerHandle)
        {
            writerHandle = NativeHandleTable.Resolve<NativeWriterHandle>(handle);
            return writerHandle is not null;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_csv_aggregate_file")]
        public static int CsvAggregateFile(
            byte* path, int pathLength, NativeCsvAggregationRaw* aggregation,
            NativeCsvParallelOptionsRaw* options, void** outState)
        {
            if (path is null || pathLength <= 0 || outState is null || !IsValidAggregation(aggregation))
            {
                return NativeStatus.InvalidArgument;
            }

            NativeCsvParallelOptionsRaw? rawOptions = options is null ? null : *options;
            int status = NativeApi.AggregateCsvFile(
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
            if (dataLength < 0 || (data is null && dataLength > 0) || outState is null || !IsValidAggregation(aggregation))
            {
                return NativeStatus.InvalidArgument;
            }

            NativeCsvParallelOptionsRaw? rawOptions = options is null ? null : *options;
            int status = NativeApi.AggregateCsvMemory(
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

        [UnmanagedCallersOnly(EntryPoint = "xl_last_error")]
        public static int LastError(byte* buffer, int capacity, int* outLength)
        {
            if (!IsValidOutBuffer(buffer, capacity, outLength))
            {
                return NativeStatus.InvalidArgument;
            }

            int status = NativeApi.LastError(new Span<byte>(buffer, capacity), out int length);
            *outLength = length;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_last_error_ptr")]
        public static byte* LastErrorPtr(int* outLength)
        {
            if (outLength is null)
            {
                return null;
            }

            nint pointer = NativeApi.LastErrorPtr(out int length);
            *outLength = length;
            return (byte*)pointer;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_abi_version")]
        public static int AbiVersion()
        {
            return NativeStatus.AbiVersion;
        }

        internal static NativeHandle? Resolve(nint handle)
        {
            return NativeHandleTable.Resolve<NativeHandle>(handle);
        }

        internal static bool TryFree(nint handle, out NativeHandle? target)
        {
            return NativeHandleTable.TryUnregister(handle, out target);
        }
    }
}
