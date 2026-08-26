using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Native.Writer;

namespace ExcelReader.Native
{
    /// <summary>
    /// The C ABI. Every function here does exactly two things: turn raw pointers into spans and a
    /// handle id (see <see cref="NativeHandleTable"/>) into a <see cref="NativeHandle"/>, then
    /// delegate to <see cref="NativeApi"/>. Keep the logic in NativeApi — managed code cannot call an
    /// [UnmanagedCallersOnly] method, so anything implemented here is untestable.
    /// </summary>
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

            NativeOpenOptionsRaw? rawOptions = options is null ? null : *options;
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

            NativeOpenOptionsRaw? rawOptions = options is null ? null : *options;
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

            // A stale (already-closed) or outright garbage handle value is reported the same way a
            // null handle is: InvalidHandle. NativeHandleTable retires an id permanently on a
            // successful unregister, so this can never free — or resolve to — the wrong workbook.
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
            // void in the ABI - no status code to report a failure through, and an exception
            // escaping [UnmanagedCallersOnly] is a fail-fast/abort for the native caller,
            // uncatchable in C/C++/Rust/Python. A caller that double-frees (holds a copy of an
            // already-freed struct - the header documents every Free* as "safe on a zeroed value",
            // not "safe on a stale copy") can make Marshal.FreeHGlobal throw; this keeps that from
            // crashing the whole process. The message still reaches xl_last_error for a caller that
            // checks after a free looked suspicious.
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
                // Decoding walks caller memory and allocates from a caller-supplied count, so it can
                // still fail in ways the guards above cannot see. Letting that escape would unwind
                // through the C caller's frame.
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
            // See FreeRows' remarks: void in the ABI, so an exception here must never escape.
            try
            {
                NativeApi.FreeTable(ref *table);
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
                // Decoding walks caller memory and allocates from a caller-supplied count, so it can
                // still fail in ways the guards above cannot see.
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
                // Same reasoning as xl_write_typed's catch: decoding walks caller memory and can still
                // fail in ways the guards above cannot see.
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
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

        // Copies a managed byte[] into unmanaged memory the caller owns until it calls xl_free_buffer.
        // Shared by xl_write_typed_to_memory and xl_write_handle_bytes. A null/empty result (including
        // a failed call, which leaves bytes null) publishes a zeroed xl_buffer rather than a 0-length
        // allocation - there is nothing for the caller to free either way, and NativeBuffer's default
        // is already exactly that.
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

        // A NULL options pointer means "every default", identical to xl_open_file_ex's contract. The
        // sheet name is UTF-8-decoded here because everything below this layer stays pointer-free.
        private static bool TryDecodeWriteOptions(NativeWriteOptionsRaw* options, out NativeWriteOptions decoded)
        {
            NativeWriteOptionsRaw raw = options is null
                ? new NativeWriteOptionsRaw { StructSize = Marshal.SizeOf<NativeWriteOptionsRaw>() }
                : *options;
            decoded = default;
            // struct_size is checked BEFORE sheet_name is touched: a size that disagrees means the
            // caller's struct layout is not this one, so the bytes sitting where sheet_name_len and
            // sheet_name should be cannot be trusted as a length and a pointer to dereference.
            if (!NativeWriteOptions.TryValidateStructSize(raw, out string? sizeError))
            {
                NativeApi.SetLastError(sizeError);
                return false;
            }
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
            // See FreeRows' remarks: void in the ABI, so an exception here must never escape.
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
                // Same hard-boundary reasoning as xl_parse_typed: spec decoding can throw on input the
                // guards above cannot rule out, and an escaping exception would unwind through C.
                NativeApi.SetLastError(exception.Message);
                *outArray = default;
                *outSchema = default;
                return NativeStatus.Error;
            }
        }

        // The argument contract every open entry point shares: a real source buffer, a non-negative
        // length for it, and somewhere to put the resulting handle id.
        private static bool IsValidOpenRequest(byte* source, int sourceLength, nint* outHandle)
        {
            return source is not null && sourceLength >= 0 && outHandle is not null;
        }

        // The argument contract every caller-supplied-buffer entry point shares. Kept in one place
        // because the expensive mistake at an ABI edge is a guard that gets tightened at four of its
        // five call sites.
        private static bool IsValidOutBuffer(byte* buffer, int capacity, int* outLength)
        {
            return capacity >= 0 && outLength is not null && (buffer is not null || capacity == 0);
        }

        // Every open path ends the same way: publish the new handle's id — 0 when the open failed and
        // produced none — and hand the status back untouched.
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

        // Shared by xl_parse_typed and xl_parse_arrow, whose column-spec input is identical. Returns
        // false for a name length that cannot describe a real header rather than passing it to
        // GetString, where it would become a read length over however much caller memory it names.
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

        // Invoked ONLY as a native function pointer value (ArrowSchema.Release) computed in
        // NativeApi.Arrow.cs — never called directly from managed code, so (like every other member
        // here) the actual free logic lives in the testable NativeApi layer.
        [UnmanagedCallersOnly]
        public static void ReleaseArrowSchemaCallback(ArrowSchema* schema)
        {
            if (schema is null)
            {
                return;
            }
            // See FreeRows' remarks: void in the ABI (Arrow's own release-callback convention), so
            // an exception here must never escape - doubly so for this one, since an Arrow consumer
            // owns when this runs, not this library.
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
            // See ReleaseArrowSchemaCallback's remarks.
            try
            {
                NativeApi.ReleaseArrowArray((IntPtr)array);
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
            }
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
            // Resolved directly (not via TryResolveWriter/RunWriterOp): those helpers return
            // XL_INVALID_HANDLE for an id that resolves to nothing, but a wrong-kind handle here
            // (opened via xl_open_write_handle, no MemoryStream) is a caller usage error against a
            // handle that IS valid, which NativeApi.GetWriteHandleBytes reports as InvalidArgument.
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

        // Shared by every writer entry point above whose body is just "resolve, try the one call,
        // translate an exception into Error". xl_start_sheet and xl_write_string are the exceptions:
        // both also decode a caller buffer, which can itself throw on malformed UTF-8, so they keep
        // their own inline try/catch around resolve-then-decode-then-call instead of using this helper.
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

        // Internal (not private) so tests can exercise this directly: the [UnmanagedCallersOnly] entry
        // points above cannot be invoked from managed code, but a plain helper method can.
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
