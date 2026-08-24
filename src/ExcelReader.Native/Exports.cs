using System.Runtime.InteropServices;
using System.Text;

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

        [UnmanagedCallersOnly(EntryPoint = "xl_next_row_decoded")]
        public static int NextRowDecoded(nint handle, NativeRow* outRow)
        {
            if (outRow is null)
            {
                return NativeStatus.InvalidArgument;
            }

            int status = NativeApi.NextRowDecoded(Resolve(handle), out NativeRow row);
            *outRow = row;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_free_row")]
        public static void FreeRow(NativeRow* row)
        {
            if (row is null)
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
                NativeApi.FreeRow(ref *row);
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
            }
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
            // See FreeRow's remarks: void in the ABI, so an exception here must never escape.
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
            // See FreeRow's remarks: void in the ABI, so an exception here must never escape.
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
            // See FreeRow's remarks: void in the ABI, so an exception here must never escape.
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
            // See FreeRow's remarks: void in the ABI (Arrow's own release-callback convention), so
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
            return NativeHandleTable.Resolve(handle);
        }

        internal static bool TryFree(nint handle, out NativeHandle? target)
        {
            return NativeHandleTable.TryUnregister(handle, out target);
        }
    }
}
