using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Native.Typed;
using ExcelReader.Native.Writer;

namespace ExcelReader.Native
{
    internal static unsafe partial class Exports
    {
        [UnmanagedCallersOnly(EntryPoint = "xl_write_typed")]
        public static int WriteTyped(byte* path, int pathLength, int format, NativeColumnSpecRaw* specs, NativeTable* table, NativeWriteOptionsRaw* options)
        {
            if (path is null || pathLength <= 0 || specs is null || table is null
                || !TypedApi.IsValidSpecCount(table->ColumnCount))
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
                return WriteApi.WriteTyped(new ReadOnlySpan<byte>(path, pathLength), format, decoded, *table, decodedOptions);
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
            if (specs is null || table is null || outBuffer is null || !TypedApi.IsValidSpecCount(table->ColumnCount))
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
                int status = WriteApi.WriteTypedToMemory(format, decoded, *table, decodedOptions, out byte[]? bytes);
                PublishBuffer(bytes, outBuffer);
                return status;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
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
                if (!TypedApi.IsValidNameLength(raw.SheetNameLen))
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
            int status = WriteApi.OpenWriteHandle(new ReadOnlySpan<byte>(path, pathLength), format, decodedOptions, out NativeWriterHandle? handle);
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
            int status = WriteApi.OpenWriteHandleToMemory(format, decodedOptions, out NativeWriterHandle? handle);
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
            int status = WriteApi.GetWriteHandleBytes(writerHandle, out byte[]? bytes);
            PublishBuffer(bytes, outBuffer);
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_start_sheet")]
        public static int StartSheet(nint handle, byte* name, int nameLength)
        {
            if (name is null || !TypedApi.IsValidNameLength(nameLength))
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
            if (value is not null && !TypedApi.IsValidNameLength(valueLength))
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
            return WriteApi.CloseWriteHandle(target);
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
    }
}
