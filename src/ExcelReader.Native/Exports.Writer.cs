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
            NativeApi.ClearLastError();
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
            NativeApi.ClearLastError();
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
            if (!NativeWriteOptionsRaw.TryRead(options, out NativeWriteOptionsRaw raw, out string? readError))
            {
                NativeApi.SetLastError(readError);
                return false;
            }
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
            NativeApi.ClearLastError();
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
            NativeApi.ClearLastError();
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
            NativeApi.ClearLastError();
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
            NativeApi.ClearLastError();
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
            NativeApi.ClearLastError();
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
            return RunWriterOp(handle, value, static (writerHandle, state) => writerHandle.WriteInt64(state));
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_write_float64")]
        public static int WriteFloat64(nint handle, double value)
        {
            return RunWriterOp(handle, value, static (writerHandle, state) => writerHandle.WriteFloat64(state));
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_write_bool")]
        public static int WriteBool(nint handle, int value)
        {
            return RunWriterOp(handle, value != 0, static (writerHandle, state) => writerHandle.WriteBool(state));
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_write_date")]
        public static int WriteDate(nint handle, int daysSinceEpoch)
        {
            return RunWriterOp(handle, daysSinceEpoch, static (writerHandle, state) => writerHandle.WriteDate(state));
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_write_time")]
        public static int WriteTime(nint handle, long microsecondsSinceMidnight)
        {
            return RunWriterOp(handle, microsecondsSinceMidnight, static (writerHandle, state) => writerHandle.WriteTime(state));
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_write_timestamp")]
        public static int WriteTimestamp(nint handle, long microsecondsSinceEpoch)
        {
            return RunWriterOp(handle, microsecondsSinceEpoch, static (writerHandle, state) => writerHandle.WriteTimestamp(state));
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_write_null")]
        public static int WriteNull(nint handle, int type)
        {
            return RunWriterOp(handle, type, static (writerHandle, state) => writerHandle.WriteNull(state));
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

        [UnmanagedCallersOnly(EntryPoint = "xl_write_sheet")]
        public static int WriteSheet(nint handle, byte* name, int nameLength, NativeColumnSpecRaw* specs, NativeTable* table, NativeSheetOptionsRaw* options)
        {
            NativeApi.ClearLastError();
            if (name is null || !TypedApi.IsValidNameLength(nameLength) || specs is null || table is null
                || !TypedApi.IsValidSpecCount(table->ColumnCount))
            {
                NativeApi.SetLastError("xl_write_sheet needs a name, specs and a table with 1..max columns.");
                return NativeStatus.InvalidArgument;
            }
            if (!TryResolveWriter(handle, out NativeWriterHandle? writerHandle))
            {
                return NativeStatus.InvalidHandle;
            }
            try
            {
                if (!NativeExcelTable.TryRead(options, out NativeExcelTable? excelTable, out string? optionsError))
                {
                    NativeApi.SetLastError(optionsError);
                    return NativeStatus.InvalidArgument;
                }
                if (!TryDecodeColumnSpecs(specs, table->ColumnCount, out NativeColumnSpec[] decoded))
                {
                    NativeApi.SetLastError("xl_write_sheet got a column spec with an invalid name.");
                    return NativeStatus.InvalidArgument;
                }
                if (!WriteApi.TryValidateWriteTable(decoded, *table, out bool hasHeader, out string? validationError))
                {
                    NativeApi.SetLastError(validationError);
                    return NativeStatus.InvalidArgument;
                }
                if (excelTable is not null && !hasHeader)
                {
                    NativeApi.SetLastError("an Excel table needs every column named: its header row holds the column names.");
                    return NativeStatus.InvalidArgument;
                }
                writerHandle.WriteSheet(Encoding.UTF8.GetString(name, nameLength), decoded, *table, hasHeader, excelTable);
                NativeApi.ClearLastError();
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_close_write_handle")]
        public static int CloseWriteHandle(nint handle)
        {
            NativeApi.ClearLastError();
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

        [UnmanagedCallersOnly(EntryPoint = "xl_abort_write_handle")]
        public static int AbortWriteHandle(nint handle)
        {
            NativeApi.ClearLastError();
            if (!NativeHandleTable.TryUnregister(handle, out NativeWriterHandle? target))
            {
                return NativeStatus.InvalidHandle;
            }
            return WriteApi.AbortWriteHandle(target);
        }

        private static int RunWriterOp(nint handle, Action<NativeWriterHandle> operation)
        {
            NativeApi.ClearLastError();
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

        private static int RunWriterOp<TState>(nint handle, TState state, Action<NativeWriterHandle, TState> operation)
        {
            NativeApi.ClearLastError();
            if (!TryResolveWriter(handle, out NativeWriterHandle? writerHandle))
            {
                return NativeStatus.InvalidHandle;
            }
            try
            {
                operation(writerHandle, state);
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
