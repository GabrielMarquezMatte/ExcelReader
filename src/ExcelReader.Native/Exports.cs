using System.Runtime.InteropServices;

namespace ExcelReader.Native
{
    /// <summary>
    /// The C ABI. Every function here does exactly two things: turn raw pointers into spans and a
    /// GCHandle into a <see cref="NativeHandle"/>, then delegate to <see cref="NativeApi"/>.
    /// Keep the logic in NativeApi — managed code cannot call an [UnmanagedCallersOnly] method, so
    /// anything implemented here is untestable.
    /// </summary>
    internal static unsafe class Exports
    {
        [UnmanagedCallersOnly(EntryPoint = "xl_open_file")]
        public static int OpenFile(byte* path, int pathLength, int format, nint* outHandle)
        {
            if (path is null || pathLength < 0 || outHandle is null)
            {
                return NativeStatus.InvalidArgument;
            }

            int status = NativeApi.OpenFile(new ReadOnlySpan<byte>(path, pathLength), format, out NativeHandle? handle);
            *outHandle = handle is null ? 0 : GCHandle.ToIntPtr(GCHandle.Alloc(handle));
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_open_memory")]
        public static int OpenMemory(byte* data, int dataLength, int format, nint* outHandle)
        {
            if (data is null || dataLength < 0 || outHandle is null)
            {
                return NativeStatus.InvalidArgument;
            }

            int status = NativeApi.OpenMemory(new ReadOnlySpan<byte>(data, dataLength), format, out NativeHandle? handle);
            *outHandle = handle is null ? 0 : GCHandle.ToIntPtr(GCHandle.Alloc(handle));
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_close")]
        public static int Close(nint handle)
        {
            if (handle == 0)
            {
                return NativeStatus.InvalidHandle;
            }

            // A stale/garbage handle value makes GCHandle.FromIntPtr throw InvalidOperationException.
            // That must not become an unhandled exception crossing the boundary (it would crash the
            // process), so a bad handle here is reported the same way a null handle is: InvalidHandle.
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
            if (capacity < 0 || outLength is null || (buffer is null && capacity > 0))
            {
                return NativeStatus.InvalidArgument;
            }

            int status = NativeApi.SheetName(Resolve(handle), new Span<byte>(buffer, capacity), out int length);
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
            if (capacity < 0 || outWritten is null || (buffer is null && capacity > 0))
            {
                return NativeStatus.InvalidArgument;
            }

            int status = NativeApi.NextRow(Resolve(handle), new Span<byte>(buffer, capacity), out int written);
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
            if (row is not null)
            {
                NativeApi.FreeRow(ref *row);
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
            if (rows is not null)
            {
                NativeApi.FreeRows(ref *rows);
            }
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_last_error")]
        public static int LastError(byte* buffer, int capacity, int* outLength)
        {
            if (capacity < 0 || outLength is null || (buffer is null && capacity > 0))
            {
                return NativeStatus.InvalidArgument;
            }

            int status = NativeApi.LastError(new Span<byte>(buffer, capacity), out int length);
            *outLength = length;
            return status;
        }

        // Internal (not private) so tests can exercise the stale/garbage-handle path directly: the
        // [UnmanagedCallersOnly] entry points above cannot be invoked from managed code, but a plain
        // helper method can.
        internal static NativeHandle? Resolve(nint handle)
        {
            if (handle == 0)
            {
                return null;
            }

            try
            {
                return GCHandle.FromIntPtr(handle).Target as NativeHandle;
            }
            catch (InvalidOperationException)
            {
                // The handle value doesn't correspond to a live GCHandle allocation (stale, already
                // freed, or outright garbage). Every caller already treats a null result as
                // InvalidHandle, so surfacing null here keeps a bad handle a clean error instead of
                // a process crash.
                return null;
            }
        }

        internal static bool TryFree(nint handle, out NativeHandle? target)
        {
            // GCHandle.FromIntPtr never validates by itself (it only rejects IntPtr.Zero) — the
            // actual validity check happens in the VM when .Target or .Free() touches the handle
            // table, so both must be inside the try, not just the FromIntPtr call.
            try
            {
                GCHandle gcHandle = GCHandle.FromIntPtr(handle);
                target = gcHandle.Target as NativeHandle;
                gcHandle.Free();
                return true;
            }
            catch (InvalidOperationException)
            {
                target = null;
                return false;
            }
        }
    }
}
