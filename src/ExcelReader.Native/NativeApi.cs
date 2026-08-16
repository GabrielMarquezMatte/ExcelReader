using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelReader.Native
{
    /// <summary>
    /// Span-based implementation behind the C ABI. Every method returns a <see cref="NativeStatus"/>
    /// code and never throws; <see cref="Exports"/> only converts pointers into spans on top of it.
    /// This split exists because managed code cannot call an [UnmanagedCallersOnly] method, so the
    /// exports themselves are untestable — this layer is what the test suite drives.
    /// </summary>
    internal static partial class NativeApi
    {
        // Thread-local because handles are single-threaded by contract: an error raised on one
        // thread must not be observable from another.
        [ThreadStatic]
        private static string? _lastError;

        // Backs xl_last_error_ptr: a borrowed-pointer alternative to xl_last_error that skips the
        // ask-the-size-then-copy round trip. A managed string has no stable address, so the UTF-8 bytes
        // are kept in a separate, pinned, thread-static buffer the pointer can point straight into.
        // Grown, never shrunk, so a pointer handed out by an earlier call before a regrow would dangle —
        // LastErrorPtr always re-reads the current buffer, so that never happens across a single call,
        // but a caller must not cache the pointer past its own thread's next ExcelReader call (documented
        // on xl_last_error_ptr in excelreader.h).
        //
        // Held via a pinned GCHandle, not a plain `byte[]` field: pinning only stops the GC from
        // *relocating* an object, not from *collecting* it, and a [ThreadStatic] field's storage is
        // reclaimed once its owning thread exits — a caller that reads xl_last_error_ptr from a
        // short-lived worker thread and copies it after that thread has already terminated would then
        // dereference collected/reused memory (observed as the buffer reading back all zero bytes).
        // GCHandle.Alloc roots the target in the runtime's handle table independent of any managed
        // reference to the handle itself, so the buffer survives its producing thread's death.
        // ponytail: the handle for a thread that errors once and never calls again leaks for the
        // process lifetime (a few hundred bytes, bounded by the number of distinct threads that ever
        // called SetLastError) — upgrade path is a bounded per-thread handle registry with cleanup on
        // thread exit, not worth the complexity for a native error-message buffer.
        [ThreadStatic]
        private static GCHandle _lastErrorHandle;
        [ThreadStatic]
        private static int _lastErrorUtf8Length;

        internal static void SetLastError(string message)
        {
            _lastError = message;
            int required = Encoding.UTF8.GetByteCount(message);
            byte[]? current = _lastErrorHandle.IsAllocated ? Unsafe.As<byte[]>(_lastErrorHandle.Target) : null;
            if (current is null || current.Length < required)
            {
                if (_lastErrorHandle.IsAllocated)
                {
                    _lastErrorHandle.Free();
                }
                current = new byte[Math.Max(required, 256)];
                _lastErrorHandle = GCHandle.Alloc(current, GCHandleType.Pinned);
            }
            _lastErrorUtf8Length = Encoding.UTF8.GetBytes(message, current);
        }

        internal static void ClearLastError()
        {
            _lastError = null;
            _lastErrorUtf8Length = 0;
        }

        /// <summary>Copies the calling thread's last error message into <paramref name="buffer"/> as UTF-8.</summary>
        internal static int LastError(Span<byte> buffer, out int length)
        {
            string? message = _lastError;
            if (string.IsNullOrEmpty(message))
            {
                length = 0;
                return NativeStatus.Ok;
            }

            int required = Encoding.UTF8.GetByteCount(message);
            length = required;
            if (buffer.Length < required)
            {
                return NativeStatus.BufferTooSmall;
            }

            Encoding.UTF8.GetBytes(message, buffer);
            return NativeStatus.Ok;
        }

        /// <summary>
        /// Borrowed pointer to the calling thread's last error message, UTF-8, not NUL-terminated.
        /// Returns <see cref="IntPtr.Zero"/> with <paramref name="length"/> zero when there is no error.
        /// The pointer is only valid until the next ExcelReader call on this thread — see the ownership
        /// note on <c>xl_last_error_ptr</c> in excelreader.h.
        /// </summary>
        /// <remarks>
        /// Returns <see cref="nint"/>, not <see cref="byte"/>*, so this method (and the test suite that
        /// drives it) doesn't need an unsafe context — <see cref="Exports.LastErrorPtr"/> is the one place
        /// this becomes an actual pointer, cast at the ABI boundary where unsafe code already lives.
        /// </remarks>
        internal static nint LastErrorPtr(out int length)
        {
            length = _lastErrorUtf8Length;
            if (length == 0 || !_lastErrorHandle.IsAllocated)
            {
                return IntPtr.Zero;
            }

            // Safe without a `fixed` block: the target is pinned by the GCHandle itself (allocated with
            // GCHandleType.Pinned), so its address is stable for as long as the handle stays allocated —
            // which outlives the producing thread, exactly what a pointer surviving past this method
            // return, into the ABI caller, requires.
            return _lastErrorHandle.AddrOfPinnedObject();
        }
    }
}
