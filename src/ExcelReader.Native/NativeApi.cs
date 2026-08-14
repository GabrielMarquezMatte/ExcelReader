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
    internal static unsafe partial class NativeApi
    {
        // Thread-local because handles are single-threaded by contract: an error raised on one
        // thread must not be observable from another.
        [ThreadStatic]
        private static string? _lastError;

        // Backs xl_last_error_ptr: a borrowed-pointer alternative to xl_last_error that skips the
        // ask-the-size-then-copy round trip. A managed string has no stable address, so the UTF-8 bytes
        // are kept in a separate, pinned (GC never relocates it), thread-static array the pointer can
        // point straight into. Grown, never shrunk, so a pointer handed out by an earlier call before a
        // regrow would dangle — LastErrorPtr always re-reads the current array, so that never happens
        // across a single call, but a caller must not cache the pointer past its own thread's next
        // ExcelReader call (documented on xl_last_error_ptr in excelreader.h).
        [ThreadStatic]
        private static byte[]? _lastErrorUtf8;
        [ThreadStatic]
        private static int _lastErrorUtf8Length;

        internal static void SetLastError(string message)
        {
            _lastError = message;
            int required = Encoding.UTF8.GetByteCount(message);
            if (_lastErrorUtf8 is null || _lastErrorUtf8.Length < required)
            {
                // Pinned so xl_last_error_ptr can hand out a stable address; the GC must not relocate it.
                _lastErrorUtf8 = GC.AllocateArray<byte>(Math.Max(required, 256), pinned: true);
            }
            _lastErrorUtf8Length = Encoding.UTF8.GetBytes(message, _lastErrorUtf8);
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
            if (length == 0 || _lastErrorUtf8 is null)
            {
                return IntPtr.Zero;
            }

            // Safe without a `fixed` block: the array is allocated pinned (GC.AllocateArray with
            // pinned: true), so its address is stable for the array's whole lifetime, not just for the
            // scope of a `fixed` statement — which is exactly what a pointer surviving past this method
            // return, into the ABI caller, requires.
            return (nint)Unsafe.AsPointer(ref MemoryMarshal.GetArrayDataReference(_lastErrorUtf8));
        }
    }
}
