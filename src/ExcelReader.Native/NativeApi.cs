using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelReader.Native
{
    internal static partial class NativeApi
    {
        [ThreadStatic]
        private static string? _lastError;

        // ponytail: the handle for a thread that errors once and never calls again leaks for the
        // thread's lifetime; free it in a thread-exit callback if that matters.
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

        internal static string LastErrorText()
        {
            return _lastError ?? "";
        }

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

        internal static nint LastErrorPtr(out int length)
        {
            length = _lastErrorUtf8Length;
            if (length == 0 || !_lastErrorHandle.IsAllocated)
            {
                return IntPtr.Zero;
            }

            return _lastErrorHandle.AddrOfPinnedObject();
        }
    }
}
