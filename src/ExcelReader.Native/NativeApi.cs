using System.Runtime.InteropServices;
using System.Text;

namespace ExcelReader.Native
{
    internal static partial class NativeApi
    {
        [ThreadStatic]
        private static string? _lastError;

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
                _lastErrorUtf8 = GC.AllocateUninitializedArray<byte>(Math.Max(required, 256), pinned: true);
            }
            _lastErrorUtf8Length = Encoding.UTF8.GetBytes(message, _lastErrorUtf8);
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
            if (length == 0 || _lastErrorUtf8 is null)
            {
                return IntPtr.Zero;
            }

            return Marshal.UnsafeAddrOfPinnedArrayElement(_lastErrorUtf8, 0);
        }
    }
}
