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

        internal static void SetLastError(string message)
        {
            _lastError = message;
        }

        internal static void ClearLastError()
        {
            _lastError = null;
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
    }
}
