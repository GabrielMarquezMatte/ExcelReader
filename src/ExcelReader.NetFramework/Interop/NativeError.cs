using System.Text;

namespace ExcelReader.NetFramework.Interop
{
    internal static unsafe class NativeError
    {
        internal static void ThrowIfFailed(int status)
        {
            if (status == (int)ExcelStatus.Ok) return;
            throw Create(status);
        }

        internal static ExcelReaderException Create(int status)
        {
            var typed = (ExcelStatus)status;
            var length = 0;
            var pointer = NativeMethods.xl_last_error_ptr(&length);
            var message = pointer == null || length <= 0 ? string.Empty : Encoding.UTF8.GetString(pointer, length);
            return new ExcelReaderException(typed, message.Length == 0 ? $"Native call failed with status {typed}." : message);
        }
    }
}
