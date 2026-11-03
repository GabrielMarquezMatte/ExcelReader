using System.Text;

namespace ExcelReader.NetFramework.Interop
{
    internal static unsafe class NativeError
    {
        internal static void ThrowIfFailed(int status)
        {
            if (status == (int)ExcelStatus.Ok)
            {
                return;
            }

            throw Create(status);
        }

        internal static ExcelReaderException Create(int status)
        {
            int length = 0;
            byte* pointer = NativeMethods.xl_last_error_ptr(&length);
            if (pointer is null || length <= 0)
            {
                return Create(status, null);
            }

            return Create(status, Encoding.UTF8.GetString(pointer, length));
        }

        internal static ExcelReaderException Create(int status, string? nativeMessage)
        {
            ExcelStatus typed = (ExcelStatus)status;
            if (nativeMessage is null || nativeMessage.Length == 0)
            {
                return new ExcelReaderException(typed, $"Native call failed with status {typed}.");
            }

            return new ExcelReaderException(typed, nativeMessage);
        }
    }
}
