using System.Text;

namespace ExcelReader.Native
{
    internal static partial class NativeApi
    {
        internal static int SheetCount(NativeHandle? handle, out int count)
        {
            count = 0;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }

            ClearLastError();
            try
            {
                count = handle.Reader.SheetCount;
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int SheetName(NativeHandle? handle, Span<byte> buffer, out int length)
        {
            length = 0;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }

            ClearLastError();
            try
            {
                return CopyUtf8(handle.Reader.SheetName, buffer, out length);
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int SheetNameAt(NativeHandle? handle, int index, Span<byte> buffer, out int length)
        {
            length = 0;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }
            if (index < 0)
            {
                return NativeStatus.InvalidArgument;
            }

            ClearLastError();
            try
            {
                return CopyUtf8(handle.Reader.SheetNameAt(index), buffer, out length);
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        private static int CopyUtf8(string value, Span<byte> buffer, out int length)
        {
            int required = Encoding.UTF8.GetByteCount(value);
            length = required;
            if (buffer.Length < required)
            {
                return NativeStatus.BufferTooSmall;
            }

            Encoding.UTF8.GetBytes(value, buffer);
            return NativeStatus.Ok;
        }

        internal static int MoveToSheet(NativeHandle? handle, int index)
        {
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }

            ClearLastError();
            try
            {
                handle.Reader.MoveToSheet(index);
                handle.ResetRows();
                handle.FaultLiveSession("xl_move_to_sheet");
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int IsDate1904(NativeHandle? handle, out int flag)
        {
            flag = 0;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }

            ClearLastError();
            try
            {
                flag = handle.Reader.IsDate1904 ? 1 : 0;
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }
    }
}
