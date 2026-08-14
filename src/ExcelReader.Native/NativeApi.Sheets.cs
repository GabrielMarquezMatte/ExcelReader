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
                string name = handle.Reader.SheetName;
                int required = Encoding.UTF8.GetByteCount(name);
                length = required;
                if (buffer.Length < required)
                {
                    return NativeStatus.BufferTooSmall;
                }

                Encoding.UTF8.GetBytes(name, buffer);
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
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
                // Row enumeration is per-sheet: the old cursor points into the previous sheet's
                // buffers, so it is dropped and rebuilt on the next row request.
                handle.ResetRows();
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
