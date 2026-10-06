using System.Text;

namespace ExcelReader.Native.Reading
{
    internal static partial class ReadApi
    {
        internal static int SheetCount(NativeHandle? handle, out int count)
        {
            count = 0;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }

            NativeApi.ClearLastError();
            try
            {
                count = handle.Workbook.SheetCount;
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
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

            NativeApi.ClearLastError();
            try
            {
                return CopyUtf8(handle.Sheet.Name, buffer, out length);
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
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

            NativeApi.ClearLastError();
            try
            {
                return CopyUtf8(handle.Workbook.SheetAt(index).Name, buffer, out length);
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
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

            NativeApi.ClearLastError();
            try
            {
                handle.MoveToSheet(index);
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
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

            NativeApi.ClearLastError();
            try
            {
                flag = handle.Workbook.IsDate1904 ? 1 : 0;
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }
    }
}
