using System.Text;
using ExcelReader.Core.Reader;

namespace ExcelReader.Native.Reading
{
    internal static partial class ReadApi
    {
        internal static int ResolveSheet(NativeHandle handle, int sheet, out IExcelSheet? resolved)
        {
            resolved = null;
            if (sheet < 0)
            {
                NativeApi.SetLastError($"sheet must be a zero-based index; got {sheet}.");
                return NativeStatus.InvalidArgument;
            }

            try
            {
                resolved = handle.Workbook.SheetAt(sheet);
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

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

        internal static int SheetVisibilityAt(NativeHandle? handle, int index, out int visibility)
        {
            visibility = 0;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }

            NativeApi.ClearLastError();
            int status = ResolveSheet(handle, index, out IExcelSheet? sheet);
            if (status == NativeStatus.Ok)
            {
                visibility = (int)sheet!.Visibility;
            }
            return status;
        }

        internal static int SheetIndex(NativeHandle? handle, ReadOnlySpan<byte> utf8Name, out int index)
        {
            index = -1;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }

            NativeApi.ClearLastError();
            try
            {
                if (handle.Workbook.TryGetSheet(Encoding.UTF8.GetString(utf8Name), out IExcelSheet? sheet))
                {
                    index = sheet!.Index;
                }
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
