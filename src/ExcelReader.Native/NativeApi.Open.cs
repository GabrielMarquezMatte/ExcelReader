using System.Text;
using ExcelReader.Core.Reader;

namespace ExcelReader.Native
{
    internal static partial class NativeApi
    {
        internal static int OpenFile(ReadOnlySpan<byte> utf8Path, int format, out NativeHandle? handle)
        {
            handle = null;
            if (!IsKnownFormat(format))
            {
                return NativeStatus.InvalidArgument;
            }

            ClearLastError();
            try
            {
                string path = Encoding.UTF8.GetString(utf8Path);
                IExcelRowReader reader = OpenReader(path, format);
                handle = new NativeHandle(reader);
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int OpenMemory(ReadOnlySpan<byte> data, int format, out NativeHandle? handle)
        {
            handle = null;
            if (!IsKnownFormat(format))
            {
                return NativeStatus.InvalidArgument;
            }

            ClearLastError();
            try
            {
                // Copied on purpose: the ABI promises the caller may free its buffer immediately,
                // and the readers keep referencing this memory for the handle's whole lifetime.
                byte[] copy = data.ToArray();
                IExcelRowReader reader = OpenReader(copy, format);
                handle = new NativeHandle(reader);
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int Close(NativeHandle? handle)
        {
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }

            handle.Dispose();
            return NativeStatus.Ok;
        }

        private static IExcelRowReader OpenReader(string path, int format)
        {
            return format switch
            {
                NativeFormat.Auto => Excel.Open(path),
                NativeFormat.Xlsx => Excel.FromXlsxFile(path),
                NativeFormat.Xlsb => Excel.FromXlsb(File.OpenRead(path), leaveOpen: false),
                NativeFormat.Xls => Excel.FromXls(File.OpenRead(path), leaveOpen: false),
                _ => Excel.FromCsv(File.OpenRead(path), leaveOpen: false),
            };
        }

        private static IExcelRowReader OpenReader(byte[] data, int format)
        {
            return format switch
            {
                NativeFormat.Auto => Excel.Open(data),
                NativeFormat.Xlsx => Excel.FromXlsx(data),
                NativeFormat.Xlsb => Excel.FromXlsb(new MemoryStream(data, writable: false), leaveOpen: false),
                NativeFormat.Xls => Excel.FromXls(new MemoryStream(data, writable: false), leaveOpen: false),
                _ => Excel.FromCsv(new MemoryStream(data, writable: false), leaveOpen: false),
            };
        }

        private static bool IsKnownFormat(int format)
        {
            return format is >= NativeFormat.Auto and <= NativeFormat.Csv;
        }
    }
}
