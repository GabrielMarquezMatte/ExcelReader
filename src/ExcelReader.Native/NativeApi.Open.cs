using System.Text;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;

namespace ExcelReader.Native
{
    internal static partial class NativeApi
    {
        internal static int OpenFile(ReadOnlySpan<byte> utf8Path, int format, out NativeHandle? handle, NativeOpenOptions? options = null)
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
                IExcelRowReader reader = OpenReader(path, format, options ?? default);
                handle = new NativeHandle(reader);
                return NativeStatus.Ok;
            }
            catch (ExcelEncryptionException ex)
            {
                return MapEncryptionException(ex);
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int OpenMemory(ReadOnlySpan<byte> data, int format, out NativeHandle? handle, NativeOpenOptions? options = null)
        {
            handle = null;
            if (!IsKnownFormat(format))
            {
                return NativeStatus.InvalidArgument;
            }

            ClearLastError();
            try
            {
                byte[] copy = data.ToArray();
                IExcelRowReader reader = OpenReader(copy, format, options ?? default);
                handle = new NativeHandle(reader);
                return NativeStatus.Ok;
            }
            catch (ExcelEncryptionException ex)
            {
                return MapEncryptionException(ex);
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        private static int MapEncryptionException(ExcelEncryptionException ex)
        {
            SetLastError(ex.Message);
            return ex.Reason switch
            {
                ExcelEncryptionReason.PasswordRequired => NativeStatus.PasswordRequired,
                ExcelEncryptionReason.PasswordIncorrect => NativeStatus.PasswordIncorrect,
                _ => NativeStatus.Error,
            };
        }

        internal static int Close(NativeHandle? handle)
        {
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }

            ClearLastError();
            try
            {
                handle.Dispose();
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        private static IExcelRowReader OpenReader(string path, int format, NativeOpenOptions options)
        {
            return Excel.Open(path, MapFormat(format), options.ToExcelReaderOptions());
        }

        private static IExcelRowReader OpenReader(byte[] data, int format, NativeOpenOptions options)
        {
            return Excel.Open(data, MapFormat(format), options.ToExcelReaderOptions());
        }

        private static ExcelFileFormat MapFormat(int format)
        {
            return format switch
            {
                NativeFormat.Xls => ExcelFileFormat.Xls,
                NativeFormat.Xlsx => ExcelFileFormat.Xlsx,
                NativeFormat.Xlsb => ExcelFileFormat.Xlsb,
                NativeFormat.Csv => ExcelFileFormat.Csv,
                _ => ExcelFileFormat.Unknown,
            };
        }

        private static bool IsKnownFormat(int format)
        {
            return format is >= NativeFormat.Auto and <= NativeFormat.Csv;
        }

        internal static int OpenFileEx(ReadOnlySpan<byte> utf8Path, int format, NativeOpenOptionsRaw? rawOptions, out NativeHandle? handle)
        {
            handle = null;
            if (!TryDecodeOpenOptions(rawOptions, out NativeOpenOptions? options, out string? error))
            {
                SetLastError(error!);
                return NativeStatus.InvalidArgument;
            }
            return OpenFile(utf8Path, format, out handle, options);
        }

        internal static int OpenMemoryEx(ReadOnlySpan<byte> data, int format, NativeOpenOptionsRaw? rawOptions, out NativeHandle? handle)
        {
            handle = null;
            if (!TryDecodeOpenOptions(rawOptions, out NativeOpenOptions? options, out string? error))
            {
                SetLastError(error!);
                return NativeStatus.InvalidArgument;
            }
            return OpenMemory(data, format, out handle, options);
        }

        internal static bool TryDecodeOpenOptions(NativeOpenOptionsRaw? rawOptions, out NativeOpenOptions? options, out string? error)
        {
            if (rawOptions is not NativeOpenOptionsRaw raw)
            {
                options = null;
                error = null;
                return true;
            }

            bool ok = NativeOpenOptions.TryDecode(raw, out NativeOpenOptions decoded, out error);
            options = ok ? decoded : null;
            return ok;
        }
    }
}
