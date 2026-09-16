using System.Text;
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
            if (format == NativeFormat.Csv)
            {
                return OpenCsvFile(path, options);
            }

            ExcelReaderOptions excelOptions = options.ToExcelReaderOptions();
            return format switch
            {
                NativeFormat.Auto => Excel.Open(path, excelOptions),
                NativeFormat.Xlsx => Excel.FromXlsxFile(path, excelOptions),
                NativeFormat.Xlsb => Excel.FromXlsb(File.OpenRead(path), leaveOpen: false, excelOptions),
                _ => Excel.FromXls(File.OpenRead(path), leaveOpen: false, excelOptions),
            };
        }

        private static IExcelRowReader OpenReader(byte[] data, int format, NativeOpenOptions options)
        {
            if (format == NativeFormat.Csv)
            {
                return OpenCsvMemory(data, options);
            }

            ExcelReaderOptions excelOptions = options.ToExcelReaderOptions();
            return format switch
            {
                NativeFormat.Auto => Excel.Open(data, excelOptions),
                NativeFormat.Xlsx => Excel.FromXlsx(data, excelOptions),
                NativeFormat.Xlsb => Excel.FromXlsb(new MemoryStream(data, writable: false), leaveOpen: false, excelOptions),
                _ => Excel.FromXls(new MemoryStream(data, writable: false), leaveOpen: false, excelOptions),
            };
        }

        private static CsvReader OpenCsvFile(string path, NativeOpenOptions options)
        {
            CsvReaderOptions csvOptions = options.ToCsvReaderOptions();
            if (options.CsvSniffDialect)
            {
                csvOptions = csvOptions.WithDialect(Excel.SniffCsvDialectFromFile(path));
            }
            return Excel.FromCsv(File.OpenRead(path), leaveOpen: false, csvOptions);
        }

        private static CsvReader OpenCsvMemory(byte[] data, NativeOpenOptions options)
        {
            CsvReaderOptions csvOptions = options.ToCsvReaderOptions();
            if (options.CsvSniffDialect)
            {
                csvOptions = csvOptions.WithDialect(Excel.SniffCsvDialect(data));
            }
            return Excel.FromCsv(new MemoryStream(data, writable: false), leaveOpen: false, csvOptions);
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
