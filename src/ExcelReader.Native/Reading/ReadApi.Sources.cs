using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Sources;

namespace ExcelReader.Native.Reading
{
    internal static partial class ReadApi
    {
        private const int DefaultSourceBlockSize = 4 * 1024 * 1024;
        private const long DefaultSourceCacheBytes = 64L * 1024 * 1024;

        internal static int OpenSource(ByteSource callback, int format, NativeOpenOptionsRaw? rawOptions, out NativeHandle? handle)
        {
            handle = null;
            bool opened = false;
            try
            {
                if (!TryDecodeOpenOptions(rawOptions, out NativeOpenOptions? decoded, out string? error))
                {
                    NativeApi.SetLastError(error!);
                    return NativeStatus.InvalidArgument;
                }
                if (!IsKnownFormat(format))
                {
                    NativeApi.SetLastError($"format must be one of the XL_FORMAT_* values; got {format}.");
                    return NativeStatus.InvalidArgument;
                }
                NativeOpenOptions options = decoded ?? default;
                int blockSize = options.SourceBlockSize ?? DefaultSourceBlockSize;
                ByteSource source = blockSize < 0
                    ? callback
                    : new CachedByteSource(callback, blockSize, options.SourceCacheBytes ?? DefaultSourceCacheBytes);
                IExcelWorkbook workbook = Excel.Open(new SourceStream(source), MapFormat(format), leaveOpen: false, options.ToExcelReaderOptions());
                handle = new NativeHandle(workbook);
                opened = true;
                return NativeStatus.Ok;
            }
            catch (ExcelEncryptionException ex)
            {
                return MapEncryptionException(ex);
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
            finally
            {
                if (!opened)
                {
                    callback.Dispose();
                }
            }
        }

        internal static int OpenStream(CallbackReadStream stream, int format, NativeOpenOptionsRaw? rawOptions, out NativeHandle? handle)
        {
            handle = null;
            bool opened = false;
            try
            {
                if (!TryDecodeOpenOptions(rawOptions, out NativeOpenOptions? decoded, out string? error))
                {
                    NativeApi.SetLastError(error!);
                    return NativeStatus.InvalidArgument;
                }
                if (!IsKnownFormat(format))
                {
                    NativeApi.SetLastError($"format must be one of the XL_FORMAT_* values; got {format}.");
                    return NativeStatus.InvalidArgument;
                }
                NativeOpenOptions options = decoded ?? default;
                ExcelReaderOptions readerOptions = options.ToExcelReaderOptions();
                IExcelWorkbook workbook;
                if (format == NativeFormat.Csv)
                {
                    workbook = Excel.FromCsv(stream, leaveOpen: false, readerOptions.Csv);
                }
                else
                {
                    ReadOnlyMemory<byte> data = StreamBuffer.ReadAll(stream, options.MaxBufferedBytes ?? Array.MaxLength);
                    stream.Dispose();
                    workbook = Excel.Open(data, MapFormat(format), readerOptions);
                }
                handle = new NativeHandle(workbook);
                opened = true;
                return NativeStatus.Ok;
            }
            catch (ExcelEncryptionException ex)
            {
                return MapEncryptionException(ex);
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
            finally
            {
                if (!opened)
                {
                    stream.Dispose();
                }
            }
        }
    }
}
