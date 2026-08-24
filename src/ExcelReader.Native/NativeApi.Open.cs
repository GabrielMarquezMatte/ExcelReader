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
                // Copied on purpose: the ABI promises the caller may free its buffer immediately,
                // and the readers keep referencing this memory for the handle's whole lifetime.
                byte[] copy = data.ToArray();
                IExcelRowReader reader = OpenReader(copy, format, options ?? default);
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

            // Every other NativeApi entry point wraps its body this way; this one didn't, so an
            // IOException from the underlying FileStream's Dispose (e.g. the source volume went
            // away between open and close) unwound straight through the [UnmanagedCallersOnly]
            // frame instead of coming back as XL_ERROR - which is a fail-fast/abort for the native
            // caller, uncatchable in C/C++/Rust/Python. The id is already retired by the time this
            // runs (Exports.Close unregisters before calling here), so a failure here just means the
            // caller learns about it via the return code instead of it staying invisible.
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
                // Wires up Excel.SniffCsvDialectFromFile, which nothing in the open path calls
                // otherwise — mirrors exactly what a C# caller would write by hand.
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

        /// <summary>The xl_open_file_ex entry point's logic: decodes <paramref name="rawOptions"/> (a null
        /// value means the caller passed a NULL options pointer — identical to xl_open_file) and, if valid,
        /// opens exactly as <see cref="OpenFile"/> does with it applied.</summary>
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

        /// <summary>The xl_open_memory_ex entry point's logic — the in-memory twin of <see cref="OpenFileEx"/>.</summary>
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

        /// <summary>Decodes and validates a raw <c>xl_open_options</c> struct, if the caller passed one.
        /// A null <paramref name="rawOptions"/> means "no struct passed" (xl_open_file/xl_open_memory, or
        /// an _ex call with a NULL options pointer) and decodes to <see langword="null"/>, meaning "use
        /// every library default" — the same as never having called an _ex function at all.</summary>
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
