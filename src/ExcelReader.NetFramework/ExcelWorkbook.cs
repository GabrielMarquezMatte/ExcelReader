using System.Collections.ObjectModel;
using ExcelReader.NetFramework.Interop;

namespace ExcelReader.NetFramework
{
    /// <summary>An open workbook. Dispose it to release the native handle.</summary>
    public sealed unsafe class ExcelWorkbook : IDisposable
    {
        private const int StackNameBytes = 256;
        private const int MaxAscii = 0x7F;

        private readonly WorkbookHandle _handle;
        private IReadOnlyList<ExcelSheetInfo>? _sheets;
        private IReadOnlyList<ExcelTableInfo>? _tables;

        private delegate int NameReader(IntPtr handle, int index, byte* buffer, int capacity, int* outLen);

        private delegate int IndexLookup(IntPtr handle, byte* name, int nameLen, int* outIndex);

        private ExcelWorkbook(WorkbookHandle handle)
        {
            _handle = handle;
        }

        internal WorkbookHandle Handle
        {
            get
            {
                ThrowIfDisposed();
                return _handle;
            }
        }

        /// <summary>The sheets of the workbook, in workbook order.</summary>
        public IReadOnlyList<ExcelSheetInfo> Sheets
        {
            get
            {
                ThrowIfDisposed();
                _sheets ??= ReadSheets();
                return _sheets;
            }
        }

        /// <summary>The Excel tables of the workbook.</summary>
        public IReadOnlyList<ExcelTableInfo> Tables
        {
            get
            {
                ThrowIfDisposed();
                _tables ??= ReadTables();
                return _tables;
            }
        }

        /// <summary>Whether dates use the 1904 epoch.</summary>
        public bool IsDate1904
        {
            get
            {
                ThrowIfDisposed();
                int flag;
                NativeError.ThrowIfFailed(NativeMethods.xl_is_date1904(_handle.DangerousGetHandle(), &flag));
                return flag != 0;
            }
        }

        /// <summary>Opens a workbook from a file.</summary>
        /// <param name="path">The file path.</param>
        /// <param name="options">The open options, or null for the defaults.</param>
        /// <returns>The open workbook.</returns>
        public static ExcelWorkbook Open(string path, ExcelOpenOptions? options = null)
        {
            if (path is null)
            {
                throw new ArgumentNullException(nameof(path));
            }

            return OpenNative(Utf8.Bytes(path), null, null, options);
        }

        /// <summary>Opens a workbook from memory. The data is copied by the native library.</summary>
        /// <param name="data">The workbook bytes.</param>
        /// <param name="options">The open options, or null for the defaults.</param>
        /// <returns>The open workbook.</returns>
        public static ExcelWorkbook Open(byte[] data, ExcelOpenOptions? options = null)
        {
            if (data is null)
            {
                throw new ArgumentNullException(nameof(data));
            }

            return OpenNative(null, data, null, options);
        }

        /// <summary>
        /// Opens a workbook from a stream. The workbook does not own the stream: the native library stops using it
        /// once the workbook is open, except for CSV, which reads it until the workbook is disposed.
        /// </summary>
        /// <param name="stream">A readable stream positioned at the start of the workbook.</param>
        /// <param name="options">The open options, or null for the defaults.</param>
        /// <returns>The open workbook.</returns>
        public static ExcelWorkbook Open(Stream stream, ExcelOpenOptions? options = null)
        {
            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            return OpenNative(null, null, stream, options);
        }

        /// <summary>Gets a sheet by exact, case-sensitive name.</summary>
        /// <param name="name">The sheet name.</param>
        /// <returns>The sheet.</returns>
        public ExcelSheet Sheet(string name)
        {
            return new ExcelSheet(this, Lookup(name, NativeMethods.xl_sheet_index, "sheet"));
        }

        /// <summary>Gets a sheet by zero-based index.</summary>
        /// <param name="index">The sheet index.</param>
        /// <returns>The sheet.</returns>
        public ExcelSheet Sheet(int index)
        {
            if (index < 0 || index >= Sheets.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index), index, "The workbook has no sheet at this index.");
            }

            return new ExcelSheet(this, index);
        }

        /// <summary>Gets a table by name, ignoring case. The table is read through its own range.</summary>
        /// <param name="name">The table name.</param>
        /// <returns>The table as a sheet.</returns>
        public ExcelSheet Table(string name)
        {
            return new ExcelSheet(this, Xl.TableSheet(Lookup(name, NativeMethods.xl_table_index, "table")));
        }

        /// <summary>Releases the native handle. Calling it again does nothing.</summary>
        public void Dispose()
        {
            _handle.Dispose();
        }

        private static ExcelWorkbook OpenNative(byte[]? path, byte[]? data, Stream? stream, ExcelOpenOptions? options)
        {
            byte[] password = Array.Empty<byte>();
            if (options?.Password is not null)
            {
                password = Utf8.Bytes(options.Password);
            }

            XlOpenOptions native = ToNative(options);
            native.PasswordLen = password.Length;
            IntPtr handle = IntPtr.Zero;
            int format = (int)(options?.Format ?? ExcelFormat.Auto);
            fixed (byte* passwordPointer = password)
            fixed (byte* pathPointer = path)
            fixed (byte* dataPointer = data)
            {
                native.Password = passwordPointer;
                if (stream is not null)
                {
                    return OpenStream(stream, format, &native);
                }

                if (path is null)
                {
                    return Wrap(NativeMethods.xl_open_memory(dataPointer, data!.Length, format, &native, &handle), handle);
                }

                return Wrap(NativeMethods.xl_open_file(pathPointer, path.Length, format, &native, &handle), handle);
            }
        }

        private static ExcelWorkbook OpenStream(Stream stream, int format, XlOpenOptions* options)
        {
            XlStream source = StreamSource.Create(stream);
            IntPtr handle = IntPtr.Zero;
            return Wrap(NativeMethods.xl_open_stream(&source, format, options, &handle), handle);
        }

        private static ExcelWorkbook Wrap(int status, IntPtr handle)
        {
            NativeError.ThrowIfFailed(status);
            return new ExcelWorkbook(new WorkbookHandle(handle));
        }

        private static XlOpenOptions ToNative(ExcelOpenOptions? options)
        {
            XlOpenOptions native = new XlOpenOptions { StructSize = sizeof(XlOpenOptions) };
            if (options is null)
            {
                return native;
            }

            native.CsvDelimiter = AsciiOrDefault(options.CsvDelimiter, nameof(options.CsvDelimiter));
            native.CsvQuote = AsciiOrDefault(options.CsvQuote, nameof(options.CsvQuote));
            if (options.CsvDetectBom is bool detectBom)
            {
                native.CsvDetectBom = detectBom ? Xl.XL_OPT_TRUE : Xl.XL_OPT_FALSE;
            }

            return native;
        }

        private static int AsciiOrDefault(char? value, string name)
        {
            if (value is null)
            {
                return 0;
            }

            if (value > MaxAscii)
            {
                throw new ArgumentOutOfRangeException(name, value, "The native library takes a single-byte character.");
            }

            return value.Value;
        }

        private int Lookup(string name, IndexLookup lookup, string kind)
        {
            ThrowIfDisposed();
            if (name is null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            byte[] bytes = Utf8.Bytes(name);
            int index = -1;
            int status;
            fixed (byte* pointer = bytes)
            {
                status = lookup(_handle.DangerousGetHandle(), pointer, bytes.Length, &index);
            }

            NativeError.ThrowIfFailed(status);
            if (index < 0)
            {
                throw new KeyNotFoundException($"The workbook has no {kind} named '{name}'.");
            }

            return index;
        }

        private ReadOnlyCollection<ExcelSheetInfo> ReadSheets()
        {
            IntPtr handle = _handle.DangerousGetHandle();
            int count = 0;
            NativeError.ThrowIfFailed(NativeMethods.xl_sheet_count(handle, &count));
            List<ExcelSheetInfo> sheets = new List<ExcelSheetInfo>(count);
            for (int index = 0; index < count; index++)
            {
                int visibility = 0;
                NativeError.ThrowIfFailed(NativeMethods.xl_sheet_visibility_at(handle, index, &visibility));
                string name = ReadName(NativeMethods.xl_sheet_name_at, index);
                sheets.Add(new ExcelSheetInfo(index, name, (ExcelSheetVisibility)visibility));
            }

            return sheets.AsReadOnly();
        }

        private ReadOnlyCollection<ExcelTableInfo> ReadTables()
        {
            IntPtr handle = _handle.DangerousGetHandle();
            int count = 0;
            NativeError.ThrowIfFailed(NativeMethods.xl_table_count(handle, &count));
            List<ExcelTableInfo> tables = new List<ExcelTableInfo>(count);
            for (int index = 0; index < count; index++)
            {
                XlTableInfo info = default;
                NativeError.ThrowIfFailed(NativeMethods.xl_table_info_at(handle, index, &info));
                string name = ReadName(NativeMethods.xl_table_name_at, index);
                tables.Add(new ExcelTableInfo(index, name, info.Sheet, info.FirstRow, info.LastRow, info.FirstColumn, info.LastColumn, info.HeaderRowCount, info.TotalsRowCount));
            }

            return tables.AsReadOnly();
        }

        private string ReadName(NameReader reader, int index)
        {
            IntPtr handle = _handle.DangerousGetHandle();
            byte* stack = stackalloc byte[StackNameBytes];
            int length = 0;
            int status = reader(handle, index, stack, StackNameBytes, &length);
            if (status == (int)ExcelStatus.Ok)
            {
                return Utf8.Decode(stack, length) ?? string.Empty;
            }

            if (status != (int)ExcelStatus.BufferTooSmall)
            {
                throw NativeError.Create(status);
            }

            byte[] large = new byte[length];
            fixed (byte* pointer = large)
            {
                NativeError.ThrowIfFailed(reader(handle, index, pointer, large.Length, &length));
                return Utf8.Decode(pointer, length) ?? string.Empty;
            }
        }

        private void ThrowIfDisposed()
        {
            if (_handle.IsClosed)
            {
                throw new ObjectDisposedException(nameof(ExcelWorkbook));
            }
        }
    }
}
