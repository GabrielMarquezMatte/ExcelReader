using ExcelReader.NetFramework.Interop;
using ExcelReader.NetFramework.Reader;
using ExcelReader.NetFramework.Writer.Internal;

namespace ExcelReader.NetFramework.Writer
{
    /// <summary>
    /// Writes a workbook cell by cell to a file or to memory. Nothing reaches the destination until
    /// <see cref="Close"/>, <see cref="CloseToArray"/> or <see cref="CloseTo"/>; disposing the writer first discards the output.
    /// </summary>
    public sealed unsafe class ExcelWorkbookWriter : IDisposable
    {
        private const long TicksPerMicrosecond = TimeSpan.TicksPerMillisecond / 1000;
        private static readonly long EpochTicks = new DateTime(1970, 1, 1).Ticks;

        private readonly WriterHandle _handle;
        private readonly bool _toMemory;
        private bool _finished;
        private bool _sheetOpen;

        private ExcelWorkbookWriter(WriterHandle handle, bool toMemory)
        {
            _handle = handle;
            _toMemory = toMemory;
        }

        /// <summary>Creates a writer that replaces the file at <paramref name="path"/> when it is closed.</summary>
        /// <param name="path">The destination file.</param>
        /// <param name="options">The write options, or null for the defaults.</param>
        /// <returns>The open writer.</returns>
        /// <exception cref="ArgumentException">The format cannot be inferred from the extension and <see cref="ExcelWriteOptions.Format"/> is not set.</exception>
        /// <exception cref="ArgumentOutOfRangeException">An option is out of range.</exception>
        /// <exception cref="NotSupportedException"><see cref="ExcelWriteOptions.Password"/> is set.</exception>
        public static ExcelWorkbookWriter Create(string path, ExcelWriteOptions? options = null)
        {
            if (path is null)
            {
                throw new ArgumentNullException(nameof(path));
            }

            NativeLoader.EnsureLoaded();
            ExcelFormat format = WriteFormats.FromPath(path, options?.Format);
            XlWriteOptions native = Prepare(options);
            byte[] pathBytes = Utf8.NonNullBytes(path, out int pathLength);
            IntPtr raw = IntPtr.Zero;
            fixed (byte* pathPointer = pathBytes)
            {
                NativeError.ThrowIfFailed(NativeMethods.xl_open_write_handle(pathPointer, pathLength, (int)format, &native, &raw));
            }

            return new ExcelWorkbookWriter(new WriterHandle(raw), false);
        }

        /// <summary>Creates a writer that builds the workbook in memory.</summary>
        /// <param name="format">The format to write.</param>
        /// <param name="options">The write options, or null for the defaults.</param>
        /// <returns>The open writer; finish it with <see cref="CloseToArray"/> or <see cref="CloseTo"/>.</returns>
        /// <exception cref="ArgumentException"><paramref name="format"/> is <see cref="ExcelFormat.Auto"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException">An option is out of range.</exception>
        /// <exception cref="NotSupportedException"><see cref="ExcelWriteOptions.Password"/> is set.</exception>
        public static ExcelWorkbookWriter CreateInMemory(ExcelFormat format, ExcelWriteOptions? options = null)
        {
            NativeLoader.EnsureLoaded();
            WriteFormats.Require(format);
            XlWriteOptions native = Prepare(options);
            IntPtr raw = IntPtr.Zero;
            NativeError.ThrowIfFailed(NativeMethods.xl_open_write_handle_to_memory((int)format, &native, &raw));
            return new ExcelWorkbookWriter(new WriterHandle(raw), true);
        }

        /// <summary>Starts a sheet.</summary>
        /// <param name="name">The sheet name: at most 31 characters, unique, without the characters Excel forbids.</param>
        /// <exception cref="ExcelReaderException">The name is invalid or a sheet is already open.</exception>
        public void StartSheet(string name)
        {
            if (name is null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            byte[] bytes = Utf8.NonNullBytes(name, out int length);
            using Lease lease = Acquire();
            fixed (byte* namePointer = bytes)
            {
                NativeError.ThrowIfFailed(NativeMethods.xl_start_sheet(lease.Raw, namePointer, length));
            }

            _sheetOpen = true;
        }

        /// <summary>Starts a row in the open sheet.</summary>
        public void StartRow()
        {
            using Lease lease = Acquire();
            NativeError.ThrowIfFailed(NativeMethods.xl_start_row(lease.Raw));
        }

        /// <summary>Ends the open row.</summary>
        public void EndRow()
        {
            using Lease lease = Acquire();
            NativeError.ThrowIfFailed(NativeMethods.xl_end_row(lease.Raw));
        }

        /// <summary>Ends the open sheet.</summary>
        public void EndSheet()
        {
            using Lease lease = Acquire();
            NativeError.ThrowIfFailed(NativeMethods.xl_end_sheet(lease.Raw));
            _sheetOpen = false;
        }

        /// <summary>
        /// Writes a whole sheet from <paramref name="rows"/>: a header row, then one row per item. Columns are the public readable
        /// properties of <typeparamref name="T"/> without <see cref="Parser.ExcelIgnoreAttribute"/>, in declaration order (base type first);
        /// each header is the first <see cref="Parser.ExcelColumnAttribute"/> name, or the property name.
        /// </summary>
        /// <typeparam name="T">The row type. Supported property types: string, long, int, short, byte, double, float, decimal, bool,
        /// <see cref="DateTime"/> (written as a date and time, its <see cref="DateTime.Kind"/> ignored), <see cref="TimeSpan"/> (a time of day), and nullable forms of those.</typeparam>
        /// <param name="name">The sheet name: at most 31 characters, unique, without the characters Excel forbids.</param>
        /// <param name="rows">The rows, enumerated once.</param>
        /// <param name="table">Writes the sheet as an Excel table (XLSX and XLSB only), or null for a plain sheet.</param>
        /// <exception cref="ArgumentException"><paramref name="table"/> has no <see cref="ExcelTableOptions.Name"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException">A <see cref="TimeSpan"/> value is negative or at least one day.</exception>
        /// <exception cref="InvalidOperationException">A sheet started with <see cref="StartSheet"/> is still open.</exception>
        /// <exception cref="NotSupportedException">A property of <typeparamref name="T"/> has an unsupported type.</exception>
        /// <exception cref="ExcelReaderException">The name or table is invalid, or the format cannot hold a table.</exception>
        /// <exception cref="ObjectDisposedException">The writer is already closed, aborted or disposed.</exception>
        public void WriteSheet<T>(string name, IEnumerable<T> rows, ExcelTableOptions? table = null)
        {
            if (name is null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            if (rows is null)
            {
                throw new ArgumentNullException(nameof(rows));
            }

            if (table is not null && string.IsNullOrEmpty(table.Name))
            {
                throw new ArgumentException("An Excel table needs a name.", nameof(table));
            }

            ThrowIfFinished();
            if (_sheetOpen)
            {
                throw new InvalidOperationException("A sheet is open; call EndSheet before writing a whole sheet.");
            }

            WriteMap<T> map = WriteMap<T>.Instance;
            T[] items = rows.ToArray();
            byte[] nameBytes = Utf8.NonNullBytes(name, out int nameLength);
            using NativeWriteTable native = NativeWriteTable.Create(map, items);
            using Lease lease = Acquire();
            fixed (byte* namePointer = nameBytes)
            {
                if (table is null)
                {
                    NativeError.ThrowIfFailed(NativeMethods.xl_write_sheet(lease.Raw, namePointer, nameLength, native.Specs, native.Table, null));
                    return;
                }

                WriteTableSheet(lease.Raw, namePointer, nameLength, native, table);
            }
        }

        /// <summary>Writes a text cell; null writes a blank cell.</summary>
        /// <param name="value">The text.</param>
        public void Write(string? value)
        {
            using Lease lease = Acquire();
            if (value is null)
            {
                NativeError.ThrowIfFailed(NativeMethods.xl_write_string(lease.Raw, null, 0));
                return;
            }

            byte[] bytes = Utf8.NonNullBytes(value, out int length);
            fixed (byte* valuePointer = bytes)
            {
                NativeError.ThrowIfFailed(NativeMethods.xl_write_string(lease.Raw, valuePointer, length));
            }
        }

        /// <summary>Writes an integer cell.</summary>
        /// <param name="value">The integer.</param>
        public void Write(long value)
        {
            using Lease lease = Acquire();
            NativeError.ThrowIfFailed(NativeMethods.xl_write_int64(lease.Raw, value));
        }

        /// <summary>Writes a floating-point cell.</summary>
        /// <param name="value">The number.</param>
        public void Write(double value)
        {
            using Lease lease = Acquire();
            NativeError.ThrowIfFailed(NativeMethods.xl_write_float64(lease.Raw, value));
        }

        /// <summary>Writes a decimal as a floating-point cell.</summary>
        /// <param name="value">The number.</param>
        public void Write(decimal value)
        {
            Write((double)value);
        }

        /// <summary>Writes a boolean cell.</summary>
        /// <param name="value">The boolean.</param>
        public void Write(bool value)
        {
            using Lease lease = Acquire();
            NativeError.ThrowIfFailed(NativeMethods.xl_write_bool(lease.Raw, value ? 1 : 0));
        }

        /// <summary>Writes the date part of <paramref name="value"/> as a date cell.</summary>
        /// <param name="value">The date; its time of day is ignored.</param>
        public void WriteDate(DateTime value)
        {
            int days = checked((int)((value.Date.Ticks - EpochTicks) / TimeSpan.TicksPerDay));
            using Lease lease = Acquire();
            NativeError.ThrowIfFailed(NativeMethods.xl_write_date(lease.Raw, days));
        }

        /// <summary>Writes a time-of-day cell.</summary>
        /// <param name="value">The time of day, from zero up to but excluding one day.</param>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="value"/> is negative or at least one day.</exception>
        public void WriteTime(TimeSpan value)
        {
            if (value < TimeSpan.Zero || value >= TimeSpan.FromDays(1))
            {
                throw new ArgumentOutOfRangeException(nameof(value), value, "A time of day is at least zero and less than one day.");
            }

            using Lease lease = Acquire();
            NativeError.ThrowIfFailed(NativeMethods.xl_write_time(lease.Raw, value.Ticks / TicksPerMicrosecond));
        }

        /// <summary>Writes a date-and-time cell. The <see cref="DateTime.Kind"/> is ignored.</summary>
        /// <param name="value">The date and time.</param>
        public void WriteDateTime(DateTime value)
        {
            using Lease lease = Acquire();
            NativeError.ThrowIfFailed(NativeMethods.xl_write_timestamp(lease.Raw, (value.Ticks - EpochTicks) / TicksPerMicrosecond));
        }

        /// <summary>Writes a blank cell.</summary>
        public void WriteEmpty()
        {
            Write((string?)null);
        }

        /// <summary>Finishes a file writer and replaces the destination file.</summary>
        /// <exception cref="InvalidOperationException">The writer was created in memory.</exception>
        /// <exception cref="ObjectDisposedException">The writer is already closed, aborted or disposed.</exception>
        public void Close()
        {
            ThrowIfFinished();
            if (_toMemory)
            {
                throw new InvalidOperationException("A memory writer finishes with CloseToArray or CloseTo.");
            }

            CloseNative();
        }

        /// <summary>Finishes a memory writer and returns the workbook bytes.</summary>
        /// <returns>The workbook bytes.</returns>
        /// <exception cref="InvalidOperationException">The writer was created for a file.</exception>
        /// <exception cref="ObjectDisposedException">The writer is already closed, aborted or disposed.</exception>
        public byte[] CloseToArray()
        {
            ThrowIfFinished();
            if (!_toMemory)
            {
                throw new InvalidOperationException("A file writer finishes with Close.");
            }

            byte[] bytes = TakeBytes();
            CloseNative();
            return bytes;
        }

        /// <summary>Finishes a memory writer and copies the workbook bytes to <paramref name="destination"/>, which stays open.</summary>
        /// <param name="destination">The stream to write to.</param>
        /// <exception cref="InvalidOperationException">The writer was created for a file.</exception>
        /// <exception cref="ObjectDisposedException">The writer is already closed, aborted or disposed.</exception>
        public void CloseTo(Stream destination)
        {
            if (destination is null)
            {
                throw new ArgumentNullException(nameof(destination));
            }

            byte[] bytes = CloseToArray();
            destination.Write(bytes, 0, bytes.Length);
        }

        /// <summary>Discards everything written; an existing destination file is left as it was.</summary>
        /// <exception cref="ObjectDisposedException">The writer is already closed, aborted or disposed.</exception>
        public void Abort()
        {
            ThrowIfFinished();
            _finished = true;
            _handle.Dispose();
        }

        /// <summary>Aborts the writer when it was not closed.</summary>
        public void Dispose()
        {
            if (_finished)
            {
                return;
            }

            Abort();
        }

        private static XlWriteOptions Prepare(ExcelWriteOptions? options)
        {
            XlWriteOptions native = WriteFormats.ToNative(options);
            if (options?.Password is not null)
            {
                throw new NotSupportedException("Password output is not supported yet.");
            }

            return native;
        }

        private static void WriteTableSheet(IntPtr raw, byte* name, int nameLength, NativeWriteTable native, ExcelTableOptions table)
        {
            byte[] tableName = Utf8.NonNullBytes(table.Name, out int tableNameLength);
            byte[]? style = null;
            int styleLength = 0;
            if (table.Style is not null)
            {
                style = Utf8.NonNullBytes(table.Style, out styleLength);
            }

            fixed (byte* tableNamePointer = tableName)
            fixed (byte* stylePointer = style)
            {
                XlSheetOptions options = new XlSheetOptions
                {
                    StructSize = sizeof(XlSheetOptions),
                    TableNameLen = tableNameLength,
                    TableName = tableNamePointer,
                    TableStyleLen = styleLength,
                    TableStyle = stylePointer,
                    TableFlags = table.Flags(),
                };
                NativeError.ThrowIfFailed(NativeMethods.xl_write_sheet(raw, name, nameLength, native.Specs, native.Table, &options));
            }
        }

        private byte[] TakeBytes()
        {
            using Lease lease = Acquire();
            XlBuffer buffer = default;
            NativeError.ThrowIfFailed(NativeMethods.xl_write_handle_bytes(lease.Raw, &buffer));
            return NativeBuffers.TakeArray(ref buffer);
        }

        private void CloseNative()
        {
            try
            {
                using Lease lease = Acquire();
                NativeError.ThrowIfFailed(NativeMethods.xl_close_write_handle(lease.Raw));
            }
            finally
            {
                _finished = true;
                _handle.MarkClosed();
                _handle.Dispose();
            }
        }

        private Lease Acquire()
        {
            ThrowIfFinished();
            return new Lease(_handle);
        }

        private void ThrowIfFinished()
        {
            if (_finished)
            {
                throw new ObjectDisposedException(nameof(ExcelWorkbookWriter));
            }
        }

        private readonly struct Lease : IDisposable
        {
            private readonly WriterHandle _owner;

            internal Lease(WriterHandle owner)
            {
                bool added = false;
                owner.DangerousAddRef(ref added);
                _owner = owner;
            }

            internal IntPtr Raw
            {
                get
                {
                    return _owner.DangerousGetHandle();
                }
            }

            public void Dispose()
            {
                _owner.DangerousRelease();
            }
        }
    }
}
