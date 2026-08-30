using System.Diagnostics.CodeAnalysis;
using System.Text;
using ExcelReader.Core.Writer;
using ExcelReader.Native.Writer;

namespace ExcelReader.Native
{
    internal static unsafe partial class NativeApi
    {
        /// <summary>Style index 1 is always the builtin date style (see IWorkbookWriter.AddStyle).</summary>
        private const int BuiltinDateStyleId = 1;

        // Named distinctly from NativeApi.Typed.cs's field of the same value to avoid S3218 (field
        // shadows an outer-class member).
        internal static readonly int WriteUnixEpochDayNumber = new DateOnly(1970, 1, 1).DayNumber;

        /// <summary>
        /// Writes <paramref name="table"/> to <paramref name="path"/> as one sheet. Mirrors
        /// <see cref="ParseTyped"/> in reverse and consumes the exact structs it produces.
        /// </summary>
        /// <remarks>
        /// Every buffer reachable from <paramref name="table"/> is borrowed, never copied and never
        /// freed here. Validation runs to completion before the file is created, so a rejected call
        /// leaves nothing behind on disk.
        /// </remarks>
        internal static int WriteTyped(ReadOnlySpan<byte> path, int format, NativeColumnSpec[] specs, NativeTable table, NativeWriteOptions options)
        {
            if (path.IsEmpty)
            {
                SetLastError("xl_write_typed needs a non-empty path.");
                return NativeStatus.InvalidArgument;
            }
            if (!IsWritableFormat(format))
            {
                SetLastError($"xl_write_typed needs an explicit format (XLS/XLSX/XLSB/CSV); got format {format}.");
                return NativeStatus.InvalidArgument;
            }
            if (!TryValidateWriteTable(specs, table, out bool hasHeader, out string? validationError))
            {
                SetLastError(validationError);
                return NativeStatus.InvalidArgument;
            }

            ClearLastError();
            try
            {
                string filePath = Encoding.UTF8.GetString(path);
                string sheetName = options.SheetName ?? "Sheet1";
                using FileStream stream = new(filePath, FileMode.Create, FileAccess.Write, FileShare.None);
                WriteToStream(stream, format, specs, table, options, sheetName, hasHeader);
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        /// <summary>
        /// Same as <see cref="WriteTyped"/>, except the result is returned as bytes rather than
        /// written to a path — everything else about validation and content is identical.
        /// </summary>
        internal static int WriteTypedToMemory(int format, NativeColumnSpec[] specs, NativeTable table, NativeWriteOptions options, out byte[]? bytes)
        {
            bytes = null;
            if (!IsWritableFormat(format))
            {
                SetLastError($"xl_write_typed_to_memory needs an explicit format (XLS/XLSX/XLSB/CSV); got format {format}.");
                return NativeStatus.InvalidArgument;
            }
            if (!TryValidateWriteTable(specs, table, out bool hasHeader, out string? validationError))
            {
                SetLastError(validationError);
                return NativeStatus.InvalidArgument;
            }

            ClearLastError();
            try
            {
                string sheetName = options.SheetName ?? "Sheet1";
                using MemoryStream stream = new();
                WriteToStream(stream, format, specs, table, options, sheetName, hasHeader);
                bytes = stream.ToArray();
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int OpenWriteHandle(ReadOnlySpan<byte> path, int format, NativeWriteOptions options, out NativeWriterHandle? handle)
        {
            handle = null;
            if (path.IsEmpty)
            {
                SetLastError("xl_open_write_handle needs a non-empty path.");
                return NativeStatus.InvalidArgument;
            }
            if (!IsWritableFormat(format))
            {
                SetLastError($"xl_open_write_handle needs an explicit format (XLS/XLSX/XLSB/CSV); got format {format}.");
                return NativeStatus.InvalidArgument;
            }
            ClearLastError();
            string filePath = Encoding.UTF8.GetString(path);
            FileStream stream = new(filePath, FileMode.Create, FileAccess.Write, FileShare.None);
            try
            {
                handle = NativeWriterHandle.Create(stream, format, options);
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                // NativeWriterHandle.Create has not taken ownership of stream if it throws (e.g. a
                // workbook writer's Start() fails), so this is the only place left to close it — an
                // un-disposed FileStream here leaks a locked file handle for the process lifetime.
                stream.Dispose();
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        /// <summary>
        /// Same as <see cref="OpenWriteHandle"/>, except the handle is backed by an in-memory buffer
        /// rather than a file — see <see cref="GetWriteHandleBytes"/> to read it out.
        /// </summary>
        internal static int OpenWriteHandleToMemory(int format, NativeWriteOptions options, out NativeWriterHandle? handle)
        {
            handle = null;
            if (!IsWritableFormat(format))
            {
                SetLastError($"xl_open_write_handle_to_memory needs an explicit format (XLS/XLSX/XLSB/CSV); got format {format}.");
                return NativeStatus.InvalidArgument;
            }
            ClearLastError();
            MemoryStream stream = new();
            try
            {
                handle = NativeWriterHandle.Create(stream, format, options);
                handle.MemoryBuffer = stream;
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                stream.Dispose();
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        /// <summary>
        /// Finishes and releases a writer handle opened by <see cref="OpenWriteHandle"/>. Always
        /// disposes <paramref name="handle"/>, even when <see cref="NativeWriterHandle.Close"/> throws
        /// — an unregistered-but-undisposed handle would otherwise leak its open <see cref="FileStream"/>.
        /// </summary>
        internal static int CloseWriteHandle(NativeWriterHandle? handle)
        {
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }
            try
            {
                handle.Close();
                ClearLastError();
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
            finally
            {
                handle.Dispose();
            }
        }

        /// <summary>
        /// Reads back everything written so far to a handle opened by
        /// <see cref="OpenWriteHandleToMemory"/>, ending the workbook's content first if that has not
        /// already happened (see <see cref="NativeWriterHandle.Close"/> — idempotent, so this is safe
        /// to call whether or not the caller ended every sheet/row itself). Does NOT dispose
        /// <paramref name="handle"/>: the caller must still call <see cref="CloseWriteHandle"/>
        /// afterward to release it, same as a file-backed handle.
        /// </summary>
        internal static int GetWriteHandleBytes(NativeWriterHandle? handle, out byte[]? bytes)
        {
            bytes = null;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }
            if (handle.MemoryBuffer is null)
            {
                SetLastError("xl_write_handle_bytes needs a handle opened by xl_open_write_handle_to_memory.");
                return NativeStatus.InvalidArgument;
            }
            try
            {
                handle.Close();
                bytes = handle.MemoryBuffer.ToArray();
                ClearLastError();
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        // XL_FORMAT_AUTO is deliberately absent: sniffing reads an existing file's signature bytes, and
        // a file being created has none.
        private static bool IsWritableFormat(int format)
        {
            return format is NativeFormat.Xls or NativeFormat.Xlsx or NativeFormat.Xlsb or NativeFormat.Csv;
        }

        // Each branch contributes only its construction line and hands off to the one generic body
        // below, since the four writers share no non-generic base.
        private static void WriteToStream(Stream stream, int format, NativeColumnSpec[] specs, NativeTable table, NativeWriteOptions options, string sheetName, bool hasHeader)
        {
            bool date1904 = options.Date1904 ?? false;
            bool sharedStrings = options.UseSharedStrings ?? false;
            switch (format)
            {
                case NativeFormat.Xlsx:
                    WriteWorkbook<XlsxSheetWriter, XlsxRowWriter>(
                        XlsxWorkbookWriter.Create(stream, useSharedStrings: sharedStrings), specs, table, sheetName, hasHeader);
                    return;
                case NativeFormat.Xlsb:
                    WriteWorkbook<XlsbSheetWriter, XlsbRowWriter>(
                        XlsbWorkbookWriter.Create(stream, date1904: date1904, useSharedStrings: sharedStrings), specs, table, sheetName, hasHeader);
                    return;
                case NativeFormat.Xls:
                    WriteWorkbook<XlsSheetWriter, XlsRowWriter>(
                        XlsWorkbookWriter.Create(stream, date1904: date1904), specs, table, sheetName, hasHeader);
                    return;
                default:
                    WriteWorkbook<CsvSheetWriter, CsvRowWriter>(
                        CsvWorkbookWriter.Create(stream, options: options.ToCsvWriterOptions()), specs, table, sheetName, hasHeader);
                    return;
            }
        }

        private static void WriteWorkbook<TSheet, TRow>(IWorkbookWriter<TSheet> workbook, NativeColumnSpec[] specs, NativeTable table, string sheetName, bool hasHeader)
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            try
            {
                workbook.Start();
                TSheet sheet = workbook.AddSheet(sheetName);
                ApplyTemporalStyles<TSheet, TRow>(workbook, sheet, table);
                sheet.Start();

                if (hasHeader)
                {
                    WriteHeaderRow<TSheet, TRow>(sheet, specs);
                }
                for (long row = 0; row < table.RowCount; row++)
                {
                    WriteDataRow<TSheet, TRow>(sheet, table, row);
                }

                sheet.End();
                sheet.Dispose();
                workbook.End();
            }
            finally
            {
                workbook.Dispose();
            }
        }

        // Must run before sheet.Start (SetColumnStyle throws afterward). Without a number format a
        // temporal cell renders in Excel as its raw serial number.
        private static void ApplyTemporalStyles<TSheet, TRow>(IWorkbookWriter<TSheet> workbook, TSheet sheet, NativeTable table)
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            int timeStyle = -1;
            int timestampStyle = -1;
            for (int index = 0; index < table.ColumnCount; index++)
            {
                switch (ColumnAt(table, index).Type)
                {
                    case NativeColumnType.Date:
                        sheet.SetColumnStyle(index, BuiltinDateStyleId);
                        break;
                    case NativeColumnType.Time:
                        timeStyle = timeStyle < 0 ? workbook.AddStyle(new CellStyle { NumberFormat = "hh:mm:ss" }) : timeStyle;
                        sheet.SetColumnStyle(index, timeStyle);
                        break;
                    case NativeColumnType.Timestamp:
                        timestampStyle = timestampStyle < 0 ? workbook.AddStyle(new CellStyle { NumberFormat = "yyyy-mm-dd hh:mm:ss" }) : timestampStyle;
                        sheet.SetColumnStyle(index, timestampStyle);
                        break;
                    default:
                        break;
                }
            }
        }

        private static void WriteHeaderRow<TSheet, TRow>(TSheet sheet, NativeColumnSpec[] specs)
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            using TRow row = sheet.StartRow();
            foreach (NativeColumnSpec spec in specs)
            {
                row.Write(spec.Names[0]);
            }
        }

        private static void WriteDataRow<TSheet, TRow>(TSheet sheet, NativeTable table, long rowIndex)
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            using TRow row = sheet.StartRow();
            for (int index = 0; index < table.ColumnCount; index++)
            {
                WriteCell(row, ColumnAt(table, index), rowIndex);
            }
        }

        internal static void WriteCell(IRowWriter row, NativeColumn column, long rowIndex)
        {
            if (!IsValidAt(column, rowIndex))
            {
                WriteNullCell(row, column.Type);
                return;
            }
            switch (column.Type)
            {
                case NativeColumnType.String:
                    // ponytail: one managed string per cell — IRowWriter has no UTF-8/span overload.
                    // Upgrade path is a Write(ReadOnlySpan<byte>) overload in ExcelReader.Core, which is
                    // a public-API change and out of this plan's scope.
                    int* offsets = (int*)column.Values;
                    int start = offsets[rowIndex];
                    // Length first: TryValidateStringOffsets permits Data == NULL when data_len is 0 (a
                    // legal all-empty-strings column), and Encoding.UTF8.GetString null-checks its
                    // pointer BEFORE its zero-count fast path, so it would throw on that valid input.
                    int length = offsets[rowIndex + 1] - start;
                    row.Write(length == 0 ? string.Empty : Encoding.UTF8.GetString((byte*)column.Data + start, length));
                    return;
                case NativeColumnType.Int64:
                    row.Write(((long*)column.Values)[rowIndex]);
                    return;
                case NativeColumnType.Float64:
                    row.Write(((double*)column.Values)[rowIndex]);
                    return;
                case NativeColumnType.Bool:
                    row.Write(((byte*)column.Values)[rowIndex] != 0);
                    return;
                case NativeColumnType.Date:
                    row.Write(DateOnly.FromDayNumber(WriteUnixEpochDayNumber + ((int*)column.Values)[rowIndex]));
                    return;
                case NativeColumnType.Time:
                    {
                        // checked: an unchecked overflow here would silently write the wrong time instead of
                        // failing the call. The pointer cast happens outside the checked block: CA2020 treats
                        // a checked native-int-to-pointer conversion as throwing on overflow starting in .NET
                        // 7, which is not what this line means to check (the multiplication is).
                        long* values = (long*)column.Values;
                        row.Write(new TimeOnly(checked(values[rowIndex] * TimeSpan.TicksPerMicrosecond)));
                        return;
                    }
                default:
                    {
                        long* values = (long*)column.Values;
                        row.Write(DateTime.UnixEpoch.AddTicks(checked(values[rowIndex] * TimeSpan.TicksPerMicrosecond)));
                        return;
                    }
            }
        }

        // The nullable overloads make a blank cell; a default value would write a real 0/false/epoch.
        internal static void WriteNullCell(IRowWriter row, int type)
        {
            switch (type)
            {
                case NativeColumnType.String:
                    row.Write((string?)null);
                    return;
                case NativeColumnType.Int64:
                    row.Write((long?)null);
                    return;
                case NativeColumnType.Float64:
                    row.Write((double?)null);
                    return;
                case NativeColumnType.Bool:
                    row.Write((bool?)null);
                    return;
                case NativeColumnType.Date:
                    row.Write((DateOnly?)null);
                    return;
                case NativeColumnType.Time:
                    row.Write((TimeOnly?)null);
                    return;
                default:
                    row.Write((DateTime?)null);
                    return;
            }
        }

        // NULL validity pointer is "no nulls in this column", not an error.
        private static bool IsValidAt(NativeColumn column, long rowIndex)
        {
            if (column.Validity == IntPtr.Zero)
            {
                return true;
            }
            byte* bitmap = (byte*)column.Validity;
            return (bitmap[rowIndex >> 3] & (1 << (int)(rowIndex & 7))) != 0;
        }

        /// <summary>
        /// Validates a caller-supplied write table before a single byte is written.
        /// </summary>
        /// <remarks>
        /// This is the trust boundary. Every pointer reachable from <paramref name="table"/> belongs to
        /// the caller and is about to be dereferenced against lengths the caller also supplied, so a
        /// value that is merely wrong here becomes an out-of-bounds read of the caller's process
        /// memory a few frames later. The string-offset walk below is the reason this runs to
        /// completion up front rather than checking each row as it is written: a partially written
        /// file plus a segfault is strictly worse than a rejected call.
        /// </remarks>
        /// <param name="hasHeader">True when every spec carries a name, so a header row must be written.</param>
        internal static bool TryValidateWriteTable(NativeColumnSpec[] specs, NativeTable table, out bool hasHeader, [NotNullWhen(false)] out string? error)
        {
            hasHeader = false;
            error = null;
            if (specs.Length == 0 || !IsValidSpecCount(table.ColumnCount))
            {
                error = $"xl_write_typed needs 1..{NativeLimits.MaxColumnSpecs} columns; got {table.ColumnCount}.";
                return false;
            }
            if (specs.Length != table.ColumnCount)
            {
                error = $"xl_write_typed got {specs.Length} spec(s) for {table.ColumnCount} column(s).";
                return false;
            }
            if (table.RowCount < 0 || table.Columns == IntPtr.Zero)
            {
                error = $"xl_write_typed needs a non-negative row_count and a non-NULL columns pointer; got {table.RowCount}.";
                return false;
            }

            hasHeader = specs[0].Names.Length > 0;
            for (int index = 0; index < table.ColumnCount; index++)
            {
                int nameCount = specs[index].Names.Length;
                if ((nameCount > 0) != hasHeader)
                {
                    error = "every column spec must have a name, or none may — xl_write_typed cannot write a partial header row.";
                    return false;
                }
                if (nameCount > 1)
                {
                    error = $"column {index} is a write spec and must have exactly one name; got {nameCount}.";
                    return false;
                }
                if (!TryValidateWriteColumn(specs[index], ColumnAt(table, index), index, table.RowCount, out error))
                {
                    return false;
                }
            }
            return true;
        }

        private static bool TryValidateWriteColumn(NativeColumnSpec spec, NativeColumn column, int index, long rowCount, [NotNullWhen(false)] out string? error)
        {
            error = null;
            if (column.Type is < NativeColumnType.String or > NativeColumnType.Timestamp)
            {
                error = $"column {index} has unknown type {column.Type}.";
                return false;
            }
            if (spec.Type != column.Type)
            {
                error = $"column {index} has type {column.Type} but its spec says {spec.Type}.";
                return false;
            }
            if (column.Length != rowCount)
            {
                error = $"column {index} has length {column.Length}, but the table's row_count is {rowCount}.";
                return false;
            }
            if (rowCount > 0 && column.Values == IntPtr.Zero)
            {
                error = $"column {index} has {rowCount} row(s) but a NULL values pointer.";
                return false;
            }
            if (column.Type != NativeColumnType.String)
            {
                return true;
            }
            return TryValidateStringOffsets(column, index, rowCount, out error);
        }

        // The whole offsets array is walked once before any of it is used as a slice bound; checking
        // lazily per row would let a hostile offset reach `data` before validation catches it.
        private static bool TryValidateStringOffsets(NativeColumn column, int index, long rowCount, [NotNullWhen(false)] out string? error)
        {
            error = null;
            if (column.DataLen < 0 || (column.DataLen > 0 && column.Data == IntPtr.Zero))
            {
                error = $"column {index} has data_len {column.DataLen} but a NULL data pointer.";
                return false;
            }
            if (rowCount == 0)
            {
                return true;
            }

            int* offsets = (int*)column.Values;
            if (offsets[0] != 0)
            {
                error = $"column {index}: the first string offset must be 0, got {offsets[0]}.";
                return false;
            }
            for (long row = 0; row < rowCount; row++)
            {
                if (offsets[row + 1] < offsets[row])
                {
                    error = $"column {index}: string offset {row + 1} ({offsets[row + 1]}) is before offset {row} ({offsets[row]}).";
                    return false;
                }
            }
            if (offsets[rowCount] != column.DataLen)
            {
                error = $"column {index}: the last string offset is {offsets[rowCount]}, but data_len is {column.DataLen}.";
                return false;
            }
            return true;
        }
    }
}
