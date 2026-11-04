using System.Diagnostics.CodeAnalysis;
using System.Text;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Internal;
using ExcelReader.Native.Typed;

namespace ExcelReader.Native.Writer
{
    internal static unsafe partial class WriteApi
    {
        internal static readonly int WriteUnixEpochDayNumber = new DateOnly(1970, 1, 1).DayNumber;

        internal static int WriteTyped(ReadOnlySpan<byte> path, int format, NativeColumnSpec[] specs, NativeTable table, NativeWriteOptions options)
        {
            if (path.IsEmpty)
            {
                NativeApi.SetLastError("xl_write_typed needs a non-empty path.");
                return NativeStatus.InvalidArgument;
            }
            if (!IsWritableFormat(format))
            {
                NativeApi.SetLastError($"xl_write_typed needs an explicit format (XLS/XLSX/XLSB/CSV); got format {format}.");
                return NativeStatus.InvalidArgument;
            }
            if (!TryValidateWriteTable(specs, table, out bool hasHeader, out string? validationError))
            {
                NativeApi.SetLastError(validationError);
                return NativeStatus.InvalidArgument;
            }

            NativeApi.ClearLastError();
            try
            {
                string filePath = Encoding.UTF8.GetString(path);
                string sheetName = options.SheetName ?? "Sheet1";
                using StagedFile staged = new(filePath);
                WriteToStream(staged.Stream, format, specs, table, options, sheetName, hasHeader);
                staged.Commit();
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int WriteTypedToMemory(int format, NativeColumnSpec[] specs, NativeTable table, NativeWriteOptions options, out byte[]? bytes)
        {
            bytes = null;
            if (!IsWritableFormat(format))
            {
                NativeApi.SetLastError($"xl_write_typed_to_memory needs an explicit format (XLS/XLSX/XLSB/CSV); got format {format}.");
                return NativeStatus.InvalidArgument;
            }
            if (!TryValidateWriteTable(specs, table, out bool hasHeader, out string? validationError))
            {
                NativeApi.SetLastError(validationError);
                return NativeStatus.InvalidArgument;
            }

            NativeApi.ClearLastError();
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
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int OpenWriteHandle(ReadOnlySpan<byte> path, int format, NativeWriteOptions options, out NativeWriterHandle? handle)
        {
            handle = null;
            if (path.IsEmpty)
            {
                NativeApi.SetLastError("xl_open_write_handle needs a non-empty path.");
                return NativeStatus.InvalidArgument;
            }
            if (!IsWritableFormat(format))
            {
                NativeApi.SetLastError($"xl_open_write_handle needs an explicit format (XLS/XLSX/XLSB/CSV); got format {format}.");
                return NativeStatus.InvalidArgument;
            }
            NativeApi.ClearLastError();
            string filePath = Encoding.UTF8.GetString(path);
            StagedFile? staged = null;
            try
            {
                staged = new StagedFile(filePath);
                handle = NativeWriterHandle.Create(staged.Stream, format, options);
                handle.StagedFile = staged;
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                staged?.Dispose();
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int OpenWriteHandleToMemory(int format, NativeWriteOptions options, out NativeWriterHandle? handle)
        {
            handle = null;
            if (!IsWritableFormat(format))
            {
                NativeApi.SetLastError($"xl_open_write_handle_to_memory needs an explicit format (XLS/XLSX/XLSB/CSV); got format {format}.");
                return NativeStatus.InvalidArgument;
            }
            NativeApi.ClearLastError();
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
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int CloseWriteHandle(NativeWriterHandle? handle)
        {
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }
            try
            {
                handle.Close();
                handle.StagedFile?.Commit();
                NativeApi.ClearLastError();
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
            finally
            {
                handle.Dispose();
            }
        }

        internal static int AbortWriteHandle(NativeWriterHandle? handle)
        {
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }
            try
            {
                handle.Dispose();
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int GetWriteHandleBytes(NativeWriterHandle? handle, out byte[]? bytes)
        {
            bytes = null;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }
            if (handle.MemoryBuffer is null)
            {
                NativeApi.SetLastError("xl_write_handle_bytes needs a handle opened by xl_open_write_handle_to_memory.");
                return NativeStatus.InvalidArgument;
            }
            try
            {
                handle.Close();
                bytes = handle.MemoryBuffer.ToArray();
                NativeApi.ClearLastError();
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        private static bool IsWritableFormat(int format)
        {
            return format is NativeFormat.Xls or NativeFormat.Xlsx or NativeFormat.Xlsb or NativeFormat.Csv;
        }

        private static void WriteToStream(Stream stream, int format, NativeColumnSpec[] specs, NativeTable table, NativeWriteOptions options, string sheetName, bool hasHeader)
        {
            using NativeWriterHandle handle = NativeWriterHandle.Create(stream, format, options);
            handle.WriteSheet(sheetName, specs, table, hasHeader, excelTable: null);
            handle.Close();
        }

        internal static SheetStyleIds RegisterSheetStyles<TSheet>(IWorkbookWriter<TSheet> workbook, Func<int, int, int> inheritNumberFormat,
            NativeSheetStyles? styles, NativeTable table)
        {
            int[] temporalStyleIds = RegisterTemporalStyles(workbook, table);
            if (styles is null)
            {
                return new SheetStyleIds(temporalStyleIds, null, null, -1);
            }
            int headerStyleId = -1;
            if (styles.HeaderStyle is { } header)
            {
                headerStyleId = workbook.AddStyle(header);
            }
            int[]? columnStyleIds = null;
            if (styles.ColumnStyles is { } columnStyles)
            {
                columnStyleIds = new int[columnStyles.Length];
                for (int index = 0; index < columnStyles.Length; index++)
                {
                    columnStyleIds[index] = -1;
                    if (columnStyles[index] is { } style)
                    {
                        columnStyleIds[index] = OverColumnStyle(workbook.AddStyle(style), temporalStyleIds[index], inheritNumberFormat);
                    }
                }
            }
            return new SheetStyleIds(temporalStyleIds, columnStyleIds, styles.ColumnWidths, headerStyleId);
        }

        private static int OverColumnStyle(int styleId, int columnStyleId, Func<int, int, int> inheritNumberFormat)
        {
            if (columnStyleId < 0)
            {
                return styleId;
            }
            return inheritNumberFormat(styleId, columnStyleId);
        }

        private static int[] RegisterTemporalStyles<TSheet>(IWorkbookWriter<TSheet> workbook, NativeTable table)
        {
            int[] styleIds = new int[table.ColumnCount];
            for (int index = 0; index < styleIds.Length; index++)
            {
                styleIds[index] = -1;
                string? format = TemporalNumberFormat(table.ColumnAt(index).Type);
                if (format is not null)
                {
                    styleIds[index] = workbook.AddStyle(new CellStyle { NumberFormat = format });
                }
            }
            return styleIds;
        }

        private static string? TemporalNumberFormat(int columnType)
        {
            switch (columnType)
            {
                case NativeColumnType.Date:
                    return StyleTable.DateStyle.NumberFormat;
                case NativeColumnType.Time:
                    return "hh:mm:ss";
                case NativeColumnType.Timestamp:
                    return "yyyy-mm-dd hh:mm:ss";
                default:
                    return null;
            }
        }

        internal static void WriteSheetBody<TSheet, TRow>(TSheet sheet, NativeColumnSpec[] specs, NativeTable table, bool hasHeader,
            NativeExcelTable? excelTable, SheetStyleIds styles, ParallelWriteOptions? parallel,
            Func<TSheet, IEnumerable<long>, Action<TRow, long>, ParallelWriteOptions, ValueTask>? writeParallel)
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            ApplyTemporalStyles<TSheet, TRow>(sheet, styles.TemporalStyleIds);
            ApplyColumnWidths<TSheet, TRow>(sheet, styles.ColumnWidths);

            if (excelTable is not null)
            {
                BeginTable<TSheet, TRow>(sheet, specs, excelTable, styles.HeaderStyleId);
                WriteStyledDataRows(sheet, table, styles.ColumnStyleIds, parallel, writeParallel);
                ((ITableSheetWriter<TRow>)sheet).EndTable();
                return;
            }
            if (!hasHeader)
            {
                WriteHeaderlessRows(sheet, table, styles.ColumnStyleIds, parallel, writeParallel);
                return;
            }
            WriteHeaderRow<TSheet, TRow>(sheet, specs, styles.HeaderStyleId);
            WriteStyledDataRows(sheet, table, styles.ColumnStyleIds, parallel, writeParallel);
        }

        private static void BeginTable<TSheet, TRow>(TSheet sheet, NativeColumnSpec[] specs, NativeExcelTable excelTable, int headerStyleId)
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            string[] columns = [.. specs.Select(spec => spec.Names[0])];
            if (headerStyleId < 0)
            {
                ((ITableSheetWriter<TRow>)sheet).BeginTable(excelTable.Name, columns, excelTable.Options);
                return;
            }
            ((IStyleBinder)sheet).BeginStyledTable(excelTable.Name, columns, excelTable.Options, headerStyleId);
        }

        // Push only once the sheet's first row has started: starting it rebuilds the lookup the overlay sits on.
        private static void WriteStyledDataRows<TSheet, TRow>(TSheet sheet, NativeTable table, int[]? columnStyleIds,
            ParallelWriteOptions? parallel, Func<TSheet, IEnumerable<long>, Action<TRow, long>, ParallelWriteOptions, ValueTask>? writeParallel)
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            if (columnStyleIds is null || sheet is not IStyleBinder binder)
            {
                WriteDataRows(sheet, table, 0, parallel, writeParallel);
                return;
            }
            binder.PushColumnStyleOverlay(columnStyleIds, firstColumn: 0);
            try
            {
                WriteDataRows(sheet, table, 0, parallel, writeParallel);
            }
            finally
            {
                binder.PopColumnStyleOverlay();
            }
        }

        private static void WriteHeaderlessRows<TSheet, TRow>(TSheet sheet, NativeTable table, int[]? columnStyleIds,
            ParallelWriteOptions? parallel, Func<TSheet, IEnumerable<long>, Action<TRow, long>, ParallelWriteOptions, ValueTask>? writeParallel)
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            if (columnStyleIds is null || sheet is not IStyleBinder binder || table.RowCount == 0)
            {
                WriteDataRows(sheet, table, 0, parallel, writeParallel);
                return;
            }
            bool pushed = false;
            try
            {
                using (TRow row = sheet.StartRow())
                {
                    binder.PushColumnStyleOverlay(columnStyleIds, firstColumn: 0);
                    pushed = true;
                    WriteDataCells(row, table, 0);
                }
                WriteDataRows(sheet, table, 1, parallel, writeParallel);
            }
            finally
            {
                if (pushed)
                {
                    binder.PopColumnStyleOverlay();
                }
            }
        }

        private static void ApplyColumnWidths<TSheet, TRow>(TSheet sheet, double[]? widths)
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            if (widths is null)
            {
                return;
            }
            for (int index = 0; index < widths.Length; index++)
            {
                if (widths[index] > 0)
                {
                    sheet.SetColumnWidth(index, widths[index]);
                }
            }
        }

        private static void WriteDataRows<TSheet, TRow>(TSheet sheet, NativeTable table, long firstRow,
            ParallelWriteOptions? parallel, Func<TSheet, IEnumerable<long>, Action<TRow, long>, ParallelWriteOptions, ValueTask>? writeParallel)
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            if (parallel is not null && writeParallel is not null)
            {
                writeParallel(sheet, RowIndices(firstRow, table.RowCount), (row, index) => WriteDataCells(row, table, index), parallel)
                    .AsTask().GetAwaiter().GetResult();
                return;
            }
            for (long row = firstRow; row < table.RowCount; row++)
            {
                WriteDataRow<TSheet, TRow>(sheet, table, row);
            }
        }

        private static void ApplyTemporalStyles<TSheet, TRow>(TSheet sheet, int[] temporalStyleIds)
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            for (int index = 0; index < temporalStyleIds.Length; index++)
            {
                if (temporalStyleIds[index] >= 0)
                {
                    sheet.SetColumnStyle(index, temporalStyleIds[index]);
                }
            }
        }

        private static void WriteHeaderRow<TSheet, TRow>(TSheet sheet, NativeColumnSpec[] specs, int headerStyleId)
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            using TRow row = sheet.StartRow();
            foreach (NativeColumnSpec spec in specs)
            {
                if (headerStyleId >= 0)
                {
                    row.Style(headerStyleId);
                }
                row.Write(spec.Names[0]);
            }
        }

        private static void WriteDataRow<TSheet, TRow>(TSheet sheet, NativeTable table, long rowIndex)
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            using TRow row = sheet.StartRow();
            WriteDataCells(row, table, rowIndex);
        }

        private static void WriteDataCells(IRowWriter row, NativeTable table, long rowIndex)
        {
            for (int index = 0; index < table.ColumnCount; index++)
            {
                WriteCell(row, table.ColumnAt(index), rowIndex);
            }
        }

        private static IEnumerable<long> RowIndices(long first, long end)
        {
            for (long index = first; index < end; index++)
            {
                yield return index;
            }
        }

        internal static void WriteCell(IRowWriter row, in NativeColumn column, long rowIndex)
        {
            if (!IsValidAt(column, rowIndex))
            {
                WriteNullCell(row, column.Type);
                return;
            }
            switch (column.Type)
            {
                case NativeColumnType.String:
                    int* offsets = (int*)column.Values;
                    int start = offsets[rowIndex];
                    int length = offsets[rowIndex + 1] - start;
                    row.WriteUtf8(new ReadOnlySpan<byte>((byte*)column.Data + start, length));
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

        private static bool IsValidAt(in NativeColumn column, long rowIndex)
        {
            if (column.Validity == IntPtr.Zero)
            {
                return true;
            }
            byte* bitmap = (byte*)column.Validity;
            return (bitmap[rowIndex >> 3] & (1 << (int)(rowIndex & 7))) != 0;
        }

        internal static bool TryValidateWriteTable(NativeColumnSpec[] specs, NativeTable table, out bool hasHeader, [NotNullWhen(false)] out string? error)
        {
            hasHeader = false;
            error = null;
            if (specs.Length == 0 || !TypedApi.IsValidSpecCount(table.ColumnCount))
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
                if (!TryValidateWriteColumn(specs[index], table.ColumnAt(index), index, table.RowCount, out error))
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
