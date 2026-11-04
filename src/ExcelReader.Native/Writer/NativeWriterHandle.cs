using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Csv;
using ExcelReader.Core.Writer.Internal;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Native.Writer
{
    internal abstract class NativeWriterHandle : IDisposable
    {
        internal abstract void StartSheet(string name);

        internal abstract void StartRow();

        internal abstract void StartRow(int styleId, double height);

        internal abstract int AddStyle(CellStyle style);

        internal abstract void SetColumnStyle(int column, int styleId);

        internal abstract void SetColumnWidth(int column, double width);

        internal abstract void StyleNextCell(int styleId);

        internal abstract void WriteString(string? value);

        internal abstract void WriteInt64(long value);

        internal abstract void WriteFloat64(double value);

        internal abstract void WriteBool(bool value);

        internal abstract void WriteDate(int daysSinceEpoch);

        internal abstract void WriteTime(long microsecondsSinceMidnight);

        internal abstract void WriteTimestamp(long microsecondsSinceEpoch);

        internal abstract void WriteNull(int type);

        internal abstract void EndRow();

        internal abstract void EndSheet();

        internal abstract void Close();

        internal MemoryStream? MemoryBuffer { get; set; }

        internal StagedFile? StagedFile { get; set; }

        public void Dispose()
        {
            try
            {
                Dispose(true);
            }
            finally
            {
                StagedFile?.Dispose();
            }
            GC.SuppressFinalize(this);
        }

        protected abstract void Dispose(bool disposing);

        internal abstract void WriteSheet(string name, NativeColumnSpec[] specs, NativeTable table, bool hasHeader, NativeExcelTable? excelTable, NativeSheetStyles? styles = null);

        internal static NativeWriterHandle Create(Stream stream, int format, NativeWriteOptions options)
        {
            bool date1904 = options.Date1904 ?? false;
            bool sharedStrings = options.UseSharedStrings ?? false;
            ParallelWriteOptions? parallel = options.ToParallelWriteOptions(format);
            switch (format)
            {
                case NativeFormat.Xlsx:
                    {
                        XlsxWorkbookWriter xlsx = XlsxWorkbookWriter.Create(stream, options: new XlsxWriterOptions { Date1904 = date1904, UseSharedStrings = sharedStrings });
                        return new NativeWriterHandle<XlsxSheetWriter, XlsxRowWriter>(xlsx, xlsx.InheritNumberFormat,
                            parallel, static (sheet, rows, writeRow, opts) => sheet.WriteRecordsParallelAsync(rows, writeRow, opts));
                    }
                case NativeFormat.Xlsb:
                    {
                        XlsbWorkbookWriter xlsb = XlsbWorkbookWriter.Create(stream, options: new XlsbWriterOptions { Date1904 = date1904, UseSharedStrings = sharedStrings });
                        return new NativeWriterHandle<XlsbSheetWriter, XlsbRowWriter>(xlsb, xlsb.InheritNumberFormat,
                            parallel, static (sheet, rows, writeRow, opts) => sheet.WriteRecordsParallelAsync(rows, writeRow, opts));
                    }
                case NativeFormat.Xls:
                    {
                        XlsWorkbookWriter xls = XlsWorkbookWriter.Create(stream, options: new XlsWriterOptions { Date1904 = date1904 });
                        return new NativeWriterHandle<XlsSheetWriter, XlsRowWriter>(xls, xls.InheritNumberFormat, null, null);
                    }
                case NativeFormat.Csv:
                    return new NativeWriterHandle<CsvSheetWriter, CsvRowWriter>(
                        CsvWorkbookWriter.Create(stream, options: options.ToCsvWriterOptions()), static (overlayStyleId, _) => overlayStyleId, null, null);
                default:
                    throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported write format.");
            }
        }
    }

    internal sealed class NativeWriterHandle<TSheet, TRow> : NativeWriterHandle
        where TSheet : class, ISheetWriter<TRow>
        where TRow : class, IRowWriter
    {
        private readonly IWorkbookWriter<TSheet> _workbook;
        private readonly Func<int, int, int> _inheritNumberFormat;
        private readonly ParallelWriteOptions? _parallel;
        private readonly Func<TSheet, IEnumerable<long>, Action<TRow, long>, ParallelWriteOptions, ValueTask>? _writeParallel;
        private TSheet? _sheet;
        private TRow? _row;
        private bool _closed;
        private Exception? _failure;

        internal NativeWriterHandle(IWorkbookWriter<TSheet> workbook, Func<int, int, int> inheritNumberFormat, ParallelWriteOptions? parallel,
            Func<TSheet, IEnumerable<long>, Action<TRow, long>, ParallelWriteOptions, ValueTask>? writeParallel)
        {
            _workbook = workbook;
            _inheritNumberFormat = inheritNumberFormat;
            _parallel = parallel;
            _writeParallel = writeParallel;
        }

        internal override void StartSheet(string name)
        {
            if (_sheet is not null)
            {
                throw new InvalidOperationException("A sheet is already open; call xl_end_sheet before starting another.");
            }
            _sheet = _workbook.AddSheet(name);
        }

        internal override void StartRow()
        {
            _row = SheetForNextRow().StartRow();
        }

        internal override void StartRow(int styleId, double height)
        {
            TSheet sheet = SheetForNextRow();
            if (height == 0)
            {
                _row = sheet.StartRow(styleId);
                return;
            }
            _row = sheet.StartRow(styleId, height);
        }

        private TSheet SheetForNextRow()
        {
            if (_sheet is null)
            {
                throw new InvalidOperationException("Cannot start a row before starting a sheet.");
            }
            if (_row is not null)
            {
                throw new InvalidOperationException("A row is already open; call xl_end_row before starting another.");
            }
            return _sheet;
        }

        internal override int AddStyle(CellStyle style)
        {
            return _workbook.AddStyle(style);
        }

        internal override void SetColumnStyle(int column, int styleId)
        {
            OpenSheet().SetColumnStyle(column, styleId);
        }

        internal override void SetColumnWidth(int column, double width)
        {
            OpenSheet().SetColumnWidth(column, width);
        }

        private TSheet OpenSheet()
        {
            return _sheet ?? throw new InvalidOperationException("Cannot set a column style or width before starting a sheet.");
        }

        internal override void StyleNextCell(int styleId)
        {
            Row().Style(styleId);
        }

        internal override void WriteString(string? value)
        {
            Row().Write(value);
        }

        internal override void WriteInt64(long value)
        {
            Row().Write(value);
        }

        internal override void WriteFloat64(double value)
        {
            Row().Write(value);
        }

        internal override void WriteBool(bool value)
        {
            Row().Write(value);
        }

        internal override void WriteDate(int daysSinceEpoch)
        {
            Row().Write(DateOnly.FromDayNumber(WriteApi.WriteUnixEpochDayNumber + daysSinceEpoch));
        }

        internal override void WriteTime(long microsecondsSinceMidnight)
        {
            Row().Write(new TimeOnly(checked(microsecondsSinceMidnight * TimeSpan.TicksPerMicrosecond)));
        }

        internal override void WriteTimestamp(long microsecondsSinceEpoch)
        {
            Row().Write(DateTime.UnixEpoch.AddTicks(checked(microsecondsSinceEpoch * TimeSpan.TicksPerMicrosecond)));
        }

        internal override void WriteNull(int type)
        {
            WriteApi.WriteNullCell(Row(), type);
        }

        private TRow Row()
        {
            return _row ?? throw new InvalidOperationException("Cannot write a cell before starting a row.");
        }

        internal override void EndRow()
        {
            if (_row is null)
            {
                throw new InvalidOperationException("Cannot end a row before starting one.");
            }
            // A row that throws for a pending style has still ended.
            try
            {
                _row.Dispose();
            }
            finally
            {
                _row = null;
            }
        }

        internal override void EndSheet()
        {
            if (_sheet is null)
            {
                throw new InvalidOperationException("Cannot end a sheet before starting one.");
            }
            if (_row is not null)
            {
                throw new InvalidOperationException("Cannot end a sheet with an open row; call xl_end_row first.");
            }
            _sheet.End();
            _sheet.Dispose();
            _sheet = null;
        }

        internal override void WriteSheet(string name, NativeColumnSpec[] specs, NativeTable table, bool hasHeader, NativeExcelTable? excelTable, NativeSheetStyles? styles = null)
        {
            if (_sheet is not null)
            {
                throw new InvalidOperationException("A sheet is already open; call xl_end_sheet before writing a whole sheet.");
            }
            if (excelTable is not null && !typeof(TSheet).IsAssignableTo(typeof(ITableSheetWriter<TRow>)))
            {
                throw new NotSupportedException("Excel tables can only be written to XLSX or XLSB.");
            }
            // Registered before the sheet exists, so a style limit fails without poisoning the handle.
            SheetStyleIds styleIds = WriteApi.RegisterSheetStyles(_workbook, _inheritNumberFormat, styles, table);
            _sheet = _workbook.AddSheet(name);
            try
            {
                WriteApi.WriteSheetBody<TSheet, TRow>(_sheet, specs, table, hasHeader, excelTable, styleIds, _parallel, _writeParallel);
                EndSheet();
            }
            catch (Exception exception)
            {
                _failure = exception;
                throw;
            }
        }

        internal override void Close()
        {
            if (_failure is not null)
            {
                throw new InvalidOperationException($"The workbook has no output because an earlier call failed: {_failure.Message}", _failure);
            }
            if (_closed)
            {
                return;
            }
            try
            {
                _row?.Dispose();
                _row = null;
                if (_sheet is not null)
                {
                    _sheet.End();
                    _sheet.Dispose();
                    _sheet = null;
                }
                _workbook.End();
                _closed = true;
            }
            catch (Exception exception)
            {
                _failure = exception;
                throw;
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (!disposing)
            {
                return;
            }
            // A row or sheet still open here belongs to output that is being discarded; ending it must
            // not replace the failure that left it open.
            FailureCleanup.Dispose(_row);
            FailureCleanup.Dispose(_sheet);
            _workbook.Dispose();
        }
    }
}
