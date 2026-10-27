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
            Dispose(true);
            StagedFile?.Dispose();
            GC.SuppressFinalize(this);
        }

        protected abstract void Dispose(bool disposing);

        internal abstract void WriteSheet(string name, NativeColumnSpec[] specs, NativeTable table, bool hasHeader, NativeExcelTable? excelTable);

        internal static NativeWriterHandle Create(Stream stream, int format, NativeWriteOptions options)
        {
            bool date1904 = options.Date1904 ?? false;
            bool sharedStrings = options.UseSharedStrings ?? false;
            ParallelWriteOptions? parallel = options.ToParallelWriteOptions(format);
            switch (format)
            {
                case NativeFormat.Xlsx:
                    return new NativeWriterHandle<XlsxSheetWriter, XlsxRowWriter>(
                        XlsxWorkbookWriter.Create(stream, options: new XlsxWriterOptions { Date1904 = date1904, UseSharedStrings = sharedStrings }),
                        parallel, static (sheet, rows, writeRow, opts) => sheet.WriteRecordsParallelAsync(rows, writeRow, opts));
                case NativeFormat.Xlsb:
                    return new NativeWriterHandle<XlsbSheetWriter, XlsbRowWriter>(
                        XlsbWorkbookWriter.Create(stream, options: new XlsbWriterOptions { Date1904 = date1904, UseSharedStrings = sharedStrings }),
                        parallel, static (sheet, rows, writeRow, opts) => sheet.WriteRecordsParallelAsync(rows, writeRow, opts));
                case NativeFormat.Xls:
                    return new NativeWriterHandle<XlsSheetWriter, XlsRowWriter>(
                        XlsWorkbookWriter.Create(stream, options: new XlsWriterOptions { Date1904 = date1904 }), null, null);
                case NativeFormat.Csv:
                    return new NativeWriterHandle<CsvSheetWriter, CsvRowWriter>(
                        CsvWorkbookWriter.Create(stream, options: options.ToCsvWriterOptions()), null, null);
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
        private readonly ParallelWriteOptions? _parallel;
        private readonly Func<TSheet, IEnumerable<long>, Action<TRow, long>, ParallelWriteOptions, ValueTask>? _writeParallel;
        private TSheet? _sheet;
        private TRow? _row;
        private bool _closed;
        private Exception? _closeFailure;

        internal NativeWriterHandle(IWorkbookWriter<TSheet> workbook, ParallelWriteOptions? parallel,
            Func<TSheet, IEnumerable<long>, Action<TRow, long>, ParallelWriteOptions, ValueTask>? writeParallel)
        {
            _workbook = workbook;
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
            if (_sheet is null)
            {
                throw new InvalidOperationException("Cannot start a row before starting a sheet.");
            }
            if (_row is not null)
            {
                throw new InvalidOperationException("A row is already open; call xl_end_row before starting another.");
            }
            _row = _sheet.StartRow();
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
            _row.Dispose();
            _row = null;
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

        internal override void WriteSheet(string name, NativeColumnSpec[] specs, NativeTable table, bool hasHeader, NativeExcelTable? excelTable)
        {
            if (_sheet is not null)
            {
                throw new InvalidOperationException("A sheet is already open; call xl_end_sheet before writing a whole sheet.");
            }
            if (excelTable is not null && !typeof(TSheet).IsAssignableTo(typeof(ITableSheetWriter<TRow>)))
            {
                throw new NotSupportedException("Excel tables can only be written to XLSX or XLSB.");
            }
            _sheet = _workbook.AddSheet(name);
            WriteApi.WriteSheetBody(_workbook, _sheet, specs, table, hasHeader, excelTable, _parallel, _writeParallel);
            EndSheet();
        }

        internal override void Close()
        {
            if (_closeFailure is not null)
            {
                throw new InvalidOperationException($"The workbook failed to finish and has no output: {_closeFailure.Message}", _closeFailure);
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
                _closeFailure = exception;
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
