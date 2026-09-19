using ExcelReader.Core.Writer;

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

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected abstract void Dispose(bool disposing);

        internal static NativeWriterHandle Create(Stream stream, int format, NativeWriteOptions options)
        {
            bool date1904 = options.Date1904 ?? false;
            bool sharedStrings = options.UseSharedStrings ?? false;
            switch (format)
            {
                case NativeFormat.Xlsx:
                    return new NativeWriterHandle<XlsxSheetWriter, XlsxRowWriter>(
                        XlsxWorkbookWriter.Create(stream, useSharedStrings: sharedStrings));
                case NativeFormat.Xlsb:
                    return new NativeWriterHandle<XlsbSheetWriter, XlsbRowWriter>(
                        XlsbWorkbookWriter.Create(stream, date1904: date1904, useSharedStrings: sharedStrings));
                case NativeFormat.Xls:
                    return new NativeWriterHandle<XlsSheetWriter, XlsRowWriter>(
                        XlsWorkbookWriter.Create(stream, date1904: date1904));
                case NativeFormat.Csv:
                    return new NativeWriterHandle<CsvSheetWriter, CsvRowWriter>(
                        CsvWorkbookWriter.Create(stream, options: options.ToCsvWriterOptions()));
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
        private TSheet? _sheet;
        private TRow? _row;
        private bool _closed;

        internal NativeWriterHandle(IWorkbookWriter<TSheet> workbook)
        {
            _workbook = workbook;
            _workbook.Start();
        }

        internal override void StartSheet(string name)
        {
            if (_sheet is not null)
            {
                throw new InvalidOperationException("A sheet is already open; call xl_end_sheet before starting another.");
            }
            _sheet = _workbook.AddSheet(name);
            _sheet.Start();
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
            Row().Write(DateOnly.FromDayNumber(NativeApi.WriteUnixEpochDayNumber + daysSinceEpoch));
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
            NativeApi.WriteNullCell(Row(), type);
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

        internal override void Close()
        {
            if (_closed)
            {
                return;
            }
            _closed = true;

            _row?.Dispose();
            _row = null;
            if (_sheet is not null)
            {
                _sheet.End();
                _sheet.Dispose();
                _sheet = null;
            }
            _workbook.End();
        }

        protected override void Dispose(bool disposing)
        {
            if (!disposing)
            {
                return;
            }
            _row?.Dispose();
            _sheet?.Dispose();
            _workbook.Dispose();
        }
    }
}
