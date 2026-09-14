using ExcelReader.Core.Writer;

namespace ExcelReader.Native.Writer
{
    // Everything one streaming write session needs on the managed side of the boundary — the
    // writer-side counterpart to NativeHandle. The caller only ever sees an opaque id
    // into NativeHandleTable.
    //
    // One sheet, one row open at a time: StartSheet must precede StartRow,
    // which must precede the WriteXxx calls for that row, which must precede
    // EndRow. Calling any of these out of order throws
    // InvalidOperationException, which Exports turns into
    // NativeStatus.Error plus a message from NativeApi.SetLastError —
    // never lets it escape across the ABI.
    internal abstract class NativeWriterHandle : IDisposable
    {
        internal abstract void StartSheet(string name);

        internal abstract void StartRow();

        // Writes a text cell, or a blank cell if value is null.
        internal abstract void WriteString(string? value);

        // Writes an integer cell.
        internal abstract void WriteInt64(long value);

        // Writes a floating-point cell.
        internal abstract void WriteFloat64(double value);

        // Writes a boolean cell.
        internal abstract void WriteBool(bool value);

        // Writes a date-only cell.
        // daysSinceEpoch: Days since 1970-01-01 (mirrors NativeColumnType.Date's wire format).
        internal abstract void WriteDate(int daysSinceEpoch);

        // Writes a time-of-day cell.
        // microsecondsSinceMidnight: Mirrors NativeColumnType.Time's wire format.
        internal abstract void WriteTime(long microsecondsSinceMidnight);

        // Writes a date/time cell.
        // microsecondsSinceEpoch: Mirrors NativeColumnType.Timestamp's wire format.
        internal abstract void WriteTimestamp(long microsecondsSinceEpoch);

        // Writes a blank cell of the given NativeColumnType.
        internal abstract void WriteNull(int type);

        internal abstract void EndRow();

        internal abstract void EndSheet();

        // Finishes the workbook: closes any row/sheet still open, then writes the workbook's trailing
        // structure (IWorkbookWriter.End) — the zip central directory for XLSX/XLSB, the BIFF
        // EOF record for XLS. Must run before Dispose for the output file to be valid;
        // NativeApi.CloseWriteHandle always calls both, in that order. Idempotent: a
        // second call is a no-op, so NativeApi.GetWriteHandleBytes can call this to
        // guarantee a complete result without caring whether the caller already ended the workbook.
        internal abstract void Close();

        // Set by NativeApi.OpenWriteHandleToMemory right after construction, to the
        // exact MemoryStream passed to Create — null for a file-backed
        // handle. NativeApi.GetWriteHandleBytes reads it back out; nothing else here
        // needs to know a handle is memory-backed rather than file-backed.
        internal MemoryStream? MemoryBuffer { get; set; }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected abstract void Dispose(bool disposing);

        // Mirrors the per-format switch in NativeApi.Write.cs's WriteToStream, except the workbook
        // writer this returns stays alive across calls instead of running start-to-end in one method.
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
                    // Unreachable: NativeApi.OpenWriteHandle rejects every format but the four above
                    // via IsWritableFormat before this runs. Kept as a hard failure rather than a
                    // silent fall-through in case that guard is ever loosened without updating this.
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
            // checked: an unchecked overflow here would silently write the wrong time instead of
            // failing the call — same reasoning as NativeApi.WriteCell's Time case.
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
            // Idempotent: GetWriteHandleBytes calls this to guarantee complete output regardless of
            // whether the caller already ended the workbook, and CloseWriteHandle may then call it
            // again on the same handle. A second _workbook.End() is not something every writer is
            // guaranteed to tolerate, so the guard lives here rather than relying on that.
            if (_closed)
            {
                return;
            }
            _closed = true;

            // Unlike EndRow/EndSheet, Close is the forceful "finish whatever is pending" step: a row
            // or sheet the caller forgot to end is closed here rather than rejected, since the whole
            // point of xl_close_write_handle is to always leave a valid file behind.
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
