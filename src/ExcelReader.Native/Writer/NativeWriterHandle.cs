using ExcelReader.Core.Writer;

namespace ExcelReader.Native.Writer
{
    /// <summary>
    /// Everything one streaming write session needs on the managed side of the boundary — the
    /// writer-side counterpart to <see cref="NativeHandle"/>. The caller only ever sees an opaque id
    /// into <see cref="NativeHandleTable"/>.
    /// </summary>
    /// <remarks>
    /// One sheet, one row open at a time: <see cref="StartSheet"/> must precede <see cref="StartRow"/>,
    /// which must precede the <c>WriteXxx</c> calls for that row, which must precede
    /// <see cref="EndRow"/>. Calling any of these out of order throws
    /// <see cref="InvalidOperationException"/>, which <see cref="Exports"/> turns into
    /// <see cref="NativeStatus.Error"/> plus a message from <see cref="NativeApi.SetLastError"/> —
    /// never lets it escape across the ABI.
    /// </remarks>
    internal abstract class NativeWriterHandle : IDisposable
    {
        internal abstract void StartSheet(string name);

        internal abstract void StartRow();

        /// <summary>Writes a text cell, or a blank cell if <paramref name="value"/> is <see langword="null"/>.</summary>
        internal abstract void WriteString(string? value);

        /// <summary>Writes an integer cell.</summary>
        internal abstract void WriteInt64(long value);

        /// <summary>Writes a floating-point cell.</summary>
        internal abstract void WriteFloat64(double value);

        /// <summary>Writes a boolean cell.</summary>
        internal abstract void WriteBool(bool value);

        /// <summary>Writes a date-only cell.</summary>
        /// <param name="daysSinceEpoch">Days since 1970-01-01 (mirrors <see cref="NativeColumnType.Date"/>'s wire format).</param>
        internal abstract void WriteDate(int daysSinceEpoch);

        /// <summary>Writes a time-of-day cell.</summary>
        /// <param name="microsecondsSinceMidnight">Mirrors <see cref="NativeColumnType.Time"/>'s wire format.</param>
        internal abstract void WriteTime(long microsecondsSinceMidnight);

        /// <summary>Writes a date/time cell.</summary>
        /// <param name="microsecondsSinceEpoch">Mirrors <see cref="NativeColumnType.Timestamp"/>'s wire format.</param>
        internal abstract void WriteTimestamp(long microsecondsSinceEpoch);

        /// <summary>Writes a blank cell of the given <see cref="NativeColumnType"/>.</summary>
        internal abstract void WriteNull(int type);

        internal abstract void EndRow();

        internal abstract void EndSheet();

        /// <summary>
        /// Finishes the workbook: closes any row/sheet still open, then writes the workbook's trailing
        /// structure (<c>IWorkbookWriter.End</c>) — the zip central directory for XLSX/XLSB, the BIFF
        /// EOF record for XLS. Must run before <see cref="Dispose"/> for the output file to be valid;
        /// <see cref="NativeApi.CloseWriteHandle"/> always calls both, in that order.
        /// </summary>
        internal abstract void Close();

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
