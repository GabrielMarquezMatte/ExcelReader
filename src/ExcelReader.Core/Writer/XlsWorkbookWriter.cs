using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
{
    // Writes a BIFF8 (.xls) workbook in the same shape as WorkbookWriter (XLSX). Unlike the ZIP
    // writer, records are buffered in memory and the OLE container is assembled in EndAsync: the
    // BoundSheet offsets in the globals and the OLE FAT both need stream sizes known only at the
    // end. Safe because BIFF8 is capped at 65,536 rows x 256 columns.
    public sealed class XlsWorkbookWriter : IAsyncDisposable
    {
        private const int MaxSheetNameLength = 31;

        private readonly Stream _stream;
        private readonly bool _leaveOpen;
        private readonly bool _date1904;
        private readonly List<XlsSheetWriter> _sheets = [];
        private WriterState _state = WriterState.Created;
        private XlsSheetWriter? _activeSheet;
        private bool _disposed;

        private XlsWorkbookWriter(Stream stream, bool leaveOpen, bool date1904)
        {
            _stream = stream;
            _leaveOpen = leaveOpen;
            _date1904 = date1904;
        }

        public static ValueTask<XlsWorkbookWriter> CreateAsync(
            Stream stream, bool leaveOpen = false, bool date1904 = false, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ct.ThrowIfCancellationRequested();
            return ValueTask.FromResult(new XlsWorkbookWriter(stream, leaveOpen, date1904));
        }

        public ValueTask StartAsync(CancellationToken ct = default)
        {
            ObjectDisposedException.ThrowIf(_state == WriterState.Ended, this);
            if (_state != WriterState.Created)
            {
                throw new InvalidOperationException("XlsWorkbookWriter has already been started.");
            }
            ct.ThrowIfCancellationRequested();
            _state = WriterState.Started;
            return ValueTask.CompletedTask;
        }

        public XlsSheetWriter AddSheet(string name)
        {
            ArgumentNullException.ThrowIfNull(name);
            ObjectDisposedException.ThrowIf(_state == WriterState.Ended, this);
            if (_state != WriterState.Started)
            {
                throw new InvalidOperationException("XlsWorkbookWriter must be started before adding sheets.");
            }
            if (name.Length is 0 or > MaxSheetNameLength)
            {
                throw new ArgumentException($"Sheet names must be 1 to {MaxSheetNameLength} characters.", nameof(name));
            }
            if (_activeSheet is not null)
            {
                throw new InvalidOperationException("The previous XlsSheetWriter must be ended before adding a new sheet.");
            }
            _activeSheet = new XlsSheetWriter(this, name, _date1904);
            return _activeSheet;
        }

        internal void RegisterSheet(XlsSheetWriter sheet)
        {
            _sheets.Add(sheet);
        }

        internal void NotifySheetEnded()
        {
            _activeSheet = null;
        }

        public async ValueTask EndAsync(CancellationToken ct = default)
        {
            ObjectDisposedException.ThrowIf(_state == WriterState.Ended, this);
            if (_state != WriterState.Started)
            {
                throw new InvalidOperationException("XlsWorkbookWriter must be started before ending.");
            }
            ct.ThrowIfCancellationRequested();
            _state = WriterState.Ended;
            if (_activeSheet is not null)
            {
                await _activeSheet.DisposeAsync().ConfigureAwait(false);
            }
            if (_sheets.Count == 0)
            {
                throw new InvalidOperationException("A workbook must contain at least one sheet.");
            }

            BiffBuffer workbook = new(8192);
            try
            {
                string[] names = new string[_sheets.Count];
                for (int i = 0; i < _sheets.Count; i++)
                {
                    names[i] = _sheets[i].Name;
                }

                int[] offsetPositions = XlsGlobals.Write(workbook, names, _date1904);
                for (int i = 0; i < _sheets.Count; i++)
                {
                    // The sheet substream begins at the current end of the buffer.
                    workbook.PatchI32(offsetPositions[i], workbook.Length);
                    _sheets[i].BuildSubstream(workbook);
                }

                await OleCompoundWriter.WriteAsync(_stream, workbook.Memory, ct).ConfigureAwait(false);
            }
            finally
            {
                workbook.Dispose();
                foreach (XlsSheetWriter sheet in _sheets)
                {
                    sheet.ReleaseBuffer();
                }
            }
        }

        public ValueTask FlushAsync(CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            return new ValueTask(_stream.FlushAsync(ct));
        }

        public async ValueTask DisposeAsync()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            if (_state == WriterState.Started)
            {
                // Auto-finalize only a usable workbook; an unused/sheet-less one is abandoned
                // quietly so disposal never throws over cleanup.
                if (_sheets.Count > 0)
                {
                    await EndAsync().ConfigureAwait(false);
                }
                else
                {
                    _state = WriterState.Ended;
                }
            }
            if (!_leaveOpen)
            {
                await _stream.DisposeAsync().ConfigureAwait(false);
            }
        }
    }
}
