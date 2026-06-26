using System.Buffers.Text;
using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
{
    public sealed class SheetWriter : ISheetWriter<RowWriter>
    {
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "WorkbookWriter is borrowed; its lifetime is managed by the caller.")]
        private readonly WorkbookWriter _owner;
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "ZipArchive is borrowed from WorkbookWriter; its lifetime exceeds this sheet.")]
        private readonly ZipArchive _zip;
        private readonly CompressionLevel _compression;
        private readonly BiffBuffer _rowBuffer = new(512);
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "Stream is explicitly disposed in EndAsync via DisposeAsync.")]
        private Stream? _stream;
        private int _rowNumber;
        private WriterState _state = WriterState.Created;
        private bool _rowActive;

        internal SheetWriter(WorkbookWriter owner, ZipArchive zip, string name, int sheetId, CompressionLevel compression)
        {
            _owner = owner;
            _zip = zip;
            Name = name;
            SheetId = sheetId;
            _compression = compression;
        }

        internal string Name { get; }
        internal int SheetId { get; }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP003:Dispose previous before re-assigning",
            Justification = "_stream is always null when StartAsync is called (state machine guarantees Created state).")]
        [SuppressMessage("AsyncFixer", "AsyncFixer02:Long-running or blocking operation invoked inside an async method",
            Justification = "ZipArchiveEntry.Open() is used intentionally; OpenAsync entry-tracking semantics differ in .NET 10.")]
        [SuppressMessage("Reliability", "CA1849:Call async methods when in an async method",
            Justification = "See AsyncFixer02 justification above.")]
        [SuppressMessage("SharpSource", "SS033:Async overload available",
            Justification = "See AsyncFixer02 justification above.")]
        [SuppressMessage("Sonar", "S6966:Await DisposeAsync instead",
            Justification = "See AsyncFixer02 justification above.")]
        public async ValueTask StartAsync(CancellationToken ct = default)
        {
            ObjectDisposedException.ThrowIf(_state == WriterState.Ended, this);
            if (_state != WriterState.Created)
            {
                throw new InvalidOperationException("SheetWriter has already been started.");
            }
            ct.ThrowIfCancellationRequested();
            ZipArchiveEntry entry = _zip.CreateEntry($"xl/worksheets/sheet{SheetId}.xml", _compression);
#if NET10_0_OR_GREATER
            Stream stream = await entry.OpenAsync(ct).ConfigureAwait(false);
#else
            ct.ThrowIfCancellationRequested();
            Stream stream = entry.Open();
#endif
            _stream = stream;
            _rowBuffer.Reset();
            _rowBuffer.WriteUtf8(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                $"<worksheet xmlns=\"{XlsxConstants.MainNs}\"><sheetData>");
            await _stream.WriteAsync(_rowBuffer.Memory, ct).ConfigureAwait(false);
            _rowBuffer.Reset();
            _state = WriterState.Started;
            _owner.RegisterSheet(Name, SheetId);
        }

        public ValueTask<RowWriter> StartRowAsync(CancellationToken ct = default)
        {
            BeginRow(ct);
            return ValueTask.FromResult(new RowWriter(this, _rowBuffer, _rowNumber));
        }

        public ValueTask WriteRow<T>(ReadOnlySpan<T> values, CancellationToken ct = default)
            where T : ISpanFormattable
        {
            int rowNumber = BeginRow(ct);
            for (int i = 0; i < values.Length; i++)
            {
                CellFormatter.WriteNumber(_rowBuffer, values[i], i, rowNumber, includeReference: false);
            }
            EndBufferedRow();
            return ValueTask.CompletedTask;
        }

        internal void NotifyRowEnded()
        {
            _rowActive = false;
        }

        [SuppressMessage("Reliability", "CA1849:Call async methods when in an async method",
            Justification = "Dispose() is called synchronously to ensure ZipArchive entry tracking is updated before this method returns.")]
        [SuppressMessage("Sonar", "S6966:Await DisposeAsync instead",
            Justification = "See CA1849 justification above.")]
        [SuppressMessage("SharpSource", "SS033:Async overload available",
            Justification = "See CA1849 justification above.")]
        [SuppressMessage("SharpSource", "SS059:Async disposable should be disposed asynchronously",
            Justification = "See CA1849 justification above.")]
        [SuppressMessage("VisualStudio.Threading", "VSTHRD103:Dispose synchronously blocks",
            Justification = "See CA1849 justification above.")]
        public async ValueTask EndAsync(CancellationToken ct = default)
        {
            ObjectDisposedException.ThrowIf(_state == WriterState.Ended, this);
            if (_state != WriterState.Started)
            {
                throw new InvalidOperationException("SheetWriter must be started before ending.");
            }
            ct.ThrowIfCancellationRequested();
            _state = WriterState.Ended;
            _rowBuffer.Reset();
            _rowBuffer.Write("</sheetData></worksheet>"u8);
            await _stream!.WriteAsync(_rowBuffer.Memory, ct).ConfigureAwait(false);
            await _stream.FlushAsync(ct).ConfigureAwait(false);
            await _stream.DisposeAsync().ConfigureAwait(false);
            _stream = null;
            _rowBuffer.Dispose();
            _owner.NotifySheetEnded();
        }

        public async ValueTask DisposeAsync()
        {
            if (_state == WriterState.Started)
            {
                await EndAsync().ConfigureAwait(false);
            }
            else if (_state == WriterState.Created)
            {
                _state = WriterState.Ended;
                _rowBuffer.Dispose();
            }
        }

        private int BeginRow(CancellationToken ct)
        {
            ObjectDisposedException.ThrowIf(_state == WriterState.Ended, this);
            if (_state != WriterState.Started)
            {
                throw new InvalidOperationException("SheetWriter must be started before adding rows.");
            }
            if (_rowActive)
            {
                throw new InvalidOperationException("The previous RowWriter must be disposed before starting a new row.");
            }
            ct.ThrowIfCancellationRequested();
            _rowNumber++;
            _rowActive = true;
            _rowBuffer.Reset();
            _rowBuffer.Write("<row r=\""u8);
            Span<byte> rowDigits = stackalloc byte[8];
            Utf8Formatter.TryFormat(_rowNumber, rowDigits, out int written);
            _rowBuffer.Write(rowDigits[..written]);
            _rowBuffer.Write("\">"u8);
            return _rowNumber;
        }

        internal void EndBufferedRow()
        {
            _rowBuffer.Write("</row>"u8);
            _stream!.Write(_rowBuffer.Span);
            _rowActive = false;
        }
    }
}
