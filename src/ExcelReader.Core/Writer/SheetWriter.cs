using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
{
    public sealed class SheetWriter : IAsyncDisposable
    {
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "WorkbookWriter is borrowed; its lifetime is managed by the caller.")]
        private readonly WorkbookWriter _owner;
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "ZipArchive is borrowed from WorkbookWriter; its lifetime exceeds this sheet.")]
        private readonly ZipArchive _zip;
        private readonly string _name;
        private readonly int _sheetId;
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "StreamWriter is explicitly disposed in EndAsync via DisposeAsync.")]
        private StreamWriter? _xml;
        private int _rowNumber;
        private WriterState _state = WriterState.Created;
        private bool _rowActive;

        internal SheetWriter(WorkbookWriter owner, ZipArchive zip, string name, int sheetId)
        {
            _owner = owner;
            _zip = zip;
            _name = name;
            _sheetId = sheetId;
        }

        internal string Name => _name;
        internal int SheetId => _sheetId;

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP003:Dispose previous before re-assigning",
            Justification = "_xml is always null when StartAsync is called (state machine guarantees Created state).")]
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
            ZipArchiveEntry entry = _zip.CreateEntry($"xl/worksheets/sheet{_sheetId}.xml", CompressionLevel.Optimal);
            Stream stream = entry.Open();
            _xml = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), leaveOpen: false);
            await _xml.WriteAsync(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                $"<worksheet xmlns=\"{XlsxConstants.MainNs}\"><sheetData>").ConfigureAwait(false);
            _state = WriterState.Started;
            _owner.RegisterSheet(_name, _sheetId);
        }

        public async ValueTask<RowWriter> StartRowAsync(CancellationToken ct = default)
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
            await _xml!.WriteAsync($"<row r=\"{_rowNumber}\">").ConfigureAwait(false);
            return new RowWriter(this, _xml!, _rowNumber);
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
            await _xml!.WriteAsync("</sheetData></worksheet>").ConfigureAwait(false);
            await _xml.FlushAsync(ct).ConfigureAwait(false);
            _xml.Dispose();
            _xml = null;
            _owner.NotifySheetEnded();
        }

        public async ValueTask DisposeAsync()
        {
            if (_state == WriterState.Started)
            {
                await EndAsync().ConfigureAwait(false);
            }
        }
    }
}
