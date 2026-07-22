using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
{
    public sealed class XlsxSheetWriter : ISheetWriter<XlsxRowWriter>
    {
        private readonly XlsxWorkbookWriter _owner;
        private readonly ZipArchive _zip;
        private readonly CompressionLevel _compression;
        private readonly BiffBuffer _rowBuffer = new(512);
        // ponytail: flush to the deflate stream once buffered rows pass 64 KB — bounds memory on huge
        // sheets while turning ~50k tiny per-row Writes into a handful of big ones. Kept at/under the
        // ArrayPool.Shared LOH threshold (a larger request would rent from the LOH and never leave it).
        private const int FlushThreshold = 64 * 1024;
        private XlsxRowWriter? _rowWriter;
        private Stream? _stream;
        private int _rowNumber;
        private WriterState _state = WriterState.Created;
        private bool _rowActive;

        internal XlsxSheetWriter(XlsxWorkbookWriter owner, ZipArchive zip, string name, int sheetId, CompressionLevel compression)
        {
            _owner = owner;
            _zip = zip;
            Name = name;
            SheetId = sheetId;
            _compression = compression;
        }

        internal string Name { get; }
        internal int SheetId { get; }
        internal bool UseSharedStrings => _owner.UseSharedStrings;

        internal int GetSharedStringIndex(string value)
        {
            return _owner.GetSharedStringIndex(value);
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP003:Dispose previous before re-assigning",
            Justification = "_stream is always null when StartAsync is called (state machine guarantees Created state).")]
        [SuppressMessage("AsyncFixer", "AsyncFixer02:Long-running or blocking operation invoked inside an async method",
            Justification = "ZipArchiveEntry.Open() is used intentionally; OpenAsync entry-tracking semantics differ in .NET 10.")]
        [SuppressMessage("Reliability", "CA1849:Call async methods when in an async method",
            Justification = "See AsyncFixer02 justification above.")]
        [SuppressMessage("Sonar", "S6966:Await DisposeAsync instead",
            Justification = "See AsyncFixer02 justification above.")]
        public async ValueTask StartAsync(CancellationToken ct = default)
        {
            ObjectDisposedException.ThrowIf(_state == WriterState.Ended, this);
            if (_state != WriterState.Created)
            {
                throw new InvalidOperationException("XlsxSheetWriter has already been started.");
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

        public ValueTask<XlsxRowWriter> StartRowAsync(CancellationToken ct = default)
        {
            int rowNumber = BeginRow(ct);
            _rowWriter ??= new XlsxRowWriter(this, _rowBuffer);
            _rowWriter.Reset(rowNumber);
            return ValueTask.FromResult(_rowWriter);
        }

        // Sync counterpart to StartRowAsync: row buffering and the (rare) threshold flush are already
        // fully synchronous internally (BeginRow, EndBufferedRow), so a caller that never needs to
        // await mid-row (e.g. SheetWriterExtensions.WriteRecordsAsync's XlsxSheetWriter-specific overload)
        // can skip the per-row ValueTask/async-disposable machinery entirely.
        public XlsxRowWriter StartRow(CancellationToken ct = default)
        {
            int rowNumber = BeginRow(ct);
            _rowWriter ??= new XlsxRowWriter(this, _rowBuffer);
            _rowWriter.Reset(rowNumber);
            return _rowWriter;
        }

        [SuppressMessage("Reliability", "CA1849:Call async methods when in an async method",
            Justification = "Dispose() is called synchronously to ensure ZipArchive entry tracking is updated before this method returns.")]
        [SuppressMessage("Sonar", "S6966:Await DisposeAsync instead",
            Justification = "See CA1849 justification above.")]
        [SuppressMessage("VisualStudio.Threading", "VSTHRD103:Dispose synchronously blocks",
            Justification = "See CA1849 justification above.")]
        public async ValueTask EndAsync(CancellationToken ct = default)
        {
            ObjectDisposedException.ThrowIf(_state == WriterState.Ended, this);
            if (_state != WriterState.Started)
            {
                throw new InvalidOperationException("XlsxSheetWriter must be started before ending.");
            }
            ct.ThrowIfCancellationRequested();
            _state = WriterState.Ended;
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
                throw new InvalidOperationException("XlsxSheetWriter must be started before adding rows.");
            }
            if (_rowActive)
            {
                throw new InvalidOperationException("The previous XlsxRowWriter must be disposed before starting a new row.");
            }
            ct.ThrowIfCancellationRequested();
            if (_rowNumber >= 1_048_576)
            {
                throw new ExcelLimitExceededException("Rows", 1_048_576, _rowNumber + 1L);
            }
            _rowNumber++;
            _rowActive = true;
            // The `r` attribute on <row> is optional per ECMA-376 (rows are positional); omitting it
            // shrinks the XML fed to deflate and skips a Utf8Formatter call on every row.
            _rowBuffer.Write("<row>"u8);
            return _rowNumber;
        }

        internal void EndBufferedRow()
        {
            _rowBuffer.Write("</row>"u8);
            if (_rowBuffer.Length >= FlushThreshold)
            {
                _stream!.Write(_rowBuffer.Span);
                _rowBuffer.Reset();
            }
            _rowActive = false;
        }

        internal ValueTask EndBufferedRowAsync(CancellationToken ct = default)
        {
            _rowBuffer.Write("</row>"u8);
            if (_rowBuffer.Length >= FlushThreshold)
            {
                return FlushRowBufferAsync(ct);
            }
            _rowActive = false;
            return ValueTask.CompletedTask;
        }

        private async ValueTask FlushRowBufferAsync(CancellationToken ct)
        {
            await _stream!.WriteAsync(_rowBuffer.Memory, ct).ConfigureAwait(false);
            _rowBuffer.Reset();
            _rowActive = false;
        }
    }
}
