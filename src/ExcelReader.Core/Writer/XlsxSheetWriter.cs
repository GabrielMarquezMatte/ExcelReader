using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
{
    /// <summary>
    /// Writes a single worksheet's XML into an XLSX ZIP archive, buffering rows and flushing them to the entry stream once they cross a size threshold.
    /// </summary>
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
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP002:Dispose member",
            Justification = "Reused per row; the caller disposes it via using after each row, and EndAsync's _rowActive guard rejects ending the sheet with it still open.")]
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP006:Implement IDisposable",
            Justification = "Reused per row; the caller disposes it via using after each row, and EndAsync's _rowActive guard rejects ending the sheet with it still open.")]
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

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The sheet has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The sheet has already been started.</exception>
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP003:Dispose previous before re-assigning",
            Justification = "_stream is always null when StartAsync is called (state machine guarantees Created state).")]
        public async ValueTask StartAsync(CancellationToken ct = default)
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireCreated(_state, nameof(XlsxSheetWriter));
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

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The sheet has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The sheet has not been started, or the previous <see cref="XlsxRowWriter"/> has not been disposed.</exception>
        /// <exception cref="ExcelLimitExceededException">The worksheet's 1,048,576-row limit has been reached.</exception>
        public ValueTask<XlsxRowWriter> StartRowAsync(CancellationToken ct = default)
        {
            int rowNumber = BeginRow(ct);
            _rowWriter ??= new XlsxRowWriter(this, _rowBuffer);
            _rowWriter.Reset(rowNumber);
            return ValueTask.FromResult(_rowWriter);
        }

        /// <summary>
        /// Starts a new row without the async/await machinery of <see cref="StartRowAsync(CancellationToken)"/>, for callers that never need to await mid-row.
        /// </summary>
        /// <remarks>
        /// Safe because row buffering and the (rare) threshold flush (<c>BeginRow</c>, <c>EndBufferedRow</c>)
        /// are already fully synchronous internally, so a caller that never needs to await mid-row (e.g.
        /// <see cref="SheetWriterExtensions"/>'s <see cref="XlsxSheetWriter"/>-specific <c>WriteRecordsAsync</c>
        /// overload) can skip the per-row <see cref="ValueTask"/>/async-disposable machinery entirely.
        /// </remarks>
        /// <param name="ct">A token checked before the row starts; the row buffering itself is fully synchronous.</param>
        /// <returns>The reusable <see cref="XlsxRowWriter"/> for the new row.</returns>
        /// <exception cref="ObjectDisposedException">The sheet has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The sheet has not been started, or the previous <see cref="XlsxRowWriter"/> has not been disposed.</exception>
        /// <exception cref="ExcelLimitExceededException">The worksheet's 1,048,576-row limit has been reached.</exception>
        public XlsxRowWriter StartRow(CancellationToken ct = default)
        {
            int rowNumber = BeginRow(ct);
            _rowWriter ??= new XlsxRowWriter(this, _rowBuffer);
            _rowWriter.Reset(rowNumber);
            return _rowWriter;
        }

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The sheet has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The sheet has not been started, or the active <see cref="XlsxRowWriter"/> has not been disposed.</exception>
        public async ValueTask EndAsync(CancellationToken ct = default)
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireStarted(_state, nameof(XlsxSheetWriter), "ending");
            WriterStateGuard.RequireNoActiveRowForEnd(_rowActive, nameof(XlsxRowWriter));
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

        /// <inheritdoc/>
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
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireStarted(_state, nameof(XlsxSheetWriter), "adding rows");
            WriterStateGuard.RequireNoActiveRowForStart(_rowActive, nameof(XlsxRowWriter));
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
