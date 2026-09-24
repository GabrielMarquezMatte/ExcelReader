using System.Buffers.Text;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Internal;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Xlsx
{
    /// <summary>
    /// Writes a single worksheet's XML into an XLSX ZIP archive, buffering rows and flushing them to the entry stream once they cross a size threshold.
    /// </summary>
    public sealed class XlsxSheetWriter : ISheetWriter<XlsxRowWriter>
    {
        private readonly XlsxWorkbookWriter _owner;
        private readonly ZipArchive _zip;
        private readonly CompressionLevel _compression;
        private readonly bool _offloadWrite;
        private readonly BiffBuffer _rowBuffer = new(512);
        // ponytail: flush to the deflate stream once buffered rows pass 64 KB — bounds memory on huge
        private const int FlushThreshold = 64 * 1024;
        private XlsxRowWriter? _rowWriter;
        private Stream? _stream;
        private int _rowNumber;
        private WriterState _state = WriterState.Created;
        private bool _rowActive;
        private Dictionary<int, int>? _columnStyles;
        private Dictionary<int, double>? _columnWidths;

        internal XlsxSheetWriter(XlsxWorkbookWriter owner, ZipArchive zip, string name, int sheetId,
            ExcelSheetVisibility visibility, CompressionLevel compression, bool offloadWrite)
        {
            _owner = owner;
            _zip = zip;
            Name = name;
            SheetId = sheetId;
            Visibility = visibility;
            _compression = compression;
            _offloadWrite = offloadWrite;
        }

        internal string Name { get; }
        internal int SheetId { get; }
        internal ExcelSheetVisibility Visibility { get; }
        internal bool UseSharedStrings => _owner.UseSharedStrings;
        internal bool ResourcesReleased => _stream is null && _rowBuffer.IsReleased;

        internal int GetSharedStringIndex(string value)
        {
            return _owner.GetSharedStringIndex(value);
        }

        internal int GetSharedStringIndex(ReadOnlySpan<char> value)
        {
            return _owner.GetSharedStringIndex(value);
        }

        internal int GetColumnStyle(int columnIndex)
        {
            return _columnStyles is not null && _columnStyles.TryGetValue(columnIndex, out int styleId) ? styleId : 0;
        }

        /// <inheritdoc/>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="columnIndex"/> is negative, or <paramref name="styleId"/> is negative or was never returned by <see cref="XlsxWorkbookWriter.AddStyle"/>.</exception>
        /// <exception cref="InvalidOperationException">The first row has already been started.</exception>
        public void SetColumnStyle(int columnIndex, int styleId)
        {
            SheetColumnValidation.SetColumnStyle(ref _columnStyles, columnIndex, styleId, _owner.StyleCount, _state, this);
        }

        /// <inheritdoc/>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="columnIndex"/> or <paramref name="width"/> is negative.</exception>
        /// <exception cref="InvalidOperationException">The first row has already been started.</exception>
        public void SetColumnWidth(int columnIndex, double width)
        {
            SheetColumnValidation.SetColumnWidth(ref _columnWidths, columnIndex, width, _state, this);
        }

        private void EnsureStarted()
        {
            if (_state != WriterState.Created)
            {
                return;
            }
            try
            {
                ZipArchiveEntry entry = _zip.CreateEntry($"xl/worksheets/sheet{SheetId}.xml", _compression);
                Stream stream = entry.Open();
                _stream = _offloadWrite ? new WriteOffloadStream(stream) : stream;
                _rowBuffer.Reset();
                _rowBuffer.WriteUtf8(
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    $"<worksheet xmlns=\"{XlsxConstants.MainNs}\">{BuildColsXml()}<sheetData>");
                _stream.Write(_rowBuffer.Span);
            }
            catch
            {
                Fault();
                throw;
            }
            _rowBuffer.Reset();
            _state = WriterState.Started;
            _owner.RegisterSheet(Name, SheetId, Visibility);
        }

        internal ValueTask EnsureStartedAsync(CancellationToken ct)
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            return _state == WriterState.Created ? StartCoreAsync(ct) : ValueTask.CompletedTask;
        }

        private async ValueTask StartCoreAsync(CancellationToken ct)
        {
            try
            {
                ZipArchiveEntry entry = _zip.CreateEntry($"xl/worksheets/sheet{SheetId}.xml", _compression);
                Stream stream = await entry.OpenAsync(ct).ConfigureAwait(false);
                _stream = _offloadWrite ? new WriteOffloadStream(stream) : stream;
                _rowBuffer.Reset();
                _rowBuffer.WriteUtf8(
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    $"<worksheet xmlns=\"{XlsxConstants.MainNs}\">{BuildColsXml()}<sheetData>");
                await _stream.WriteAsync(_rowBuffer.Memory, ct).ConfigureAwait(false);
            }
            catch
            {
                await FaultAsync().ConfigureAwait(false);
                throw;
            }
            _rowBuffer.Reset();
            _state = WriterState.Started;
            _owner.RegisterSheet(Name, SheetId, Visibility);
        }

        private string BuildColsXml()
        {
            if (_columnStyles is null && _columnWidths is null)
            {
                return string.Empty;
            }
            var columns = new SortedSet<int>();
            if (_columnStyles is not null)
            {
                columns.UnionWith(_columnStyles.Keys);
            }
            if (_columnWidths is not null)
            {
                columns.UnionWith(_columnWidths.Keys);
            }
            var sb = new StringBuilder("<cols>");
            foreach (int columnIndex in columns)
            {
                int oneBased = columnIndex + 1;
                sb.Append(CultureInfo.InvariantCulture, $"<col min=\"{oneBased}\" max=\"{oneBased}\"");
                if (_columnStyles is not null && _columnStyles.TryGetValue(columnIndex, out int styleId))
                {
                    sb.Append(CultureInfo.InvariantCulture, $" style=\"{styleId}\"");
                }
                if (_columnWidths is not null && _columnWidths.TryGetValue(columnIndex, out double width))
                {
                    sb.Append(CultureInfo.InvariantCulture, $" width=\"{width}\" customWidth=\"1\"");
                }
                sb.Append("/>");
            }
            sb.Append("</cols>");
            return sb.ToString();
        }

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The sheet has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The previous <see cref="XlsxRowWriter"/> has not been disposed.</exception>
        /// <exception cref="ExcelLimitExceededException">The worksheet's 1,048,576-row limit has been reached.</exception>
        public ValueTask<XlsxRowWriter> StartRowAsync(CancellationToken ct = default)
        {
            return StartRowAsync(styleId: 0, ct);
        }

        /// <inheritdoc/>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="styleId"/> is negative or was never returned by <see cref="XlsxWorkbookWriter.AddStyle"/>.</exception>
        /// <exception cref="ObjectDisposedException">The sheet has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The previous <see cref="XlsxRowWriter"/> has not been disposed.</exception>
        /// <exception cref="ExcelLimitExceededException">The worksheet's 1,048,576-row limit has been reached.</exception>
        public ValueTask<XlsxRowWriter> StartRowAsync(int styleId, CancellationToken ct = default)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(styleId);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(styleId, _owner.StyleCount);
            ct.ThrowIfCancellationRequested();
            if (_state == WriterState.Created)
            {
                return StartFirstRowAsync(styleId, ct);
            }
            return ValueTask.FromResult(StartRowCore(styleId));
        }

        private async ValueTask<XlsxRowWriter> StartFirstRowAsync(int styleId, CancellationToken ct)
        {
            await StartCoreAsync(ct).ConfigureAwait(false);
            return StartRowCore(styleId);
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="StartRowAsync(CancellationToken)"/>. Row buffering is
        /// synchronous either way, so a caller that never awaits mid-row can skip the per-row
        /// <see cref="ValueTask"/> machinery.
        /// </summary>
        /// <returns>The reusable <see cref="XlsxRowWriter"/> for the new row.</returns>
        /// <exception cref="ObjectDisposedException">The sheet has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The previous <see cref="XlsxRowWriter"/> has not been disposed.</exception>
        /// <exception cref="ExcelLimitExceededException">The worksheet's 1,048,576-row limit has been reached.</exception>
        public XlsxRowWriter StartRow()
        {
            return StartRowCore(styleId: 0);
        }

        /// <summary>Synchronous counterpart to <see cref="StartRowAsync(int, CancellationToken)"/>.</summary>
        /// <param name="styleId">The style to apply to every cell of this row.</param>
        /// <returns>The reusable <see cref="XlsxRowWriter"/> for the new row.</returns>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="styleId"/> is negative or was never returned by <see cref="XlsxWorkbookWriter.AddStyle"/>.</exception>
        /// <exception cref="ObjectDisposedException">The sheet has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The previous <see cref="XlsxRowWriter"/> has not been disposed.</exception>
        /// <exception cref="ExcelLimitExceededException">The worksheet's 1,048,576-row limit has been reached.</exception>
        public XlsxRowWriter StartRow(int styleId)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(styleId);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(styleId, _owner.StyleCount);
            return StartRowCore(styleId);
        }

        private XlsxRowWriter StartRowCore(int styleId)
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            EnsureStarted();
            int rowNumber = BeginRow(styleId);
            _rowWriter ??= new XlsxRowWriter(this, _rowBuffer);
            _rowWriter.Reset(rowNumber, styleId);
            return _rowWriter;
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="EndAsync"/>, for native/unmanaged callers whose ABI is
        /// synchronous.
        /// </summary>
        /// <exception cref="ObjectDisposedException">The sheet has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The active <see cref="XlsxRowWriter"/> has not been disposed.</exception>
        public void End()
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireNoActiveRowForEnd(_rowActive, nameof(XlsxRowWriter));
            try
            {
                EnsureStarted();
                _rowBuffer.Write("</sheetData></worksheet>"u8);
#pragma warning disable CS8602
                _stream.Write(_rowBuffer.Span);
                _stream.Flush();
                _stream.Dispose();
#pragma warning restore CS8602
            }
            catch
            {
                Fault();
                throw;
            }
            Release(faulted: false);
        }

        // Ended is only ever set by Release, after cleanup, so an already-ended sheet has nothing left
        // to release: a fault raised inside a nested step (start, flush) is handled exactly once.
        private void Fault()
        {
            if (_state == WriterState.Ended)
            {
                return;
            }
            FailureCleanup.Dispose(_stream);
            Release(faulted: true);
        }

        private async ValueTask FaultAsync()
        {
            if (_state == WriterState.Ended)
            {
                return;
            }
            await FailureCleanup.DisposeAsync(_stream).ConfigureAwait(false);
            Release(faulted: true);
        }

        private void Release(bool faulted)
        {
            _state = WriterState.Ended;
            _stream = null;
            _rowBuffer.Dispose();
            _owner.NotifySheetEnded(faulted);
        }

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The sheet has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The active <see cref="XlsxRowWriter"/> has not been disposed.</exception>
        public async ValueTask EndAsync(CancellationToken ct = default)
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireNoActiveRowForEnd(_rowActive, nameof(XlsxRowWriter));
            ct.ThrowIfCancellationRequested();
            try
            {
                if (_state == WriterState.Created)
                {
                    await StartCoreAsync(ct).ConfigureAwait(false);
                }
                _rowBuffer.Write("</sheetData></worksheet>"u8);
#pragma warning disable CS8602
                await _stream.WriteAsync(_rowBuffer.Memory, ct).ConfigureAwait(false);
#pragma warning restore CS8602
                await _stream.FlushAsync(ct).ConfigureAwait(false);
                await _stream.DisposeAsync().ConfigureAwait(false);
            }
            catch
            {
                await FaultAsync().ConfigureAwait(false);
                throw;
            }
            Release(faulted: false);
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="DisposeAsync"/>, for native/unmanaged callers whose ABI
        /// is synchronous.
        /// </summary>
        public void Dispose()
        {
            if (_state != WriterState.Ended)
            {
                End();
            }
        }

        /// <inheritdoc/>
        public async ValueTask DisposeAsync()
        {
            if (_state != WriterState.Ended)
            {
                await EndAsync().ConfigureAwait(false);
            }
        }

        private int BeginRow(int styleId)
        {
            WriterStateGuard.RequireNoActiveRowForStart(_rowActive, nameof(XlsxRowWriter));
            if (_rowNumber >= ExcelLimits.MaxRows)
            {
                ExcelLimits.ThrowRowLimit(_rowNumber + 1L);
            }
            _rowNumber++;
            _rowActive = true;
            if (styleId == 0)
            {
                _rowBuffer.Write("<row>"u8);
            }
            else
            {
                _rowBuffer.Write("<row s=\""u8);
                Utf8Formatter.TryFormat(styleId, _rowBuffer.GetSpan(11), out int written);
                _rowBuffer.Advance(written);
                _rowBuffer.Write("\" customFormat=\"1\">"u8);
            }
            return _rowNumber;
        }

        internal void EndBufferedRow()
        {
            _rowBuffer.Write("</row>"u8);
            _rowActive = false;
            if (_rowBuffer.Length >= FlushThreshold)
            {
                FlushRowBuffer();
            }
        }

        private void FlushRowBuffer()
        {
            try
            {
                if (_stream is WriteOffloadStream offload)
                {
                    byte[] detached = _rowBuffer.Detach(out int length);
                    offload.EnqueueOwned(detached, length);
                    return;
                }
                _stream!.Write(_rowBuffer.Span);
                _rowBuffer.Reset();
            }
            catch
            {
                Fault();
                throw;
            }
        }

        internal ValueTask EndBufferedRowAsync(CancellationToken ct = default)
        {
            _rowBuffer.Write("</row>"u8);
            _rowActive = false;
            return _rowBuffer.Length >= FlushThreshold ? FlushRowBufferAsync(ct) : ValueTask.CompletedTask;
        }

        private async ValueTask FlushRowBufferAsync(CancellationToken ct)
        {
            try
            {
                if (_stream is WriteOffloadStream offload)
                {
                    byte[] detached = _rowBuffer.Detach(out int length);
                    await offload.EnqueueOwnedAsync(detached, length, ct).ConfigureAwait(false);
                    return;
                }
#pragma warning disable CS8602
                await _stream.WriteAsync(_rowBuffer.Memory, ct).ConfigureAwait(false);
#pragma warning restore CS8602
                _rowBuffer.Reset();
            }
            catch
            {
                await FaultAsync().ConfigureAwait(false);
                throw;
            }
        }
    }
}
