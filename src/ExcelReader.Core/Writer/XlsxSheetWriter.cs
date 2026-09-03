using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO.Compression;
using System.Text;
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
        private readonly bool _offloadWrite;
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
        private Dictionary<int, int>? _columnStyles;
        private Dictionary<int, double>? _columnWidths;

        internal XlsxSheetWriter(XlsxWorkbookWriter owner, ZipArchive zip, string name, int sheetId, CompressionLevel compression, bool offloadWrite)
        {
            _owner = owner;
            _zip = zip;
            Name = name;
            SheetId = sheetId;
            _compression = compression;
            _offloadWrite = offloadWrite;
        }

        internal string Name { get; }
        internal int SheetId { get; }
        internal bool UseSharedStrings => _owner.UseSharedStrings;

        internal int GetSharedStringIndex(string value)
        {
            return _owner.GetSharedStringIndex(value);
        }

        internal int GetColumnStyle(int columnIndex)
        {
            return _columnStyles is not null && _columnStyles.TryGetValue(columnIndex, out int styleId) ? styleId : 0;
        }

        /// <inheritdoc/>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="columnIndex"/> is negative, or <paramref name="styleId"/> is negative or was never returned by <see cref="XlsxWorkbookWriter.AddStyle"/>.</exception>
        /// <exception cref="InvalidOperationException">The sheet has already been started.</exception>
        public void SetColumnStyle(int columnIndex, int styleId)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
            ArgumentOutOfRangeException.ThrowIfNegative(styleId);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(styleId, _owner.StyleCount);
            RequireNotStarted();
            _columnStyles ??= [];
            _columnStyles[columnIndex] = styleId;
        }

        /// <inheritdoc/>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="columnIndex"/> or <paramref name="width"/> is negative.</exception>
        /// <exception cref="InvalidOperationException">The sheet has already been started.</exception>
        public void SetColumnWidth(int columnIndex, double width)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
            ArgumentOutOfRangeException.ThrowIfNegative(width);
            RequireNotStarted();
            _columnWidths ??= [];
            _columnWidths[columnIndex] = width;
        }

        private void RequireNotStarted()
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            if (_state != WriterState.Created)
            {
                throw new InvalidOperationException($"{nameof(SetColumnStyle)}/{nameof(SetColumnWidth)} must be called before {nameof(StartAsync)}.");
            }
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="StartAsync"/>, for native/unmanaged callers whose ABI
        /// is synchronous.
        /// </summary>
        /// <exception cref="ObjectDisposedException">The sheet has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The sheet has already been started.</exception>
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP003:Dispose previous before re-assigning",
            Justification = "_stream is always null when Start is called (state machine guarantees Created state).")]
        public void Start()
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireCreated(_state, nameof(XlsxSheetWriter));
            ZipArchiveEntry entry = _zip.CreateEntry($"xl/worksheets/sheet{SheetId}.xml", _compression);
            Stream stream = entry.Open();
            _stream = _offloadWrite ? new WriteOffloadStream(stream) : stream;
            _rowBuffer.Reset();
            _rowBuffer.WriteUtf8(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                $"<worksheet xmlns=\"{XlsxConstants.MainNs}\">{BuildColsXml()}<sheetData>");
            _stream.Write(_rowBuffer.Span);
            _rowBuffer.Reset();
            _state = WriterState.Started;
            _owner.RegisterSheet(Name, SheetId);
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
            _stream = _offloadWrite ? new WriteOffloadStream(stream) : stream;
            _rowBuffer.Reset();
            _rowBuffer.WriteUtf8(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                $"<worksheet xmlns=\"{XlsxConstants.MainNs}\">{BuildColsXml()}<sheetData>");
            await _stream.WriteAsync(_rowBuffer.Memory, ct).ConfigureAwait(false);
            _rowBuffer.Reset();
            _state = WriterState.Started;
            _owner.RegisterSheet(Name, SheetId);
        }

        // <cols> must precede <sheetData>, so it's built once up front from whatever
        // SetColumnStyle/SetColumnWidth calls landed before StartAsync.
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
        /// <exception cref="InvalidOperationException">The sheet has not been started, or the previous <see cref="XlsxRowWriter"/> has not been disposed.</exception>
        /// <exception cref="ExcelLimitExceededException">The worksheet's 1,048,576-row limit has been reached.</exception>
        public ValueTask<XlsxRowWriter> StartRowAsync(CancellationToken ct = default)
        {
            return StartRowAsync(styleId: 0, ct);
        }

        /// <inheritdoc/>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="styleId"/> is negative or was never returned by <see cref="XlsxWorkbookWriter.AddStyle"/>.</exception>
        /// <exception cref="ObjectDisposedException">The sheet has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The sheet has not been started, or the previous <see cref="XlsxRowWriter"/> has not been disposed.</exception>
        /// <exception cref="ExcelLimitExceededException">The worksheet's 1,048,576-row limit has been reached.</exception>
        public ValueTask<XlsxRowWriter> StartRowAsync(int styleId, CancellationToken ct = default)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(styleId);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(styleId, _owner.StyleCount);
            return ValueTask.FromResult(StartRow(styleId, ct));
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
            return StartRow(styleId: 0, ct);
        }

        private XlsxRowWriter StartRow(int styleId, CancellationToken ct)
        {
            int rowNumber = BeginRow(styleId, ct);
            _rowWriter ??= new XlsxRowWriter(this, _rowBuffer);
            _rowWriter.Reset(rowNumber, styleId);
            return _rowWriter;
        }

        // Explicit, not public: a public zero-arg StartRow would collide with the existing
        // StartRow(CancellationToken ct = default) overload. Reached via the ISheetWriter<TRow>
        // constraint (see NativeApi.Write.cs).
        XlsxRowWriter ISheetWriter<XlsxRowWriter>.StartRow()
        {
            return StartRow(styleId: 0, default);
        }

        XlsxRowWriter ISheetWriter<XlsxRowWriter>.StartRow(int styleId)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(styleId);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(styleId, _owner.StyleCount);
            return StartRow(styleId, default(CancellationToken));
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="EndAsync"/>, for native/unmanaged callers whose ABI is
        /// synchronous.
        /// </summary>
        /// <exception cref="ObjectDisposedException">The sheet has already been ended.</exception>
        /// <exception cref="InvalidOperationException">The sheet has not been started, or the active <see cref="XlsxRowWriter"/> has not been disposed.</exception>
        public void End()
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireStarted(_state, nameof(XlsxSheetWriter), "ending");
            WriterStateGuard.RequireNoActiveRowForEnd(_rowActive, nameof(XlsxRowWriter));
            _state = WriterState.Ended;
            _rowBuffer.Write("</sheetData></worksheet>"u8);
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            _stream.Write(_rowBuffer.Span);
            _stream.Flush();
            _stream.Dispose();
#pragma warning restore CS8602 // Dereference of a possibly null reference.
            _stream = null;
            _rowBuffer.Dispose();
            _owner.NotifySheetEnded();
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
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            await _stream.WriteAsync(_rowBuffer.Memory, ct).ConfigureAwait(false);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
            await _stream.FlushAsync(ct).ConfigureAwait(false);
            await _stream.DisposeAsync().ConfigureAwait(false);
            _stream = null;
            _rowBuffer.Dispose();
            _owner.NotifySheetEnded();
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="DisposeAsync"/>, for native/unmanaged callers whose ABI
        /// is synchronous.
        /// </summary>
        public void Dispose()
        {
            if (_state == WriterState.Started)
            {
                End();
            }
            else if (_state == WriterState.Created)
            {
                _state = WriterState.Ended;
                _rowBuffer.Dispose();
            }
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

        private int BeginRow(int styleId, CancellationToken ct)
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireStarted(_state, nameof(XlsxSheetWriter), "adding rows");
            WriterStateGuard.RequireNoActiveRowForStart(_rowActive, nameof(XlsxRowWriter));
            ct.ThrowIfCancellationRequested();
            if (_rowNumber >= ExcelLimits.MaxRows)
            {
                ExcelLimits.ThrowRowLimit(_rowNumber + 1L);
            }
            _rowNumber++;
            _rowActive = true;
            // The `r` attribute on <row> is optional per ECMA-376 (rows are positional); omitting it
            // shrinks the XML and skips a format call on every row.
            if (styleId == 0)
            {
                _rowBuffer.Write("<row>"u8);
            }
            else
            {
                _rowBuffer.WriteUtf8($"<row s=\"{styleId}\" customFormat=\"1\">");
            }
            return _rowNumber;
        }

        internal void EndBufferedRow()
        {
            _rowBuffer.Write("</row>"u8);
            if (_rowBuffer.Length >= FlushThreshold)
            {
                FlushRowBuffer();
            }
            _rowActive = false;
        }

        // When offloading, hands the buffer's array to the background writer directly instead of
        // copying; Detach rents _rowBuffer its own replacement so it keeps working immediately.
        private void FlushRowBuffer()
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
            if (_stream is WriteOffloadStream offload)
            {
                byte[] detached = _rowBuffer.Detach(out int length);
                await offload.EnqueueOwnedAsync(detached, length, ct).ConfigureAwait(false);
                _rowActive = false;
                return;
            }
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            await _stream.WriteAsync(_rowBuffer.Memory, ct).ConfigureAwait(false);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
            _rowBuffer.Reset();
            _rowActive = false;
        }
    }
}
