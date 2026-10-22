using System.Buffers.Text;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Internal;
using ExcelReader.Core.Writer.Internal;
using ExcelReader.Core.Writer.Parallel;
using ExcelReader.Core.Writer.Zip;

namespace ExcelReader.Core.Writer.Xlsx
{
    /// <summary>
    /// Writes a single worksheet's XML into an XLSX ZIP archive, buffering rows and flushing them to the entry stream once they cross a size threshold.
    /// </summary>
    public sealed class XlsxSheetWriter : ISheetWriter<XlsxRowWriter>, ITableSheetWriter<XlsxRowWriter>, IBatchSheet<XlsxSheetWriter, XlsxRowWriter>, IZipSheet
    {
        private readonly TableTracker _tables = new();
        private readonly XlsxWorkbookWriter _owner;
        private readonly ZipPackageWriter _zip;
        private readonly CompressionLevel _compression;
        private readonly bool _offloadWrite;
        private readonly BiffBuffer _rowBuffer;
        // ponytail: flush to the deflate stream once buffered rows pass 64 KB — bounds memory on huge
        private const int FlushThreshold = 64 * 1024;
        private readonly int _flushThreshold = FlushThreshold;
        private readonly WorkerStringCache? _workerStrings;
        [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed", Justification = "An alias of the entry that _stream is or wraps; disposing _stream disposes it.")]
        private ZipEntryStream? _entry;
        private XlsxRowWriter? _rowWriter;
        private Stream? _stream;
        private int _rowNumber;
        private WriterState _state = WriterState.Created;
        private bool _rowActive;
        private Dictionary<int, int>? _columnStyles;
        private Dictionary<int, double>? _columnWidths;

        internal XlsxSheetWriter(XlsxWorkbookWriter owner, ZipPackageWriter zip, string name, int sheetId,
            ExcelSheetVisibility visibility, CompressionLevel compression, bool offloadWrite)
        {
            _owner = owner;
            _zip = zip;
            Name = name;
            SheetId = sheetId;
            Visibility = visibility;
            _compression = compression;
            _offloadWrite = offloadWrite;
            _rowBuffer = new BiffBuffer(512);
        }

        // A detached copy formats batches for the parallel record path: it has its own row buffer,
        // never flushes it and never touches the ZIP.
        private XlsxSheetWriter(XlsxSheetWriter source, int batchRows)
        {
            _rowBuffer = new BiffBuffer(DetachedCapacity(batchRows));
            _owner = source._owner;
            _zip = source._zip;
            Name = source.Name;
            SheetId = source.SheetId;
            Visibility = source.Visibility;
            _compression = source._compression;
            _columnStyles = source._columnStyles;
            _flushThreshold = int.MaxValue;
            _state = WriterState.Started;
            if (_owner.UseSharedStrings)
            {
                _workerStrings = new WorkerStringCache(_owner.SharedStrings!);
            }
        }

        private static int DetachedCapacity(int batchRows)
        {
            return (int)Math.Min(1 << 20, (long)batchRows * 128);
        }

        internal string Name { get; }
        internal int SheetId { get; }
        internal ExcelSheetVisibility Visibility { get; }
        internal bool UseSharedStrings => _owner.UseSharedStrings;

        internal bool Date1904 => _owner.Date1904;
        internal bool ResourcesReleased => _stream is null && _rowBuffer.IsReleased;

        internal int GetSharedStringIndex(string value)
        {
            if (_workerStrings is null)
            {
                return _owner.GetSharedStringIndex(value);
            }
            return _workerStrings.GetOrAdd(value);
        }

        internal int GetSharedStringIndex(ReadOnlySpan<char> value)
        {
            if (_workerStrings is null)
            {
                return _owner.GetSharedStringIndex(value);
            }
            return _workerStrings.GetOrAdd(value);
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
            SheetColumnValidation.SetColumnStyle(ref _columnStyles, columnIndex, styleId, _owner.StyleCount, ExcelLimits.MaxColumns, _state, this);
        }

        /// <inheritdoc/>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="columnIndex"/> or <paramref name="width"/> is negative.</exception>
        /// <exception cref="InvalidOperationException">The first row has already been started.</exception>
        public void SetColumnWidth(int columnIndex, double width)
        {
            SheetColumnValidation.SetColumnWidth(ref _columnWidths, columnIndex, width, ExcelLimits.MaxColumns, _state, this);
        }

        private void EnsureStarted()
        {
            if (_state != WriterState.Created)
            {
                return;
            }
            try
            {
                OpenEntry().Write(_rowBuffer.Span);
            }
            catch
            {
                Fault();
                throw;
            }
            MarkStarted();
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
                await OpenEntry().WriteAsync(_rowBuffer.Memory, ct).ConfigureAwait(false);
            }
            catch
            {
                await FaultAsync().ConfigureAwait(false);
                throw;
            }
            MarkStarted();
        }

        // Leaves the worksheet header in the row buffer for the caller to write.
        private Stream OpenEntry()
        {
            ZipEntryStream entry = _zip.OpenEntry($"xl/worksheets/sheet{SheetId}.xml", _compression);
            _entry = entry;
            _stream = _offloadWrite ? new WriteOffloadStream(entry) : entry;
            _rowBuffer.Reset();
            _rowBuffer.WriteUtf8(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                $"<worksheet xmlns=\"{XlsxConstants.MainNs}\" xmlns:r=\"{XlsxConstants.RelationshipsNs}\">{BuildColsXml()}<sheetData>");
            return _stream;
        }

        private void MarkStarted()
        {
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
            var sb = new StringBuilder("<cols>");
            foreach (int columnIndex in SheetColumnValidation.Columns(_columnStyles, _columnWidths))
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
                WriteTrailer();
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
            _owner.RegisterTables(SheetId, _tables.Finished);
            Release(faulted: false);
        }

        private void WriteTrailer()
        {
            _tables.CloseOpen(_rowNumber - 1);
            _rowBuffer.Write("</sheetData>"u8);
            if (_tables.Finished.Count > 0)
            {
                _rowBuffer.WriteUtf8(XlsxTableXml.TableParts(_tables.Finished.Count));
            }
            _rowBuffer.Write("</worksheet>"u8);
        }

        /// <inheritdoc/>
        public void BeginTable(string name, IReadOnlyList<string> columns, ExcelTableOptions? options = null)
        {
            (string[] names, ExcelTableOptions effective) = PrepareTable(name, columns, options);
            using (XlsxRowWriter row = StartRow())
            {
                TableTracker.WriteHeader(row, effective.FirstColumn, names);
            }
            _tables.Begin(_owner.Tables.Claim(name), name, names, effective, _rowNumber - 1);
        }

        /// <inheritdoc/>
        public async ValueTask BeginTableAsync(string name, IReadOnlyList<string> columns, ExcelTableOptions? options = null, CancellationToken ct = default)
        {
            (string[] names, ExcelTableOptions effective) = PrepareTable(name, columns, options);
            XlsxRowWriter row = await StartRowAsync(ct).ConfigureAwait(false);
            await using (row.ConfigureAwait(false))
            {
                TableTracker.WriteHeader(row, effective.FirstColumn, names);
            }
            _tables.Begin(_owner.Tables.Claim(name), name, names, effective, _rowNumber - 1);
        }

        /// <inheritdoc/>
        public void EndTable()
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            if (_tables.End(_rowNumber - 1))
            {
                BeginRow(styleId: 0);
                CloseBufferedRow();
            }
        }

        private (string[] Names, ExcelTableOptions Options) PrepareTable(string name, IReadOnlyList<string> columns, ExcelTableOptions? options)
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            ExcelTableOptions effective = options ?? ExcelTableOptions.Default;
            string[] names = TableValidation.Validate(name, columns, effective);
            _tables.RequireNoneOpen();
            _owner.Tables.RequireAvailable(name);
            TableTracker.RequireRoom(_rowNumber);
            return (names, effective);
        }

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

        void IZipSheet.Abandon()
        {
            Fault();
        }

        ValueTask IZipSheet.AbandonAsync()
        {
            return FaultAsync();
        }

        private void Release(bool faulted)
        {
            _state = WriterState.Ended;
            _stream = null;
            _entry = null;
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
                WriteTrailer();
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
            _owner.RegisterTables(SheetId, _tables.Finished);
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

        private void CloseBufferedRow()
        {
            _rowBuffer.Write("</row>"u8);
            _rowActive = false;
        }

        internal void EndBufferedRow()
        {
            CloseBufferedRow();
            if (_rowBuffer.Length >= _flushThreshold)
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
            CloseBufferedRow();
            return _rowBuffer.Length >= _flushThreshold ? FlushRowBufferAsync(ct) : ValueTask.CompletedTask;
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

        CompressionLevel IBatchSheet<XlsxSheetWriter, XlsxRowWriter>.BatchCompression
        {
            get
            {
                return _compression;
            }
        }

        async ValueTask<int> IBatchSheet<XlsxSheetWriter, XlsxRowWriter>.BeginSegmentsAsync(CancellationToken ct)
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireNoActiveRowForStart(_rowActive, nameof(XlsxRowWriter));
            await EnsureStartedAsync(ct).ConfigureAwait(false);
            if (_rowBuffer.Length > 0)
            {
                await FlushRowBufferAsync(ct).ConfigureAwait(false);
            }
            if (_stream is WriteOffloadStream offload)
            {
                try
                {
                    await offload.DrainAsync(ct).ConfigureAwait(false);
                }
                catch
                {
                    await FaultAsync().ConfigureAwait(false);
                    throw;
                }
            }
            // Held for the whole stretch, so StartRow and End are refused until it is over.
            _rowActive = true;
            return _rowNumber;
        }

        async ValueTask IBatchSheet<XlsxSheetWriter, XlsxRowWriter>.AppendSegmentAsync(BatchSegment segment, CancellationToken ct)
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            await _entry!.AppendSegmentAsync(segment.Data.Memory, segment.Crc, segment.RawLength, ct).ConfigureAwait(false);
            _rowNumber += segment.Rows;
        }

        void IBatchSheet<XlsxSheetWriter, XlsxRowWriter>.EndSegments()
        {
            _rowActive = false;
        }

        async ValueTask IBatchSheet<XlsxSheetWriter, XlsxRowWriter>.FailSegmentsAsync()
        {
            _rowActive = false;
            await FaultAsync().ConfigureAwait(false);
        }

        XlsxSheetWriter IBatchSheet<XlsxSheetWriter, XlsxRowWriter>.CreateDetached(int batchRows)
        {
            return new XlsxSheetWriter(this, batchRows);
        }

        void IBatchSheet<XlsxSheetWriter, XlsxRowWriter>.BeginBatch(int rowsBefore)
        {
            _rowBuffer.Reset();
            _rowNumber = rowsBefore;
            _rowActive = false;
        }

        XlsxRowWriter IBatchSheet<XlsxSheetWriter, XlsxRowWriter>.StartBatchRow()
        {
            return StartRow();
        }

        ReadOnlySpan<byte> IBatchSheet<XlsxSheetWriter, XlsxRowWriter>.EndBatch(BatchSegment segment)
        {
            _workerStrings?.FlushReferences();
            return _rowBuffer.Span;
        }

        void IBatchSheet<XlsxSheetWriter, XlsxRowWriter>.ReleaseDetached()
        {
            _rowBuffer.Dispose();
        }
    }
}
