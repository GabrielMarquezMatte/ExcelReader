using System.Buffers.Binary;
using System.IO.Compression;
using System.Runtime.InteropServices;
using ExcelReader.Core.Reader.Internal;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Xlsb
{
    /// <summary>Writes a single worksheet's rows into an .xlsb workbook produced by <see cref="XlsbWorkbookWriter"/>.</summary>
    public sealed class XlsbSheetWriter : ISheetWriter<XlsbRowWriter>
    {
        private const int SpillThreshold = 64 * 1024;

        private readonly XlsbWorkbookWriter _owner;
        private readonly ZipArchive _zip;
        private readonly bool _date1904;
        private readonly CompressionLevel _compression;
        private readonly bool _offloadWrite;
        private readonly BiffBuffer _records = new(4096);
        private Stream? _stream;
        private WriterState _state = WriterState.Created;
        private bool _rowActive;
        private bool _registered;
        private bool _buffersDisposed;
        private int _rowNumber = -1;
        private XlsbRowWriter? _rowWriter;
        private Dictionary<int, int>? _columnStyles;
        private Dictionary<int, double>? _columnWidths;
        private int _activeRowStyle;
        private int _dimPayloadPos = -1;
        private int _usedRowFirst = -1;
        private int _usedRowLast = -1;
        private int _usedColFirst = int.MaxValue;
        private int _usedColLast = -1;
        private int _rowHeaderPos;
        private int _rowColFirst;
        private int _rowColLast;

        internal XlsbSheetWriter(
            XlsbWorkbookWriter owner,
            ZipArchive zip,
            string name,
            int sheetId,
            ExcelSheetVisibility visibility,
            bool date1904,
            CompressionLevel compression,
            bool offloadWrite)
        {
            _owner = owner;
            _zip = zip;
            Name = name;
            SheetId = sheetId;
            Visibility = visibility;
            _date1904 = date1904;
            _compression = compression;
            _offloadWrite = offloadWrite;
        }

        internal string Name { get; }
        internal int SheetId { get; }
        internal ExcelSheetVisibility Visibility { get; }
        internal BiffBuffer Payload { get; } = new(256);
        internal bool UseSharedStrings => _owner.UseSharedStrings;

        /// <inheritdoc/>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="columnIndex"/> is negative, or <paramref name="styleId"/> is negative or was never returned by <see cref="XlsbWorkbookWriter.AddStyle"/>.</exception>
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

        private int EffectiveStyle(int columnIndex)
        {
            if (_activeRowStyle != 0)
            {
                return _activeRowStyle;
            }
            return _columnStyles is not null && _columnStyles.TryGetValue(columnIndex, out int styleId) ? styleId : 0;
        }

        private void EnsureStarted()
        {
            if (_state != WriterState.Created)
            {
                return;
            }
            _state = WriterState.Started;
            WriteRecord(Brt.BeginSheet);
            Biff12RecordWriter.WriteFixedRecord(_records, Brt.WsDim, DimLength, out Span<byte> dim);
            dim.Clear();
            _dimPayloadPos = _records.Length - DimLength;
            WriteWorksheetView();
            WriteRecord(Brt.BeginColInfos);
            WriteColInfos();
            WriteRecord(Brt.EndColInfos);
            WriteRecord(Brt.BeginSheetData);
        }

        private const double DefaultColumnWidth = 8.43;

        private const int DimLength = 16;

        private const int ColInfoUserSet = 0x0002;

        private void WriteColInfos()
        {
            if (_columnStyles is null && _columnWidths is null)
            {
                return;
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
            foreach (int columnIndex in columns)
            {
                int styleId = _columnStyles is not null && _columnStyles.TryGetValue(columnIndex, out int s) ? s : 0;
                bool hasWidth = false;
                double width = DefaultColumnWidth;
                if (_columnWidths is not null && _columnWidths.TryGetValue(columnIndex, out double explicitWidth))
                {
                    hasWidth = true;
                    width = explicitWidth;
                }
                Payload.Reset();
                Payload.WriteU32((uint)columnIndex);
                Payload.WriteU32((uint)columnIndex);
                Payload.WriteU32((uint)Math.Round(width * 256));
                Payload.WriteU32((uint)styleId);
                Payload.WriteU16(hasWidth ? ColInfoUserSet : 0);
                WriteRecord(Brt.ColInfo, Payload.Span);
            }
        }

        /// <inheritdoc/>
        public ValueTask<XlsbRowWriter> StartRowAsync(CancellationToken ct = default)
        {
            return StartRowAsync(styleId: 0, ct);
        }

        /// <inheritdoc/>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="styleId"/> is negative or was never returned by <see cref="XlsbWorkbookWriter.AddStyle"/>.</exception>
        public ValueTask<XlsbRowWriter> StartRowAsync(int styleId, CancellationToken ct = default)
        {
            return ValueTask.FromResult(StartRow(styleId));
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="ISheetWriter{TRow}.StartRow()"/>, for native/unmanaged
        /// callers whose ABI is synchronous.
        /// </summary>
        public XlsbRowWriter StartRow()
        {
            return StartRow(styleId: 0);
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="ISheetWriter{TRow}.StartRow(int)"/>, for native/unmanaged
        /// callers whose ABI is synchronous.
        /// </summary>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="styleId"/> is negative or was never returned by <see cref="XlsbWorkbookWriter.AddStyle"/>.</exception>
        public XlsbRowWriter StartRow(int styleId)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(styleId);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(styleId, _owner.StyleCount);
            BeginRow(styleId);
            _rowWriter ??= new XlsbRowWriter(this);
            _rowWriter.Reset();
            return _rowWriter;
        }

        /// <summary>Writes an entire row in one call, mapping each element of <paramref name="values"/> to a column starting at 0.</summary>
        public void WriteRow(ReadOnlySpan<XlsbCell> values)
        {
            BeginRow(styleId: 0);
            for (int i = 0; i < values.Length; i++)
            {
                WriteCell(i, values[i]);
            }
            FinishRow();
        }

        internal void NotifyRowEnded()
        {
            FinishRow();
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="EndAsync"/>, for native/unmanaged callers whose ABI is
        /// synchronous.
        /// </summary>
        public void End()
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireNoActiveRowForEnd(_rowActive, nameof(XlsbRowWriter));
            EnsureStarted();
            _state = WriterState.Ended;
            WriteRecord(Brt.EndSheetData);
            WriteSheetMetadata();
            WriteRecord(Brt.EndSheet);
            if (_stream is null)
            {
                PatchDimension(final: true);
                WriteBufferedSheet();
            }
            else
            {
                FlushRecords();
                _stream.Dispose();
                _stream = null;
            }
            ReleaseBuffers();
            if (!_registered)
            {
                _owner.RegisterSheet(this);
                _registered = true;
            }
            _owner.NotifySheetEnded();
        }

        /// <inheritdoc/>
        public async ValueTask EndAsync(CancellationToken ct = default)
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireNoActiveRowForEnd(_rowActive, nameof(XlsbRowWriter));
            ct.ThrowIfCancellationRequested();
            EnsureStarted();
            _state = WriterState.Ended;
            WriteRecord(Brt.EndSheetData);
            WriteSheetMetadata();
            WriteRecord(Brt.EndSheet);
            if (_stream is null)
            {
                PatchDimension(final: true);
                await WriteBufferedSheetAsync(ct).ConfigureAwait(false);
            }
            else
            {
                FlushRecords();
                await _stream.DisposeAsync().ConfigureAwait(false);
                _stream = null;
            }
            ReleaseBuffers();
            if (!_registered)
            {
                _owner.RegisterSheet(this);
                _registered = true;
            }
            _owner.NotifySheetEnded();
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

        internal void WriteRecord(int id, ReadOnlySpan<byte> payload = default)
        {
            Biff12RecordWriter.WriteRecord(_records, id, payload);
            MaybeFlush();
        }

        private void MaybeFlush()
        {
            if (!_rowActive && _records.Length >= SpillThreshold)
            {
                FlushRecords();
            }
        }

        private void EnsureStream()
        {
            if (_stream is not null)
            {
                return;
            }
            ZipArchiveEntry entry = _zip.CreateEntry($"xl/worksheets/sheet{SheetId}.bin", _compression);
            Stream stream = entry.Open();
            _stream = _offloadWrite ? new WriteOffloadStream(stream) : stream;
        }

        private void FlushRecords()
        {
            if (_records.Length == 0)
            {
                return;
            }
            PatchDimension(final: false);
            EnsureStream();
            if (_stream is WriteOffloadStream offload)
            {
                byte[] detached = _records.Detach(out int length);
                offload.EnqueueOwned(detached, length);
                return;
            }
            _stream!.Write(_records.Span);
            _records.Reset();
        }

        private void WriteBufferedSheet()
        {
            ZipArchiveEntry entry = _zip.CreateEntry($"xl/worksheets/sheet{SheetId}.bin", _compression);
            using Stream stream = entry.Open();
            stream.Write(_records.Span);
        }

        private async ValueTask WriteBufferedSheetAsync(CancellationToken ct)
        {
            ZipArchiveEntry entry = _zip.CreateEntry($"xl/worksheets/sheet{SheetId}.bin", _compression);
            Stream stream = await entry.OpenAsync(ct).ConfigureAwait(false);
            await using (stream.ConfigureAwait(false))
            {
                await stream.WriteAsync(_records.Memory, ct).ConfigureAwait(false);
            }
        }

        private void ReleaseBuffers()
        {
            if (_buffersDisposed)
            {
                return;
            }
            _buffersDisposed = true;
            _records.Dispose();
            Payload.Dispose();
        }

        private void BeginRow(int styleId)
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireNoActiveRowForStart(_rowActive, nameof(XlsbRowWriter));
            EnsureStarted();
            if (_rowNumber >= ExcelLimits.MaxRows)
            {
                ExcelLimits.ThrowRowLimit(_rowNumber + 1L);
            }
            _rowNumber++;
            _activeRowStyle = styleId;
            _rowColFirst = int.MaxValue;
            _rowColLast = -1;
            _rowHeaderPos = _records.Length;
            WriteRowHeader(_rowNumber, spanCount: 1);
            _rowActive = true;
        }

        private const int RowHeaderFixedLength = 17;
        private const int ColumnBlockShift = 10;

        private void WriteRowHeader(int rowNumber, int spanCount)
        {
            Biff12RecordWriter.WriteFixedRecord(_records, Brt.RowHdr, RowHeaderFixedLength + (spanCount * 8), out Span<byte> p);
            p.Clear();
            BinaryPrimitives.WriteUInt32LittleEndian(p, (uint)rowNumber);
            BinaryPrimitives.WriteUInt32LittleEndian(p.Slice(13, 4), (uint)spanCount);
        }

        private void FinishRow()
        {
            _rowActive = false;
            if (_rowColLast >= 0)
            {
                WriteRowSpans();
                if (_usedRowFirst < 0)
                {
                    _usedRowFirst = _rowNumber;
                }
                _usedRowLast = _rowNumber;
                _usedColFirst = Math.Min(_usedColFirst, _rowColFirst);
                _usedColLast = Math.Max(_usedColLast, _rowColLast);
            }
            MaybeFlush();
        }

        /// <summary>
        /// Patches the open row's header with the columns it actually used, one span per 1024-column
        /// block as Excel writes them. The rare multi-block row grows the header in place.
        /// </summary>
        private void WriteRowSpans()
        {
            int firstBlock = _rowColFirst >> ColumnBlockShift;
            int spanCount = (_rowColLast >> ColumnBlockShift) - firstBlock + 1;
            if (spanCount > 1)
            {
                int cellsStart = _rowHeaderPos + 2 + RowHeaderFixedLength + 8;
                int cellsLength = _records.Length - cellsStart;
                Span<byte> cells = cellsLength <= 1024 ? stackalloc byte[cellsLength] : new byte[cellsLength];
                _records.Slice(cellsStart, cellsLength).CopyTo(cells);
                _records.Truncate(_rowHeaderPos);
                WriteRowHeader(_rowNumber, spanCount);
                _records.Write(cells);
            }
            int idAndSizeLength = spanCount > 13 ? 3 : 2;
            Span<byte> spans = _records.Slice(_rowHeaderPos + idAndSizeLength + RowHeaderFixedLength, spanCount * 8);
            for (int i = 0; i < spanCount; i++)
            {
                int blockStart = (firstBlock + i) << ColumnBlockShift;
                int colMic = Math.Max(_rowColFirst, blockStart);
                int colLast = Math.Min(_rowColLast, blockStart + (1 << ColumnBlockShift) - 1);
                BinaryPrimitives.WriteUInt32LittleEndian(spans.Slice(i * 8, 4), (uint)colMic);
                BinaryPrimitives.WriteUInt32LittleEndian(spans.Slice((i * 8) + 4, 4), (uint)colLast);
            }
        }

        private void TrackColumn(int columnIndex)
        {
            if (columnIndex < _rowColFirst)
            {
                _rowColFirst = columnIndex;
            }
            if (columnIndex > _rowColLast)
            {
                _rowColLast = columnIndex;
            }
        }

        /// <summary>
        /// Fills in BrtWsDim before the sheet header leaves the buffer. A sheet that never spills gets its
        /// exact used range at End; one that spills gets the first used row and the columns seen so far,
        /// with the last row left open at the sheet maximum.
        /// </summary>
        private void PatchDimension(bool final)
        {
            if (_dimPayloadPos < 0)
            {
                return;
            }
            // ponytail: a spilled sheet's column range is what the rows before the first 64 KB flush used;
            bool any = _usedRowLast >= 0;
            int rowLast = final ? Math.Max(_usedRowLast, 0) : ExcelLimits.MaxRows - 1;
            Span<byte> dim = _records.Slice(_dimPayloadPos, DimLength);
            BinaryPrimitives.WriteUInt32LittleEndian(dim, (uint)(any ? _usedRowFirst : 0));
            BinaryPrimitives.WriteUInt32LittleEndian(dim.Slice(4, 4), (uint)rowLast);
            BinaryPrimitives.WriteUInt32LittleEndian(dim.Slice(8, 4), (uint)(any ? _usedColFirst : 0));
            BinaryPrimitives.WriteUInt32LittleEndian(dim.Slice(12, 4), (uint)(any ? _usedColLast : 0));
            _dimPayloadPos = -1;
        }
        private void WriteBlobRecord(int id, ReadOnlySpan<byte> blob)
        {
            Payload.Reset();
            Payload.Write(blob);
            WriteRecord(id, Payload.Span);
        }

        private static ReadOnlySpan<byte> InitialWorksheetViewPayload => [0x9C, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x04, 0x00, 0x00, 0x00, 0x64, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00];
        private static ReadOnlySpan<byte> SecondPayload => [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xF0, 0x3F, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x01];
        private void WriteWorksheetView()
        {
            WriteRecord(Brt.BeginWsViews);
            WriteBlobRecord(Brt.BeginWsView, InitialWorksheetViewPayload);
            WriteBlobRecord(Brt.Pane, SecondPayload);
            WriteRecord(Brt.EndWsView);
            WriteRecord(Brt.EndWsViews);
        }
        private static ReadOnlySpan<byte> SheetMetadataPayload => [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00];
        private static ReadOnlySpan<byte> TableStyleClientPayload => [0x04, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00];
        private void WriteSheetMetadata()
        {
            WriteBlobRecord(Brt.BeginCellMetadata, SheetMetadataPayload);
            WriteRecord(Brt.EndCellMetadata);
            WriteRecord(Brt.BeginTableStyles);
            WriteBlobRecord(Brt.TableStyleClient, TableStyleClientPayload);
            WriteRecord(Brt.EndTableStyles);
        }

        private void WriteCell(int columnIndex, XlsbCell cell)
        {
            switch (cell.Kind)
            {
                case XlsbCellKind.Empty:
                    break;
                case XlsbCellKind.String:
                    WriteStringCell(columnIndex, cell.Text);
                    break;
                case XlsbCellKind.Boolean:
                    WriteBoolCell(columnIndex, cell.Boolean);
                    break;
                case XlsbCellKind.Number:
                    WriteDoubleCell(columnIndex, cell.Number);
                    break;
                case XlsbCellKind.Date:
                    WriteDateSerialCell(columnIndex, cell.Number);
                    break;
                default:
                    throw new InvalidOperationException($"Unsupported XLSB cell kind: {cell.Kind}.");
            }
        }

        private const int CellHeaderLength = 8;
        private const int RkIntMin = -(1 << 29);
        private const int RkIntMax = (1 << 29) - 1;

        internal void WriteStringCell(int columnIndex, string? value)
        {
            ValidateColumn(columnIndex);
            if (value is not null)
            {
                WriteTextCell(columnIndex, value, value);
            }
        }

        internal void WriteStringCell(int columnIndex, ReadOnlySpan<char> value)
        {
            ValidateColumn(columnIndex);
            WriteTextCell(columnIndex, value, owned: null);
        }

        private void WriteTextCell(int columnIndex, ReadOnlySpan<char> value, string? owned)
        {
            ExcelLimits.ThrowIfCellTextTooLong(value.Length, nameof(value));
            TrackColumn(columnIndex);
            int style = EffectiveStyle(columnIndex);
            if (_owner.UseSharedStrings)
            {
                int index = owned is null ? _owner.GetSharedStringIndex(value) : _owner.GetSharedStringIndex(owned);
                const int Length = CellHeaderLength + 4;
                Biff12RecordWriter.WriteFixedRecord(_records, Brt.CellIsst, Length, out Span<byte> shared);
                Biff12RecordWriter.WriteCellHeader(shared, columnIndex, style);
                BinaryPrimitives.WriteUInt32LittleEndian(shared.Slice(8, 4), (uint)index);
                MaybeFlush();
                return;
            }
            int length = CellHeaderLength + 4 + checked(value.Length * 2);
            Biff12RecordWriter.WriteFixedRecord(_records, Brt.CellSt, length, out Span<byte> p);
            Biff12RecordWriter.WriteCellHeader(p, columnIndex, style);
            BinaryPrimitives.WriteUInt32LittleEndian(p.Slice(8, 4), (uint)value.Length);
            MemoryMarshal.AsBytes(value).CopyTo(p[12..]);
            MaybeFlush();
        }

        internal void WriteBoolCell(int columnIndex, bool value)
        {
            ValidateColumn(columnIndex);
            TrackColumn(columnIndex);
            const int Length = CellHeaderLength + 1;
            Biff12RecordWriter.WriteFixedRecord(_records, Brt.CellBool, Length, out Span<byte> p);
            Biff12RecordWriter.WriteCellHeader(p, columnIndex, EffectiveStyle(columnIndex));
            p[8] = value ? (byte)1 : (byte)0;
            MaybeFlush();
        }

        internal void WriteDateSerialCell(int columnIndex, double serial)
        {
            int styleId = EffectiveStyle(columnIndex);
            WriteDoubleCellCore(columnIndex, ExcelEpoch.OADateToSerial(serial, _date1904), styleId == 0 ? 1 : styleId);
        }

        internal void WriteDoubleCell(int columnIndex, double value)
        {
            WriteDoubleCellCore(columnIndex, value, EffectiveStyle(columnIndex));
        }

        private void WriteDoubleCellCore(int columnIndex, double value, int style)
        {
            ValidateColumn(columnIndex);
            CellValueGuards.ThrowIfNonFinite(value, nameof(value));
            TrackColumn(columnIndex);
            if (TryEncodeRkInt(value, out uint rk))
            {
                const int RkLength = CellHeaderLength + 4;
                Biff12RecordWriter.WriteFixedRecord(_records, Brt.CellRk, RkLength, out Span<byte> rkPayload);
                Biff12RecordWriter.WriteCellHeader(rkPayload, columnIndex, style);
                BinaryPrimitives.WriteUInt32LittleEndian(rkPayload.Slice(8, 4), rk);
                MaybeFlush();
                return;
            }
            const int Length = CellHeaderLength + 8;
            Biff12RecordWriter.WriteFixedRecord(_records, Brt.CellReal, Length, out Span<byte> p);
            Biff12RecordWriter.WriteCellHeader(p, columnIndex, style);
            BinaryPrimitives.WriteDoubleLittleEndian(p.Slice(8, 8), value);
            MaybeFlush();
        }

        private static bool TryEncodeRkInt(double value, out uint rk)
        {
            rk = 0;
            if (value != Math.Truncate(value) || value < RkIntMin || value > RkIntMax
                || (value == 0 && double.IsNegative(value)))
            {
                return false;
            }
            rk = ((uint)(int)value << 2) | 0x02;
            return true;
        }

        private static void ValidateColumn(int columnIndex)
        {
            ExcelLimits.ThrowIfColumnOutOfRange(columnIndex);
        }
    }
}
