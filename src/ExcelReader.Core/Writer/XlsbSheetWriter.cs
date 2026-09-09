using System.Buffers.Binary;
using System.IO.Compression;
using System.Runtime.InteropServices;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
{
    /// <summary>Writes a single worksheet's rows into an .xlsb workbook produced by <see cref="XlsbWorkbookWriter"/>.</summary>
    public sealed class XlsbSheetWriter : ISheetWriter<XlsbRowWriter>
    {
        // Kept under the LOH threshold instead of parking the pooled backing array there permanently.
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

        internal XlsbSheetWriter(
            XlsbWorkbookWriter owner,
            ZipArchive zip,
            string name,
            int sheetId,
            bool date1904,
            CompressionLevel compression,
            bool offloadWrite)
        {
            _owner = owner;
            _zip = zip;
            Name = name;
            SheetId = sheetId;
            _date1904 = date1904;
            _compression = compression;
            _offloadWrite = offloadWrite;
        }

        internal string Name { get; }
        internal int SheetId { get; }
        internal BiffBuffer Payload { get; } = new(256);
        internal bool UseSharedStrings => _owner.UseSharedStrings;

        internal int GetSharedStringIndex(string value)
        {
            return _owner.GetSharedStringIndex(value);
        }

        /// <inheritdoc/>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="columnIndex"/> is negative, or <paramref name="styleId"/> is negative or was never returned by <see cref="XlsbWorkbookWriter.AddStyle"/>.</exception>
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

        // A row style always wins over a column style; falls back to 0. Every BIFF12 cell record
        // carries a mandatory ixfe field, so this is consulted for every cell write, not only dates.
        private int EffectiveStyle(int columnIndex)
        {
            if (_activeRowStyle != 0)
            {
                return _activeRowStyle;
            }
            return _columnStyles is not null && _columnStyles.TryGetValue(columnIndex, out int styleId) ? styleId : 0;
        }

        private void StartCore()
        {
            _state = WriterState.Started;
            WriteRecord(Brt.BeginSheet);
            WriteWorksheetView();
            WriteRecord(Brt.BeginColInfos);
            WriteColInfos();
            WriteRecord(Brt.EndColInfos);
            WriteRecord(Brt.BeginSheetData);
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="StartAsync"/>, for native/unmanaged callers whose ABI
        /// is synchronous.
        /// </summary>
        public void Start()
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireCreated(_state, nameof(XlsbSheetWriter));
            StartCore();
        }

        /// <inheritdoc/>
        public ValueTask StartAsync(CancellationToken ct = default)
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireCreated(_state, nameof(XlsbSheetWriter));
            ct.ThrowIfCancellationRequested();
            StartCore();
            return ValueTask.CompletedTask;
        }

        // Excel's own default column width, in characters.
        private const double DefaultColumnWidth = 8.43;

        // BrtColInfo flags: bit 1 is fUserSet. Excel ignores coldx on a column that doesn't claim it,
        // so an explicit SetColumnWidth must set this or the width silently has no effect.
        private const int ColInfoUserSet = 0x0002;

        // Payload layout verified byte-for-byte against a real Excel-authored .xlsb (18 bytes:
        // colFirst/colLast/coldx/ixfe as u32, then flags as u16 — ixfe is a u32, not u16).
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
                Payload.WriteU32((uint)columnIndex);                 // colFirst
                Payload.WriteU32((uint)columnIndex);                 // colLast
                Payload.WriteU32((uint)Math.Round(width * 256));     // coldx (1/256th of a character)
                Payload.WriteU32((uint)styleId);                     // ixfe
                Payload.WriteU16(hasWidth ? ColInfoUserSet : 0);     // flags
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
            _rowActive = false;
        }

        internal void NotifyRowEnded()
        {
            _rowActive = false;
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="EndAsync"/>, for native/unmanaged callers whose ABI is
        /// synchronous.
        /// </summary>
        public void End()
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireStarted(_state, nameof(XlsbSheetWriter), "ending");
            WriterStateGuard.RequireNoActiveRowForEnd(_rowActive, nameof(XlsbRowWriter));
            _state = WriterState.Ended;
            WriteRecord(Brt.EndSheetData);
            WriteSheetMetadata();
            WriteRecord(Brt.EndSheet);
            if (_stream is null)
            {
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
            WriterStateGuard.RequireStarted(_state, nameof(XlsbSheetWriter), "ending");
            WriterStateGuard.RequireNoActiveRowForEnd(_rowActive, nameof(XlsbRowWriter));
            ct.ThrowIfCancellationRequested();
            _state = WriterState.Ended;
            WriteRecord(Brt.EndSheetData);
            WriteSheetMetadata();
            WriteRecord(Brt.EndSheet);
            if (_stream is null)
            {
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
            if (_state == WriterState.Started)
            {
                End();
            }
            else if (_state == WriterState.Created)
            {
                ReleaseBuffers();
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
                ReleaseBuffers();
            }
        }

        internal void WriteRecord(int id, ReadOnlySpan<byte> payload = default)
        {
            Biff12RecordWriter.WriteRecord(_records, id, payload);
            MaybeFlush();
        }

        private void MaybeFlush()
        {
            if (_records.Length >= SpillThreshold)
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

        // When offloading, hands the buffer's array to the background writer directly instead of
        // copying; Detach rents _records its own replacement so it keeps working immediately.
        private void FlushRecords()
        {
            if (_records.Length == 0)
            {
                return;
            }
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
#if NET10_0_OR_GREATER
            Stream stream = await entry.OpenAsync(ct).ConfigureAwait(false);
#else
            Stream stream = entry.Open();
#endif
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
            WriterStateGuard.RequireStarted(_state, nameof(XlsbSheetWriter), "adding rows");
            WriterStateGuard.RequireNoActiveRowForStart(_rowActive, nameof(XlsbRowWriter));
            if (_rowNumber >= ExcelLimits.MaxRows)
            {
                ExcelLimits.ThrowRowLimit(_rowNumber + 1L);
            }
            _rowNumber++;
            _activeRowStyle = styleId;
            WriteRowHeader(_rowNumber);
            _rowActive = true;
        }

        private void WriteRowHeader(int rowNumber)
        {
            const int Length = (6 * 4) + 1; // 6 x u32 + 1 byte
            Biff12RecordWriter.WriteFixedRecord(_records, Brt.RowHdr, Length, out Span<byte> p);
            BinaryPrimitives.WriteUInt32LittleEndian(p, (uint)rowNumber);
            BinaryPrimitives.WriteUInt32LittleEndian(p.Slice(4, 4), 0);
            BinaryPrimitives.WriteUInt32LittleEndian(p.Slice(8, 4), 0);
            BinaryPrimitives.WriteUInt32LittleEndian(p.Slice(12, 4), 1);
            BinaryPrimitives.WriteUInt32LittleEndian(p.Slice(16, 4), 0);
            BinaryPrimitives.WriteUInt32LittleEndian(p.Slice(20, 4), 16383);
            p[24] = 0;
            MaybeFlush();
        }
        // Fixed byte blobs Excel expects verbatim; nothing in them varies per workbook.
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

        private const int CellHeaderLength = 8; // column u32 + style u32
        private const int RkIntMin = -(1 << 29);  // RkNumber's fInt field is a 30-bit signed integer
        private const int RkIntMax = (1 << 29) - 1;

        internal void WriteStringCell(int columnIndex, string? value)
        {
            ValidateColumn(columnIndex);
            if (value is null)
            {
                return;
            }
            ExcelLimits.ThrowIfCellTextTooLong(value.Length, nameof(value));
            int style = EffectiveStyle(columnIndex);
            if (_owner.UseSharedStrings)
            {
                const int Length = CellHeaderLength + 4; // + shared-string index u32
                Biff12RecordWriter.WriteFixedRecord(_records, Brt.CellIsst, Length, out Span<byte> shared);
                Biff12RecordWriter.WriteCellHeader(shared, columnIndex, style);
                BinaryPrimitives.WriteUInt32LittleEndian(shared.Slice(8, 4), (uint)_owner.GetSharedStringIndex(value));
                MaybeFlush();
                return;
            }
            // Wide-string payload is u32 length + UTF-16LE chars: an exact, known-upfront byte count.
            int length = CellHeaderLength + 4 + checked(value.Length * 2);
            Biff12RecordWriter.WriteFixedRecord(_records, Brt.CellSt, length, out Span<byte> p);
            Biff12RecordWriter.WriteCellHeader(p, columnIndex, style);
            BinaryPrimitives.WriteUInt32LittleEndian(p.Slice(8, 4), (uint)value.Length);
            MemoryMarshal.AsBytes(value.AsSpan()).CopyTo(p[12..]);
            MaybeFlush();
        }

        internal void WriteBoolCell(int columnIndex, bool value)
        {
            ValidateColumn(columnIndex);
            const int Length = CellHeaderLength + 1; // + bool byte
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
            if (TryEncodeRkInt(value, out uint rk))
            {
                const int RkLength = CellHeaderLength + 4; // + RkNumber u32
                Biff12RecordWriter.WriteFixedRecord(_records, Brt.CellRk, RkLength, out Span<byte> rkPayload);
                Biff12RecordWriter.WriteCellHeader(rkPayload, columnIndex, style);
                BinaryPrimitives.WriteUInt32LittleEndian(rkPayload.Slice(8, 4), rk);
                MaybeFlush();
                return;
            }
            const int Length = CellHeaderLength + 8; // + double
            Biff12RecordWriter.WriteFixedRecord(_records, Brt.CellReal, Length, out Span<byte> p);
            Biff12RecordWriter.WriteCellHeader(p, columnIndex, style);
            BinaryPrimitives.WriteDoubleLittleEndian(p.Slice(8, 8), value);
            MaybeFlush();
        }

        // RkNumber ([MS-XLSB] 2.5.122) with fInt set: its upper 30 bits are a signed integer, so a
        // whole number in [-2^29, 2^29) costs 4 payload bytes instead of Xnum's 8 — the encoding Excel
        // emits for integers, and the one Biff12.Rk reads back on the way in. A sheet of integers is
        // about half the size for it. The fX100 form is deliberately left out: it would trade the same
        // 4 bytes for a divide-by-100 round trip that is not bit-exact for every multiple of 0.01.
        // Negative zero is excluded too — fInt has no sign bit for zero, so RK would hand back +0.0.
        private static bool TryEncodeRkInt(double value, out uint rk)
        {
            rk = 0;
            if (value != Math.Truncate(value) || value < RkIntMin || value > RkIntMax
                || (value == 0 && double.IsNegative(value)))
            {
                return false;
            }
            rk = ((uint)(int)value << 2) | 0x02; // fInt set, fX100 clear
            return true;
        }

        private static void ValidateColumn(int columnIndex)
        {
            ExcelLimits.ThrowIfColumnOutOfRange(columnIndex);
        }
    }
}
