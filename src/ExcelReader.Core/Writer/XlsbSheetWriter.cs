using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
{
    public sealed class XlsbSheetWriter : ISheetWriter<XlsbRowWriter>
    {
        // Kept under the LOH threshold instead of parking the pooled backing array there permanently.
        private const int SpillThreshold = 64 * 1024;

        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "XlsbWorkbookWriter is borrowed; its lifetime is managed by the caller.")]
        private readonly XlsbWorkbookWriter _owner;
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "ZipArchive is borrowed from XlsbWorkbookWriter; its lifetime exceeds this sheet.")]
        private readonly ZipArchive _zip;
        private readonly bool _date1904;
        private readonly CompressionLevel _compression;
        private readonly BiffBuffer _records = new(4096);
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "Stream is explicitly disposed in EndAsync or DisposeAsync.")]
        private Stream? _stream;
        private WriterState _state = WriterState.Created;
        private bool _rowActive;
        private bool _registered;
        private bool _buffersDisposed;
        private int _rowNumber = -1;
        private XlsbRowWriter? _rowWriter;

        internal XlsbSheetWriter(
            XlsbWorkbookWriter owner,
            ZipArchive zip,
            string name,
            int sheetId,
            bool date1904,
            CompressionLevel compression)
        {
            _owner = owner;
            _zip = zip;
            Name = name;
            SheetId = sheetId;
            _date1904 = date1904;
            _compression = compression;
        }

        internal string Name { get; }
        internal int SheetId { get; }
        internal BiffBuffer Payload { get; } = new(256);
        internal bool UseSharedStrings => _owner.UseSharedStrings;

        internal int GetSharedStringIndex(string value)
        {
            return _owner.GetSharedStringIndex(value);
        }

        public ValueTask StartAsync(CancellationToken ct = default)
        {
            ObjectDisposedException.ThrowIf(_state == WriterState.Ended, this);
            if (_state != WriterState.Created)
            {
                throw new InvalidOperationException("XlsbSheetWriter has already been started.");
            }
            ct.ThrowIfCancellationRequested();
            _state = WriterState.Started;
            WriteRecord(Brt.BeginSheet);
            WriteWorksheetView();
            WriteRecord(Brt.BeginColInfos);
            WriteRecord(Brt.EndColInfos);
            WriteRecord(Brt.BeginSheetData);
            return ValueTask.CompletedTask;
        }

        public ValueTask<XlsbRowWriter> StartRowAsync(CancellationToken ct = default)
        {
            BeginRow();
            _rowWriter ??= new XlsbRowWriter(this);
            _rowWriter.Reset();
            return ValueTask.FromResult(_rowWriter);
        }

        public void WriteRow(ReadOnlySpan<XlsbCell> values)
        {
            BeginRow();
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

        [SuppressMessage("Reliability", "CA1849:Call async methods when in an async method",
            Justification = "The sheet body is written synchronously by row writers; EndAsync only finalizes and closes the entry.")]
        [SuppressMessage("SharpSource", "SS033:Async overload available",
            Justification = "See CA1849 justification above.")]
        public async ValueTask EndAsync(CancellationToken ct = default)
        {
            ObjectDisposedException.ThrowIf(_state == WriterState.Ended, this);
            if (_state != WriterState.Started)
            {
                throw new InvalidOperationException("XlsbSheetWriter must be started before ending.");
            }
            if (_rowActive)
            {
                throw new InvalidOperationException("The active XlsbRowWriter must be disposed before ending the sheet.");
            }
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

        [SuppressMessage("Reliability", "CA1849:Call async methods when in an async method",
            Justification = "Rows write records synchronously to keep the per-cell API synchronous.")]
        [SuppressMessage("SharpSource", "SS033:Async overload available",
            Justification = "See CA1849 justification above.")]
        internal void WriteRecord(int id, ReadOnlySpan<byte> payload = default)
        {
            Biff12RecordWriter.WriteRecord(_records, id, payload);
            MaybeFlush();
        }

        // Fixed-length records (known-size cells, row headers) write header + fields straight into
        // _records, skipping the Payload-buffer round trip that WriteRecord/Payload.Reset() needs for
        // variable-length records.
        private void MaybeFlush()
        {
            if (_records.Length >= SpillThreshold)
            {
                FlushRecords();
            }
        }

        [SuppressMessage("Reliability", "CA1849:Call async methods when in an async method",
            Justification = "Opening the entry from the synchronous row-writing hot path avoids an async API on every cell.")]
        [SuppressMessage("SharpSource", "SS033:Async overload available",
            Justification = "See CA1849 justification above.")]
        private void EnsureStream()
        {
            if (_stream is not null)
            {
                return;
            }
            ZipArchiveEntry entry = _zip.CreateEntry($"xl/worksheets/sheet{SheetId}.bin", _compression);
            _stream = entry.Open();
        }

        private void FlushRecords()
        {
            if (_records.Length == 0)
            {
                return;
            }
            EnsureStream();
            _stream!.Write(_records.Span);
            _records.Reset();
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

        private void BeginRow()
        {
            ObjectDisposedException.ThrowIf(_state == WriterState.Ended, this);
            if (_state != WriterState.Started)
            {
                throw new InvalidOperationException("XlsbSheetWriter must be started before adding rows.");
            }
            if (_rowActive)
            {
                throw new InvalidOperationException("The previous XlsbRowWriter must be disposed before starting a new row.");
            }
            _rowNumber++;
            WriteRowHeader(_rowNumber);
            _rowActive = true;
        }

        private void WriteRowHeader(int rowNumber)
        {
            const int Length = (6 * 4) + 1; // 6 x u32 + 1 byte
            Biff12RecordWriter.WriteRecordHeader(_records, Brt.RowHdr, Length);
            _records.WriteU32((uint)rowNumber);
            _records.WriteU32(0);
            _records.WriteU32(0);
            _records.WriteU32(1);
            _records.WriteU32(0);
            _records.WriteU32(16384);
            _records.WriteByte(0);
            MaybeFlush();
        }
        private static ReadOnlySpan<byte> InitialWorksheetViewPayload => [0x9C, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x04, 0x00, 0x00, 0x00, 0x64, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00];
        private static ReadOnlySpan<byte> SecondPayload => [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xF0, 0x3F, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x01];
        private void WriteWorksheetView()
        {
            WriteRecord(Brt.BeginWsViews);
            Payload.Reset();
            Payload.Write(InitialWorksheetViewPayload);
            WriteRecord(Brt.BeginWsView, Payload.Span);
            Payload.Reset();
            Payload.Write(SecondPayload);
            WriteRecord(Brt.Pane, Payload.Span);
            WriteRecord(Brt.EndWsView);
            WriteRecord(Brt.EndWsViews);
        }
        private static ReadOnlySpan<byte> SheetMetadataPayload => [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00];
        private static ReadOnlySpan<byte> TableStyleClientPayload => [0x04, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00];
        private void WriteSheetMetadata()
        {
            Payload.Reset();
            Payload.Write(SheetMetadataPayload);
            WriteRecord(Brt.BeginCellMetadata, Payload.Span);
            WriteRecord(Brt.EndCellMetadata);
            WriteRecord(Brt.BeginTableStyles);
            Payload.Reset();
            Payload.Write(TableStyleClientPayload);
            WriteRecord(Brt.TableStyleClient, Payload.Span);
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
                    WriteDoubleCell(columnIndex, cell.Number, style: 0);
                    break;
                case XlsbCellKind.Date:
                    WriteDateSerialCell(columnIndex, cell.Number);
                    break;
                default:
                    throw new InvalidOperationException($"Unsupported XLSB cell kind: {cell.Kind}.");
            }
        }

        // internal: shared with XlsbRowWriter, whose per-cell Write(...) overloads delegate here so
        // the streaming and batch (WriteRow) paths emit BIFF12 cell records through one place.
        private const int CellHeaderLength = 8; // column u32 + style u32

        internal void WriteStringCell(int columnIndex, string? value)
        {
            if (value is null)
            {
                return;
            }
            if (_owner.UseSharedStrings)
            {
                const int Length = CellHeaderLength + 4; // + shared-string index u32
                Biff12RecordWriter.WriteRecordHeader(_records, Brt.CellIsst, Length);
                Biff12RecordWriter.WriteCellHeader(_records, columnIndex, 0);
                _records.WriteU32((uint)_owner.GetSharedStringIndex(value));
                MaybeFlush();
                return;
            }
            // WriteWideString emits u32 length + UTF-16LE chars: an exact, known-upfront byte count.
            int length = CellHeaderLength + 4 + checked(value.Length * 2);
            Biff12RecordWriter.WriteRecordHeader(_records, Brt.CellSt, length);
            Biff12RecordWriter.WriteCellHeader(_records, columnIndex, 0);
            Biff12RecordWriter.WriteWideString(_records, value);
            MaybeFlush();
        }

        internal void WriteBoolCell(int columnIndex, bool value)
        {
            const int Length = CellHeaderLength + 1; // + bool byte
            Biff12RecordWriter.WriteRecordHeader(_records, Brt.CellBool, Length);
            Biff12RecordWriter.WriteCellHeader(_records, columnIndex, 0);
            _records.WriteByte(value ? (byte)1 : (byte)0);
            MaybeFlush();
        }

        internal void WriteDateSerialCell(int columnIndex, double serial)
        {
            WriteDoubleCell(columnIndex, DateSerial.ForEpoch(serial, _date1904), style: 1);
        }

        internal void WriteDoubleCell(int columnIndex, double value, int style)
        {
            const int Length = CellHeaderLength + 8; // + double
            Biff12RecordWriter.WriteRecordHeader(_records, Brt.CellReal, Length);
            Biff12RecordWriter.WriteCellHeader(_records, columnIndex, style);
            _records.WriteDouble(value);
            MaybeFlush();
        }
    }
}
