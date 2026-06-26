using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
{
    public sealed class XlsbSheetWriter : ISheetWriter<XlsbRowWriter>
    {
        private const int SpillThreshold = 8 * 1024 * 1024;

        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "XlsbWorkbookWriter is borrowed; its lifetime is managed by the caller.")]
        private readonly XlsbWorkbookWriter _owner;
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "ZipArchive is borrowed from XlsbWorkbookWriter; its lifetime exceeds this sheet.")]
        private readonly ZipArchive _zip;
        private readonly bool _date1904;
        private readonly CompressionLevel _compression;
        private readonly BiffBuffer _records = new(4096);
        private readonly BiffBuffer _payload = new(256);
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "Stream is explicitly disposed in EndAsync or DisposeAsync.")]
        private Stream? _stream;
        private WriterState _state = WriterState.Created;
        private bool _rowActive;
        private bool _registered;
        private bool _buffersDisposed;

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
        internal BiffBuffer Payload => _payload;

        public ValueTask StartAsync(CancellationToken ct = default)
        {
            ObjectDisposedException.ThrowIf(_state == WriterState.Ended, this);
            if (_state != WriterState.Created)
            {
                throw new InvalidOperationException("XlsbSheetWriter has already been started.");
            }
            ct.ThrowIfCancellationRequested();
            _state = WriterState.Started;
            return ValueTask.CompletedTask;
        }

        public ValueTask<XlsbRowWriter> StartRowAsync(CancellationToken ct = default)
        {
            BeginRow();
            return ValueTask.FromResult(new XlsbRowWriter(this, _date1904));
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

        public void WriteRow<T>(ReadOnlySpan<T> values)
            where T : ISpanFormattable
        {
            BeginRow();
            for (int i = 0; i < values.Length; i++)
            {
                WriteDoubleCell(i, XlsbRowWriter.ToDouble(values[i]), style: 0);
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
            _payload.Dispose();
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
            WriteRecord(Brt.RowHdr);
            _rowActive = true;
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

        private void WriteStringCell(int columnIndex, string? value)
        {
            if (value is null)
            {
                return;
            }
            _payload.Reset();
            Biff12RecordWriter.WriteCellHeader(_payload, columnIndex, 0);
            Biff12RecordWriter.WriteWideString(_payload, value);
            WriteRecord(Brt.CellSt, _payload.Span);
        }

        private void WriteBoolCell(int columnIndex, bool value)
        {
            _payload.Reset();
            Biff12RecordWriter.WriteCellHeader(_payload, columnIndex, 0);
            _payload.WriteByte(value ? (byte)1 : (byte)0);
            WriteRecord(Brt.CellBool, _payload.Span);
        }

        private void WriteDateSerialCell(int columnIndex, double serial)
        {
            if (_date1904)
            {
                serial -= 1462.0;
            }
            WriteDoubleCell(columnIndex, serial, style: 1);
        }

        private void WriteDoubleCell(int columnIndex, double value, int style)
        {
            _payload.Reset();
            Biff12RecordWriter.WriteCellHeader(_payload, columnIndex, style);
            _payload.WriteDouble(value);
            WriteRecord(Brt.CellReal, _payload.Span);
        }
    }
}
