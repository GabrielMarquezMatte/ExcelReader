using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
{
    public sealed class XlsbSheetWriter : ISheetWriter<XlsbRowWriter>
    {
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "XlsbWorkbookWriter is borrowed; its lifetime is managed by the caller.")]
        private readonly XlsbWorkbookWriter _owner;
        private readonly bool _date1904;
        private readonly BiffBuffer _records = new(4096);
        private WriterState _state = WriterState.Created;
        private bool _rowActive;
        private bool _registered;

        internal XlsbSheetWriter(XlsbWorkbookWriter owner, string name, int sheetId, bool date1904)
        {
            _owner = owner;
            Name = name;
            SheetId = sheetId;
            _date1904 = date1904;
        }

        internal string Name { get; }
        internal int SheetId { get; }
        internal ReadOnlyMemory<byte> Memory => _records.Memory;

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
            ObjectDisposedException.ThrowIf(_state == WriterState.Ended, this);
            if (_state != WriterState.Started)
            {
                throw new InvalidOperationException("XlsbSheetWriter must be started before adding rows.");
            }
            if (_rowActive)
            {
                throw new InvalidOperationException("The previous XlsbRowWriter must be disposed before starting a new row.");
            }
            ct.ThrowIfCancellationRequested();
            Biff12RecordWriter.WriteRecord(_records, Brt.RowHdr);
            _rowActive = true;
            return ValueTask.FromResult(new XlsbRowWriter(this, _records, _date1904));
        }

        internal void NotifyRowEnded()
        {
            _rowActive = false;
        }

        public ValueTask EndAsync(CancellationToken ct = default)
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
            Biff12RecordWriter.WriteRecord(_records, Brt.EndSheetData);
            if (!_registered)
            {
                _owner.RegisterSheet(this);
                _registered = true;
            }
            _owner.NotifySheetEnded();
            return ValueTask.CompletedTask;
        }

        public async ValueTask DisposeAsync()
        {
            if (_state == WriterState.Started)
            {
                await EndAsync().ConfigureAwait(false);
            }
        }

        internal void ReleaseBuffer()
        {
            _records.Dispose();
        }
    }
}