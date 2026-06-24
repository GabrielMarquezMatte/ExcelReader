using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
{
    public sealed class XlsSheetWriter : IAsyncDisposable
    {
        private const int MaxRow = 65535;
        private const int MaxColumn = 255;

        // Fixed framing added around the cell records when the substream is assembled.
        private const int FramingBytes = 20 + 18 + 4; // BOF + DIMENSION + EOF

        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "XlsWorkbookWriter is borrowed; its lifetime is managed by the caller.")]
        private readonly XlsWorkbookWriter _owner;
        private readonly bool _date1904;
        private readonly BiffBuffer _cells = new();
        private int _maxRow = -1;
        private int _maxCol = -1;
        private int _rowNumber = -1;
        private WriterState _state = WriterState.Created;
        private bool _rowActive;

        internal XlsSheetWriter(XlsWorkbookWriter owner, string name, bool date1904)
        {
            _owner = owner;
            Name = name;
            _date1904 = date1904;
        }

        internal string Name { get; }

        // Full substream byte length once framed — used to compute BoundSheet offsets.
        internal int SubstreamLength => FramingBytes + _cells.Length;

        public ValueTask StartAsync(CancellationToken ct = default)
        {
            ObjectDisposedException.ThrowIf(_state == WriterState.Ended, this);
            if (_state != WriterState.Created)
            {
                throw new InvalidOperationException("XlsSheetWriter has already been started.");
            }
            ct.ThrowIfCancellationRequested();
            _state = WriterState.Started;
            _owner.RegisterSheet(this);
            return ValueTask.CompletedTask;
        }

        public ValueTask<XlsRowWriter> StartRowAsync(CancellationToken ct = default)
        {
            ObjectDisposedException.ThrowIf(_state == WriterState.Ended, this);
            if (_state != WriterState.Started)
            {
                throw new InvalidOperationException("XlsSheetWriter must be started before adding rows.");
            }
            if (_rowActive)
            {
                throw new InvalidOperationException("The previous XlsRowWriter must be disposed before starting a new row.");
            }
            ct.ThrowIfCancellationRequested();
            _rowNumber++;
            if (_rowNumber > MaxRow)
            {
                throw new InvalidOperationException($"BIFF8 worksheets are limited to {MaxRow + 1} rows.");
            }
            _rowActive = true;
            return new ValueTask<XlsRowWriter>(new XlsRowWriter(this, _rowNumber));
        }

        internal void NotifyRowEnded()
        {
            _rowActive = false;
        }

        internal void EmitNumber(int row, int col, double value)
        {
            ValidateColumn(col);
            BiffRecordWriter.WriteNumber(_cells, row, col, XlsGlobals.GeneralXf, value);
            Track(row, col);
        }

        internal void EmitDate(int row, int col, DateTime value)
        {
            ValidateColumn(col);
            double serial = value.ToOADate();
            if (_date1904)
            {
                serial -= 1462.0;
            }
            BiffRecordWriter.WriteNumber(_cells, row, col, XlsGlobals.DateXf, serial);
            Track(row, col);
        }

        internal void EmitLabel(int row, int col, ReadOnlySpan<char> value)
        {
            ValidateColumn(col);
            BiffRecordWriter.WriteLabel(_cells, row, col, XlsGlobals.GeneralXf, value);
            Track(row, col);
        }

        internal void EmitBool(int row, int col, bool value)
        {
            ValidateColumn(col);
            BiffRecordWriter.WriteBool(_cells, row, col, XlsGlobals.GeneralXf, value);
            Track(row, col);
        }

        // Writes the full worksheet substream (BOF + DIMENSION + cells + EOF) into the destination.
        internal void BuildSubstream(BiffBuffer destination)
        {
            BiffRecordWriter.WriteBof(destination, BiffRecord.SubstreamWorksheet);
            BiffRecordWriter.WriteDimension(destination, _maxRow + 1, _maxCol + 1);
            destination.Write(_cells.Span);
            BiffRecordWriter.WriteEof(destination);
        }

        internal void ReleaseBuffer()
        {
            _cells.Dispose();
        }

        private void Track(int row, int col)
        {
            if (row > _maxRow) { _maxRow = row; }
            if (col > _maxCol) { _maxCol = col; }
        }

        private static void ValidateColumn(int col)
        {
            if ((uint)col > MaxColumn)
            {
                throw new InvalidOperationException($"BIFF8 worksheets are limited to {MaxColumn + 1} columns.");
            }
        }

        public ValueTask EndAsync(CancellationToken ct = default)
        {
            ObjectDisposedException.ThrowIf(_state == WriterState.Ended, this);
            if (_state != WriterState.Started)
            {
                throw new InvalidOperationException("XlsSheetWriter must be started before ending.");
            }
            ct.ThrowIfCancellationRequested();
            _state = WriterState.Ended;
            _owner.NotifySheetEnded();
            return ValueTask.CompletedTask;
        }

        public ValueTask DisposeAsync()
        {
            if (_state == WriterState.Started)
            {
                return EndAsync();
            }
            return ValueTask.CompletedTask;
        }
    }
}
