using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
{
    public sealed class XlsSheetWriter : IDisposable, ISheetWriter<XlsRowWriter>
    {
        private const int MaxRow = 65535;
        private const int MaxColumn = 255;
        private const int MaxSheetNameLength = 31;

        // Fixed framing added around the cell records when the substream is assembled.
        private const int FramingBytes = 20 + 18 + 22 + 4; // BOF + DIMENSION + WINDOW2 + EOF

        private readonly XlsWorkbookWriter _owner;
        private readonly bool _date1904;
        private readonly bool _isContinuation;
        private readonly string _baseName;
        [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed",
            Justification = "The cell buffer outlives Dispose; XlsWorkbookWriter releases it via ReleaseBuffer after writing the bytes in EndAsync.")]
        private readonly BiffBuffer _cells = new();
        private XlsSheetWriter? _continuation;
        private int _maxRow = -1;
        private int _maxCol = -1;
        private int _rowNumber = -1;
        private WriterState _state = WriterState.Created;
        private bool _rowActive;
        private XlsRowWriter? _rowWriter;

        internal XlsSheetWriter(XlsWorkbookWriter owner, string name, bool date1904, bool isContinuation = false, string? baseName = null)
        {
            _owner = owner;
            Name = name;
            _date1904 = date1904;
            _isContinuation = isContinuation;
            _baseName = baseName ?? name;
        }

        internal string Name { get; }

        // Full substream byte length once framed — used to compute BoundSheet offsets.
        internal int SubstreamLength => FramingBytes + _cells.Length;

        internal int RowCount => _maxRow + 1;
        internal int ColCount => _maxCol + 1;
        internal ReadOnlyMemory<byte> CellsMemory => _cells.Memory;

        public void Start()
        {
            ObjectDisposedException.ThrowIf(_state == WriterState.Ended, this);
            if (_state != WriterState.Created)
            {
                throw new InvalidOperationException("XlsSheetWriter has already been started.");
            }
            _state = WriterState.Started;
            _owner.RegisterSheet(this);
        }

        public XlsRowWriter StartRow()
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
            _rowNumber++;
            if (_rowNumber > MaxRow)
            {
                // ponytail: auto-split into continuation sheets; BIFF8 row index is 16-bit so each sheet holds 65536 rows
                _continuation ??= CreateContinuation();
                return _continuation.StartRow();
            }
            _rowActive = true;
            _rowWriter ??= new XlsRowWriter(this, _rowNumber);
            _rowWriter.Reset(_rowNumber);
            return _rowWriter;
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
            double serial = ExcelEpoch.OADateToSerial(value.ToOADate(), _date1904);
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

        private XlsSheetWriter CreateContinuation()
        {
            string suffix = $" ({_owner.SheetCount + 1})";
            string contName = _baseName.Length + suffix.Length <= MaxSheetNameLength
                ? _baseName + suffix
                : _baseName[..(MaxSheetNameLength - suffix.Length)] + suffix;
            var cont = new XlsSheetWriter(_owner, contName, _date1904, isContinuation: true, baseName: _baseName);
            cont.Start();
            return cont;
        }

        public void End()
        {
            ObjectDisposedException.ThrowIf(_state == WriterState.Ended, this);
            if (_state != WriterState.Started)
            {
                throw new InvalidOperationException("XlsSheetWriter must be started before ending.");
            }
            _state = WriterState.Ended;
            _continuation?.End();
            if (!_isContinuation)
            {
                _owner.NotifySheetEnded();
            }
        }

        public void Dispose()
        {
            if (_state == WriterState.Started)
            {
                End();
            }
        }

        // XLS buffers everything in memory, so these wrap the synchronous path in a completed ValueTask.
        public ValueTask StartAsync(CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            Start();
            return ValueTask.CompletedTask;
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP004:Don't ignore created IDisposable",
            Justification = "The row writer is returned to the caller, who disposes it to end the row.")]
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "The row writer is returned to the caller, who disposes it to end the row.")]
        public ValueTask<XlsRowWriter> StartRowAsync(CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            return ValueTask.FromResult(StartRow());
        }

        public ValueTask EndAsync(CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            End();
            return ValueTask.CompletedTask;
        }

        public ValueTask DisposeAsync()
        {
            Dispose();
            return ValueTask.CompletedTask;
        }
    }
}
