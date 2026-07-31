using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
{
    /// <summary>
    /// Writes a single sheet of a BIFF8 (.xls) workbook, buffering cell records in memory until the
    /// workbook is finalized. Rows beyond the BIFF8 65,536-row cap spill into an auto-generated
    /// continuation sheet.
    /// </summary>
    public sealed class XlsSheetWriter : IDisposable, ISheetWriter<XlsRowWriter>
    {
        private const int MaxRow = 65535;
        // Internal (not private) so XlsRowWriter.Skip can bound itself against the same limit instead
        // of duplicating the literal — BIFF8's grid is 256 columns, not the 16,384 of XLSX/XLSB.
        internal const int MaxColumn = 255;
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
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP002:Dispose member",
            Justification = "Reused per row; the caller disposes it via using after each row, and End's _rowActive guard rejects ending the sheet with it still open.")]
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

        /// <summary>
        /// Marks the sheet as started and registers it with the owning workbook.
        /// </summary>
        public void Start()
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireCreated(_state, nameof(XlsSheetWriter));
            _state = WriterState.Started;
            _owner.RegisterSheet(this);
        }

        /// <summary>
        /// Begins writing the next row, transparently spilling into a continuation sheet once the
        /// BIFF8 row cap is reached.
        /// </summary>
        public XlsRowWriter StartRow()
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireStarted(_state, nameof(XlsSheetWriter), "adding rows");
            WriterStateGuard.RequireNoActiveRowForStart(_rowActive, nameof(XlsRowWriter));
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

        /// <summary>
        /// Marks the sheet (and any continuation sheet it spilled into) as ended.
        /// </summary>
        public void End()
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireStarted(_state, nameof(XlsSheetWriter), "ending");
            WriterStateGuard.RequireNoActiveRowForEnd(_rowActive, nameof(XlsRowWriter));
            _state = WriterState.Ended;
            _continuation?.End();
            if (!_isContinuation)
            {
                _owner.NotifySheetEnded();
            }
        }

        /// <inheritdoc/>
        public void Dispose()
        {
            if (_state == WriterState.Started)
            {
                End();
            }
        }

        /// <inheritdoc/>
        /// <remarks>
        /// XLS buffers everything in memory, so this (and the other <c>*Async</c> members below) simply
        /// wraps the synchronous path in a completed <see cref="ValueTask"/>.
        /// </remarks>
        public ValueTask StartAsync(CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            Start();
            return ValueTask.CompletedTask;
        }

        /// <inheritdoc/>
        public ValueTask<XlsRowWriter> StartRowAsync(CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            return ValueTask.FromResult(StartRow());
        }

        /// <inheritdoc/>
        public ValueTask EndAsync(CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            End();
            return ValueTask.CompletedTask;
        }

        /// <inheritdoc/>
        public ValueTask DisposeAsync()
        {
            Dispose();
            return ValueTask.CompletedTask;
        }
    }
}
