using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Xls
{
    /// <summary>
    /// Writes a single sheet of a BIFF8 (.xls) workbook, buffering cell records in memory until the
    /// workbook is finalized. Rows beyond the BIFF8 65,536-row cap spill into an auto-generated
    /// continuation sheet.
    /// </summary>
    public sealed class XlsSheetWriter : ISheetWriter<XlsRowWriter>
    {
        private const int MaxRow = 65535;
        internal const int MaxColumn = 255;
        private const int MaxSheetNameLength = 31;

        private const int FramingBytes = 20 + 18 + 22 + 4;

        private readonly XlsWorkbookWriter _owner;
        private readonly bool _date1904;
        private readonly bool _isContinuation;
        private readonly string _baseName;
        private const int InitialCellsCapacity = 256 * 1024;

        [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed",
            Justification = "The cell buffer outlives Dispose; XlsWorkbookWriter releases it via ReleaseBuffer after writing the bytes in EndAsync.")]
        private readonly BiffBuffer _cells = new(InitialCellsCapacity);
        private XlsSheetWriter? _continuation;
        private int _maxRow = -1;
        private int _maxCol = -1;
        private int _rowNumber = -1;
        private WriterState _state = WriterState.Created;
        private bool _rowActive;
        private Dictionary<int, int>? _columnStyles;
        private Dictionary<int, double>? _columnWidths;
        private int _activeRowStyle;
        private XlsRowWriter? _rowWriter;

        internal XlsSheetWriter(XlsWorkbookWriter owner, string name, bool date1904,
            ExcelSheetVisibility visibility = ExcelSheetVisibility.Visible, bool isContinuation = false, string? baseName = null)
        {
            _owner = owner;
            Name = name;
            Visibility = visibility;
            _date1904 = date1904;
            _isContinuation = isContinuation;
            _baseName = baseName ?? name;
        }

        internal string Name { get; }

        internal ExcelSheetVisibility Visibility { get; }

        internal int SubstreamLength => FramingBytes + _colInfos.Length + _cells.Length;

        internal int RowCount => _maxRow + 1;
        internal int ColCount => _maxCol + 1;
        internal ReadOnlyMemory<byte> CellsMemory => _cells.Memory;
        internal ReadOnlyMemory<byte> ColInfoMemory => _colInfos.Memory;

        [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed",
            Justification = "Outlives Dispose; XlsWorkbookWriter releases it via ReleaseBuffer after writing the bytes in EndAsync, same as _cells.")]
        private readonly BiffBuffer _colInfos = new(64);

        /// <inheritdoc/>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="columnIndex"/> is negative, or <paramref name="styleId"/> is negative or was never returned by <see cref="XlsWorkbookWriter.AddStyle"/>.</exception>
        /// <exception cref="InvalidOperationException">The sheet has already been started.</exception>
        public void SetColumnStyle(int columnIndex, int styleId)
        {
            SheetColumnValidation.SetColumnStyle(ref _columnStyles, columnIndex, styleId, _owner.StyleCount, _state, this);
        }

        /// <inheritdoc/>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="columnIndex"/> or <paramref name="width"/> is negative.</exception>
        /// <exception cref="InvalidOperationException">The sheet has already been started.</exception>
        public void SetColumnWidth(int columnIndex, double width)
        {
            SheetColumnValidation.SetColumnWidth(ref _columnWidths, columnIndex, width, _state, this);
        }

        private int EffectiveStyle(int columnIndex)
        {
            int abstractStyle = 0;
            if (_activeRowStyle != 0)
            {
                abstractStyle = _activeRowStyle;
            }
            else if (_columnStyles is not null && _columnStyles.TryGetValue(columnIndex, out int styleId))
            {
                abstractStyle = styleId;
            }
            return abstractStyle;
        }

        private void EnsureStarted()
        {
            if (_state != WriterState.Created)
            {
                return;
            }
            _state = WriterState.Started;
            WriteColInfos();
            _owner.RegisterSheet(this);
        }

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
                int abstractStyle = _columnStyles is not null && _columnStyles.TryGetValue(columnIndex, out int s) ? s : 0;
                double width = _columnWidths is not null && _columnWidths.TryGetValue(columnIndex, out double w) ? w : 8.43;
                BiffRecordWriter.WriteColInfo(_colInfos, columnIndex, (int)Math.Round(width * 256), XlsGlobals.CustomXf(abstractStyle));
            }
        }

        /// <summary>
        /// Begins writing the next row, transparently spilling into a continuation sheet once the
        /// BIFF8 row cap is reached.
        /// </summary>
        public XlsRowWriter StartRow()
        {
            return StartRow(styleId: 0);
        }

        /// <summary>
        /// Begins writing the next row with <paramref name="styleId"/> applied to its cells,
        /// transparently spilling into a continuation sheet once the BIFF8 row cap is reached.
        /// </summary>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="styleId"/> is negative or was never returned by <see cref="XlsWorkbookWriter.AddStyle"/>.</exception>
        public XlsRowWriter StartRow(int styleId)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(styleId);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(styleId, _owner.StyleCount);
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireNoActiveRowForStart(_rowActive, nameof(XlsRowWriter));
            EnsureStarted();
            _rowNumber++;
            if (_rowNumber > MaxRow)
            {
                // ponytail: auto-split into continuation sheets; BIFF8 row index is 16-bit so each sheet holds 65536 rows
                _continuation ??= CreateContinuation();
                return _continuation.StartRow(styleId);
            }
            _activeRowStyle = styleId;
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
            int abstractStyle = EffectiveStyle(col);
            BiffRecordWriter.WriteNumber(_cells, row, col, abstractStyle == 0 ? XlsGlobals.GeneralXf : XlsGlobals.CustomXf(abstractStyle), value);
            Track(row, col);
        }

        internal void EmitDate(int row, int col, DateTime value)
        {
            ValidateColumn(col);
            int abstractStyle = EffectiveStyle(col);
            double serial = ExcelEpoch.OADateToSerial(value.ToOADate(), _date1904);
            BiffRecordWriter.WriteNumber(_cells, row, col, abstractStyle == 0 ? XlsGlobals.DateXf : XlsGlobals.CustomXf(abstractStyle), serial);
            Track(row, col);
        }

        internal void EmitLabel(int row, int col, ReadOnlySpan<char> value)
        {
            ValidateColumn(col);
            int abstractStyle = EffectiveStyle(col);
            BiffRecordWriter.WriteLabel(_cells, row, col, abstractStyle == 0 ? XlsGlobals.GeneralXf : XlsGlobals.CustomXf(abstractStyle), value);
            Track(row, col);
        }

        internal void EmitBool(int row, int col, bool value)
        {
            ValidateColumn(col);
            int abstractStyle = EffectiveStyle(col);
            BiffRecordWriter.WriteBool(_cells, row, col, abstractStyle == 0 ? XlsGlobals.GeneralXf : XlsGlobals.CustomXf(abstractStyle), value);
            Track(row, col);
        }

        internal void ReleaseBuffer()
        {
            _cells.Dispose();
            _colInfos.Dispose();
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
            var cont = new XlsSheetWriter(_owner, contName, _date1904, Visibility, isContinuation: true, baseName: _baseName);
            cont.EnsureStarted();
            return cont;
        }

        /// <summary>
        /// Marks the sheet (and any continuation sheet it spilled into) as ended.
        /// </summary>
        public void End()
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireNoActiveRowForEnd(_rowActive, nameof(XlsRowWriter));
            EnsureStarted();
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
            if (_state != WriterState.Ended)
            {
                End();
            }
        }

        /// <inheritdoc/>
        /// <remarks>
        /// XLS buffers everything in memory, so this (and the other <c>*Async</c> members below) simply
        /// wraps the synchronous path in a completed <see cref="ValueTask"/>.
        /// </remarks>
        public ValueTask<XlsRowWriter> StartRowAsync(CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            return ValueTask.FromResult(StartRow());
        }

        /// <inheritdoc/>
        public ValueTask<XlsRowWriter> StartRowAsync(int styleId, CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            return ValueTask.FromResult(StartRow(styleId));
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
