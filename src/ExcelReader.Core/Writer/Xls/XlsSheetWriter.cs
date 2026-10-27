using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using ExcelReader.Core.Reader.Internal;
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
        [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed",
            Justification = "Read while the workbook is written; released with _cells via ReleaseBuffer.")]
        private readonly BiffBuffer _rows = new(4096);
        private int _rowCellStart;
        private int _rowColMic;
        private int _rowColMax;

        // Numbers in consecutive columns of the current row are appended to one open MULRK record, ended
        // by a gap or another kind of cell; see BiffRecordWriter.BeginRkRun.
        private int _rkRecord;
        private int _rkCount;
        private int _rkFirstCol;
        private WriterState _state = WriterState.Created;
        private bool _rowActive;
        private bool _rowHasCells;
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

        internal int RowCount => _maxRow + 1;
        internal int ColCount => _maxCol + 1;
        internal ReadOnlyMemory<byte> CellsMemory => _cells.Memory;
        internal ReadOnlyMemory<byte> ColInfoMemory => _colInfos.Memory;
        internal ReadOnlySpan<XlsRowExtent> Rows => MemoryMarshal.Cast<byte, XlsRowExtent>(_rows.Span);
        internal bool ResourcesReleased => _cells.IsReleased && _colInfos.IsReleased && _rows.IsReleased;
        private bool HasActiveRow => _rowActive || (_continuation?.HasActiveRow ?? false);

        [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed",
            Justification = "Outlives Dispose; XlsWorkbookWriter releases it via ReleaseBuffer after writing the bytes in EndAsync, same as _cells.")]
        private readonly BiffBuffer _colInfos = new(64);

        /// <inheritdoc/>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="columnIndex"/> is negative, or <paramref name="styleId"/> is negative or was never returned by <see cref="XlsWorkbookWriter.AddStyle"/>.</exception>
        /// <exception cref="InvalidOperationException">The sheet has already been started.</exception>
        public void SetColumnStyle(int columnIndex, int styleId)
        {
            SheetColumnValidation.SetColumnStyle(ref _columnStyles, columnIndex, styleId, _owner.StyleCount, MaxColumn + 1, _state, this);
        }

        /// <inheritdoc/>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="columnIndex"/> or <paramref name="width"/> is negative.</exception>
        /// <exception cref="InvalidOperationException">The sheet has already been started.</exception>
        public void SetColumnWidth(int columnIndex, double width)
        {
            SheetColumnValidation.SetColumnWidth(ref _columnWidths, columnIndex, width, MaxColumn + 1, _state, this);
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
            _rowHasCells = false;
            _rowCellStart = _cells.Length;
            _rowWriter ??= new XlsRowWriter(this, _rowNumber);
            _rowWriter.Reset(_rowNumber);
            return _rowWriter;
        }

        /// <summary>A row left without cells gets one blank cell, since BIFF8 has no other trace of it for a reader to find.</summary>
        internal void NotifyRowEnded()
        {
            FlushRkRun();
            if (_rowActive && !_rowHasCells)
            {
                BiffRecordWriter.WriteBlank(_cells, _rowNumber, 0, XlsGlobals.GeneralXf);
                Track(_rowNumber, 0);
            }
            if (_rowActive)
            {
                var extent = new XlsRowExtent(_rowCellStart, _rowColMic, _rowColMax + 1);
                MemoryMarshal.Write(_rows.GetSpan(Unsafe.SizeOf<XlsRowExtent>()), in extent);
                _rows.Advance(Unsafe.SizeOf<XlsRowExtent>());
            }
            _rowActive = false;
        }

        internal void EmitNumber(int row, int col, double value)
        {
            ValidateColumn(col);
            int abstractStyle = EffectiveStyle(col);
            WriteNumberCell(row, col, abstractStyle == 0 ? XlsGlobals.GeneralXf : XlsGlobals.CustomXf(abstractStyle), value);
            Track(row, col);
        }

        internal void EmitDate(int row, int col, DateTime value)
        {
            ValidateColumn(col);
            int abstractStyle = EffectiveStyle(col);
            double serial = ExcelEpoch.OADateToSerial(value.ToOADate(), _date1904);
            WriteNumberCell(row, col, abstractStyle == 0 ? XlsGlobals.DateXf : XlsGlobals.CustomXf(abstractStyle), serial);
            Track(row, col);
        }

        internal void EmitLabel(int row, int col, ReadOnlySpan<char> value)
        {
            ValidateColumn(col);
            ExcelLimits.ThrowIfCellTextTooLong(value.Length, nameof(value));
            int abstractStyle = EffectiveStyle(col);
            int xf = abstractStyle == 0 ? XlsGlobals.GeneralXf : XlsGlobals.CustomXf(abstractStyle);
            FlushRkRun();
            if (value.Length > XlsSharedStrings.MaxLabelLength)
            {
                BiffRecordWriter.WriteLabelSst(_cells, row, col, xf, _owner.AddSharedString(value));
            }
            else
            {
                BiffRecordWriter.WriteLabel(_cells, row, col, xf, value);
            }
            Track(row, col);
        }

        internal void EmitBool(int row, int col, bool value)
        {
            ValidateColumn(col);
            int abstractStyle = EffectiveStyle(col);
            FlushRkRun();
            BiffRecordWriter.WriteBool(_cells, row, col, abstractStyle == 0 ? XlsGlobals.GeneralXf : XlsGlobals.CustomXf(abstractStyle), value);
            Track(row, col);
        }

        private void WriteNumberCell(int row, int col, int xf, double value)
        {
            CellValueGuards.ThrowIfNonFinite(value, nameof(value));
            if (!XlsRk.TryEncode(value, out int rk))
            {
                FlushRkRun();
                BiffRecordWriter.WriteNumber(_cells, row, col, xf, value);
                return;
            }
            if (_rkCount > 0 && col != _rkFirstCol + _rkCount)
            {
                FlushRkRun();
            }
            if (_rkCount == 0)
            {
                _rkFirstCol = col;
                _rkRecord = BiffRecordWriter.BeginRkRun(_cells, row, col);
            }
            BiffRecordWriter.AppendRk(_cells, xf, rk);
            _rkCount++;
        }

        private void FlushRkRun()
        {
            if (_rkCount > 0)
            {
                BiffRecordWriter.EndRkRun(_cells, _rkRecord, _rkFirstCol, _rkCount);
                _rkCount = 0;
            }
        }

        internal void ReleaseBuffer()
        {
            _cells.Dispose();
            _colInfos.Dispose();
            _rows.Dispose();
        }

        private void Track(int row, int col)
        {
            if (!_rowHasCells)
            {
                _rowColMic = col;
            }
            _rowColMax = col;
            _rowHasCells = true;
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
            string contName;
            for (int n = _owner.SheetCount + 1; ; n++)
            {
                string suffix = $" ({n})";
                contName = _baseName.Length + suffix.Length <= MaxSheetNameLength
                    ? _baseName + suffix
                    : _baseName[..(MaxSheetNameLength - suffix.Length)] + suffix;
                if (_owner.TryClaimSheetName(contName))
                {
                    break;
                }
            }
            var cont = new XlsSheetWriter(_owner, contName, _date1904, Visibility, isContinuation: true, baseName: _baseName)
            {
                _columnStyles = _columnStyles,
                _columnWidths = _columnWidths,
            };
            cont.EnsureStarted();
            return cont;
        }

        /// <summary>
        /// Marks the sheet (and any continuation sheet it spilled into) as ended.
        /// </summary>
        public void End()
        {
            WriterStateGuard.ThrowIfEnded(_state, this);
            WriterStateGuard.RequireNoActiveRowForEnd(HasActiveRow, nameof(XlsRowWriter));
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

    [StructLayout(LayoutKind.Auto)]
    internal readonly record struct XlsRowExtent(int CellStart, int ColMic, int ColMac);
}
