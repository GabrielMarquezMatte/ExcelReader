using ExcelReader.Core.Writer.Internal;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Writer
{
    /// <summary>
    /// Writes the cells of a single worksheet row into an XLSX row buffer; reused across rows by <see cref="XlsxSheetWriter"/>.
    /// </summary>
    public sealed class XlsxRowWriter : IRowWriter
    {
        private readonly XlsxSheetWriter _owner;
        private readonly BiffBuffer _row;
        private int _rowNumber;
        private int _columnIndex;
        private int _rowStyleId;
        private bool _useCellReferences;
        private bool _disposed;

        internal XlsxRowWriter(XlsxSheetWriter owner, BiffBuffer row)
        {
            _owner = owner;
            _row = row;
        }

        internal void Reset(int rowNumber, int styleId = 0)
        {
            _rowNumber = rowNumber;
            _columnIndex = 0;
            _rowStyleId = styleId;
            _useCellReferences = false;
            _disposed = false;
        }

        // A row style always wins over a column style; falls back to 0 (no explicit attribute).
        private int EffectiveStyle()
        {
            return _rowStyleId != 0 ? _rowStyleId : _owner.GetColumnStyle(_columnIndex);
        }

        private void ThrowIfDisposed()
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
        }

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The row has already been disposed.</exception>
        /// <exception cref="ArgumentException"><paramref name="value"/> exceeds Excel's 32,767-character per-cell limit.</exception>
        public void Write(string? value)
        {
            ThrowIfDisposed();
            if (value is null)
            {
                WriteEmptyCell();
                return;
            }
            ExcelLimits.ThrowIfCellTextTooLong(value.Length, nameof(value));
            if (_owner.UseSharedStrings)
            {
                int index = _owner.GetSharedStringIndex(value);
                CellFormatter.WriteSharedString(_row, index, _columnIndex, _rowNumber, ConsumeCellReference(), EffectiveStyle());
                _columnIndex++;
                return;
            }
            CellFormatter.WriteString(_row, value, _columnIndex, _rowNumber, ConsumeCellReference(), EffectiveStyle());
            _columnIndex++;
        }

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The row has already been disposed.</exception>
        public void Write(bool value)
        {
            ThrowIfDisposed();
            CellFormatter.WriteBool(_row, value, _columnIndex, _rowNumber, ConsumeCellReference(), EffectiveStyle());
            _columnIndex++;
        }

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The row has already been disposed.</exception>
        public void Write(bool? value)
        {
            ThrowIfDisposed();
            if (value is null)
            {
                WriteEmptyCell();
                return;
            }
            Write(value.Value);
        }

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The row has already been disposed.</exception>
        public void Write(DateTime value)
        {
            ThrowIfDisposed();
            int styleId = EffectiveStyle();
            CellFormatter.WriteDateTime(_row, value, _columnIndex, _rowNumber, ConsumeCellReference(), styleId == 0 ? 1 : styleId);
            _columnIndex++;
        }

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The row has already been disposed.</exception>
        public void Write(DateTime? value)
        {
            ThrowIfDisposed();
            if (value is null)
            {
                WriteEmptyCell();
                return;
            }
            Write(value.Value);
        }

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The row has already been disposed.</exception>
        /// <remarks>
        /// Shares the <see cref="DateTime"/> date-serial cell format (midnight), so it round-trips as a date.
        /// </remarks>
        public void Write(DateOnly value)
        {
            ThrowIfDisposed();
            int styleId = EffectiveStyle();
            CellFormatter.WriteDateTime(_row, value.ToDateTime(TimeOnly.MinValue), _columnIndex, _rowNumber, ConsumeCellReference(), styleId == 0 ? 1 : styleId);
            _columnIndex++;
        }

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The row has already been disposed.</exception>
        public void Write(DateOnly? value)
        {
            ThrowIfDisposed();
            if (value is null)
            {
                WriteEmptyCell();
                return;
            }
            Write(value.Value);
        }

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The row has already been disposed.</exception>
        /// <remarks>
        /// Written as an Excel time serial: the fraction of a 24h day, in [0,1). This is a plain number
        /// cell (no date style); <c>ExcelParser&lt;T&gt;</c> reconstructs the <see cref="TimeOnly"/> from the fraction.
        /// </remarks>
        public void Write(TimeOnly value)
        {
            ThrowIfDisposed();
            CellFormatter.WriteNumber(_row, value.Ticks / (double)TimeSpan.TicksPerDay, _columnIndex, _rowNumber, ConsumeCellReference());
            _columnIndex++;
        }

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The row has already been disposed.</exception>
        public void Write(TimeOnly? value)
        {
            ThrowIfDisposed();
            if (value is null)
            {
                WriteEmptyCell();
                return;
            }
            Write(value.Value);
        }

        /// <summary>
        /// Writes <paramref name="value"/> to the current cell as a number and advances to the next column.
        /// </summary>
        /// <exception cref="ObjectDisposedException">The row has already been disposed.</exception>
        public void Write(int value)
        {
            ThrowIfDisposed();
            CellFormatter.WriteNumber(_row, value, _columnIndex, _rowNumber, ConsumeCellReference(), EffectiveStyle());
            _columnIndex++;
        }

        /// <summary>
        /// Writes <paramref name="value"/> to the current cell as a number, or an empty cell when it is <see langword="null"/>, and advances to the next column.
        /// </summary>
        /// <exception cref="ObjectDisposedException">The row has already been disposed.</exception>
        public void Write(int? value)
        {
            ThrowIfDisposed();
            if (value is null)
            {
                WriteEmptyCell();
                return;
            }
            Write(value.Value);
        }

        /// <summary>
        /// Writes <paramref name="value"/> to the current cell as a number and advances to the next column.
        /// </summary>
        /// <exception cref="ObjectDisposedException">The row has already been disposed.</exception>
        public void Write(long value)
        {
            ThrowIfDisposed();
            CellFormatter.WriteNumber(_row, value, _columnIndex, _rowNumber, ConsumeCellReference(), EffectiveStyle());
            _columnIndex++;
        }

        /// <summary>
        /// Writes <paramref name="value"/> to the current cell as a number, or an empty cell when it is <see langword="null"/>, and advances to the next column.
        /// </summary>
        /// <exception cref="ObjectDisposedException">The row has already been disposed.</exception>
        public void Write(long? value)
        {
            ThrowIfDisposed();
            if (value is null)
            {
                WriteEmptyCell();
                return;
            }
            Write(value.Value);
        }

        /// <summary>
        /// Writes <paramref name="value"/> to the current cell as a number and advances to the next column.
        /// </summary>
        /// <exception cref="ObjectDisposedException">The row has already been disposed.</exception>
        public void Write(double value)
        {
            ThrowIfDisposed();
            CellFormatter.WriteNumber(_row, value, _columnIndex, _rowNumber, ConsumeCellReference(), EffectiveStyle());
            _columnIndex++;
        }

        /// <summary>
        /// Writes <paramref name="value"/> to the current cell as a number, or an empty cell when it is <see langword="null"/>, and advances to the next column.
        /// </summary>
        /// <exception cref="ObjectDisposedException">The row has already been disposed.</exception>
        public void Write(double? value)
        {
            ThrowIfDisposed();
            if (value is null)
            {
                WriteEmptyCell();
                return;
            }
            Write(value.Value);
        }

        /// <summary>
        /// Writes <paramref name="value"/> to the current cell as a number and advances to the next column.
        /// </summary>
        /// <exception cref="ObjectDisposedException">The row has already been disposed.</exception>
        public void Write(decimal value)
        {
            ThrowIfDisposed();
            CellFormatter.WriteNumber(_row, value, _columnIndex, _rowNumber, ConsumeCellReference(), EffectiveStyle());
            _columnIndex++;
        }

        /// <summary>
        /// Writes <paramref name="value"/> to the current cell as a number, or an empty cell when it is <see langword="null"/>, and advances to the next column.
        /// </summary>
        /// <exception cref="ObjectDisposedException">The row has already been disposed.</exception>
        public void Write(decimal? value)
        {
            ThrowIfDisposed();
            if (value is null)
            {
                WriteEmptyCell();
                return;
            }
            Write(value.Value);
        }

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The row has already been disposed.</exception>
        public void Write<T>(T value)
            where T : IUtf8SpanFormattable
        {
            ThrowIfDisposed();
            CellFormatter.WriteNumber(_row, value, _columnIndex, _rowNumber, ConsumeCellReference(), EffectiveStyle());
            _columnIndex++;
        }

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The row has already been disposed.</exception>
        public void Write<T>(T? value)
            where T : struct, IUtf8SpanFormattable
        {
            ThrowIfDisposed();
            if (value is null)
            {
                WriteEmptyCell();
                return;
            }
            Write(value.Value);
        }

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The row has already been disposed.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="count"/> is negative.</exception>
        /// <exception cref="ExcelLimitExceededException">Skipping would advance past Excel's 16,384-column limit.</exception>
        public void Skip(int count = 1)
        {
            ThrowIfDisposed();
            ArgumentOutOfRangeException.ThrowIfNegative(count);
            if (count > 0)
            {
                if (_columnIndex > ExcelLimits.MaxColumns - count)
                {
                    throw new ExcelLimitExceededException("Columns", ExcelLimits.MaxColumns, (long)_columnIndex + count);
                }
                _useCellReferences = true;
            }
            _columnIndex += count;
        }

        // Only the first cell written after a real Skip() needs an explicit r="..." reference.
        private bool ConsumeCellReference()
        {
            ExcelLimits.ThrowIfColumnOutOfRange(_columnIndex);
            bool result = _useCellReferences;
            _useCellReferences = false;
            return result;
        }

        private void WriteEmptyCell()
        {
            CellFormatter.WriteEmpty(_row, _columnIndex, _rowNumber, includeReference: ConsumeCellReference(), styleId: EffectiveStyle());
            _columnIndex++;
        }

        /// <inheritdoc/>
        public ValueTask DisposeAsync()
        {
            if (_disposed)
            {
                return ValueTask.CompletedTask;
            }
            _disposed = true;
            return _owner.EndBufferedRowAsync();
        }

        /// <summary>
        /// Synchronous counterpart to <see cref="DisposeAsync"/>: ends the row and, on the rare buffer-threshold flush, writes to the destination stream synchronously.
        /// </summary>
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            _owner.EndBufferedRow();
        }
    }
}
