using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Internal;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Xlsx
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
        private int _pendingStyle = -1;
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
            _pendingStyle = -1;
            _useCellReferences = false;
            _disposed = false;
        }

        private int EffectiveStyle()
        {
            if (_pendingStyle >= 0)
            {
                int pending = _pendingStyle;
                _pendingStyle = -1;
                return pending;
            }
            return _rowStyleId != 0 ? _rowStyleId : _owner.GetColumnStyle(_columnIndex);
        }

        /// <summary>Styles the next cell written, and only that cell, replacing the row and column style. Style 0 forces the general style.</summary>
        /// <param name="styleId">The style to apply, from <see cref="XlsxWorkbookWriter.AddStyle"/>.</param>
        /// <returns>This row writer.</returns>
        /// <exception cref="ObjectDisposedException">The row has already been disposed.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="styleId"/> is negative or was never returned by <see cref="XlsxWorkbookWriter.AddStyle"/>.</exception>
        public XlsxRowWriter Style(int styleId)
        {
            ThrowIfDisposed();
            ArgumentOutOfRangeException.ThrowIfNegative(styleId);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(styleId, _owner.StyleCount);
            _pendingStyle = styleId;
            return this;
        }

        IRowWriter IRowWriter.Style(int styleId)
        {
            return Style(styleId);
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
            WriteText(value, value);
        }

        /// <inheritdoc/>
        /// <exception cref="ObjectDisposedException">The row has already been disposed.</exception>
        public void WriteUtf8(ReadOnlySpan<byte> utf8)
        {
            ThrowIfDisposed();
            if (!_owner.UseSharedStrings && CellFormatter.IsPlainUtf8Text(utf8))
            {
                CellFormatter.WritePlainUtf8String(_row, utf8, _columnIndex, _rowNumber, ConsumeCellReference(), EffectiveStyle());
                _columnIndex++;
                return;
            }
            Utf8Chars.Decode(utf8, this, static (chars, self) => self.WriteText(chars, owned: null));
        }

        private void WriteText(ReadOnlySpan<char> value, string? owned)
        {
            ExcelLimits.ThrowIfCellTextTooLong(value.Length, nameof(value));
            if (_owner.UseSharedStrings)
            {
                int index = owned is null ? _owner.GetSharedStringIndex(value) : _owner.GetSharedStringIndex(owned);
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
            CellFormatter.WriteDateTime(_row, value, _owner.Date1904, _columnIndex, _rowNumber, ConsumeCellReference(), styleId == 0 ? 1 : _owner.DateStyle(styleId));
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
            CellFormatter.WriteDateTime(_row, value.ToDateTime(TimeOnly.MinValue), _owner.Date1904, _columnIndex, _rowNumber, ConsumeCellReference(), styleId == 0 ? 1 : _owner.DateStyle(styleId));
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
            CellFormatter.WriteNumber(_row, value.Ticks / (double)TimeSpan.TicksPerDay, _columnIndex, _rowNumber, ConsumeCellReference(), EffectiveStyle());
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
                if (_pendingStyle >= 0)
                {
                    WriterStateGuard.ThrowPendingStyle("skipping");
                }
                if (_columnIndex > ExcelLimits.MaxColumns - count)
                {
                    throw new ExcelLimitExceededException("Columns", ExcelLimits.MaxColumns, (long)_columnIndex + count);
                }
                _useCellReferences = true;
            }
            _columnIndex += count;
        }

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
            ValueTask ended = _owner.EndBufferedRowAsync();
            if (_pendingStyle < 0)
            {
                return ended;
            }
            _pendingStyle = -1;
            return ThrowPendingStyleAsync(ended);
        }

        private static async ValueTask ThrowPendingStyleAsync(ValueTask ended)
        {
            await ended.ConfigureAwait(false);
            WriterStateGuard.ThrowPendingStyle("disposing the row");
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
            if (_pendingStyle < 0)
            {
                return;
            }
            _pendingStyle = -1;
            WriterStateGuard.ThrowPendingStyle("disposing the row");
        }
    }
}
