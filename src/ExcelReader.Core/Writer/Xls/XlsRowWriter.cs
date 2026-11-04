using ExcelReader.Core.Writer.Internal;
using ExcelReader.Core.Writer.Xlsb;

namespace ExcelReader.Core.Writer.Xls
{
    /// <summary>
    /// Writes the cells of a single row of a BIFF8 (.xls) sheet.
    /// </summary>
    public sealed class XlsRowWriter : IRowWriter
    {
        private readonly XlsSheetWriter _owner;
        private int _rowNumber;
        private int _columnIndex;
        private bool _disposed;

        internal XlsRowWriter(XlsSheetWriter owner, int rowNumber)
        {
            _owner = owner;
            _rowNumber = rowNumber;
        }

        internal void Reset(int rowNumber)
        {
            _rowNumber = rowNumber;
            _columnIndex = 0;
            _disposed = false;
        }

        /// <inheritdoc/>
        public void Write(string? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                _owner.EmitLabel(_rowNumber, _columnIndex, value);
                _columnIndex++;
                return;
            }
            WriteNull();
        }

        /// <summary>Styles the next cell written, and only that cell, replacing the row and column style. Style 0 forces the general style.</summary>
        /// <param name="styleId">The style to apply, from <see cref="XlsWorkbookWriter.AddStyle"/>.</param>
        /// <returns>This row writer.</returns>
        /// <exception cref="ObjectDisposedException">The row has already been disposed.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="styleId"/> is negative or was never returned by <see cref="XlsWorkbookWriter.AddStyle"/>.</exception>
        public XlsRowWriter Style(int styleId)
        {
            ThrowIfDisposed();
            _owner.SetPendingStyle(styleId);
            return this;
        }

        IRowWriter IRowWriter.Style(int styleId)
        {
            return Style(styleId);
        }

        private void WriteNull()
        {
            if (!_owner.HasPendingStyle)
            {
                Skip(1);
                return;
            }
            ThrowIfDisposed();
            _owner.EmitBlank(_rowNumber, _columnIndex);
            _columnIndex++;
        }

        /// <inheritdoc/>
        public void WriteUtf8(ReadOnlySpan<byte> utf8)
        {
            ThrowIfDisposed();
            Utf8Chars.Decode(utf8, this, static (chars, self) => self._owner.EmitLabel(self._rowNumber, self._columnIndex, chars));
            _columnIndex++;
        }

        /// <inheritdoc/>
        public void Write(bool value)
        {
            ThrowIfDisposed();
            _owner.EmitBool(_rowNumber, _columnIndex, value);
            _columnIndex++;
        }

        /// <inheritdoc/>
        public void Write(bool? value)
        {
            if (value is null)
            {
                WriteNull();
                return;
            }
            Write(value.Value);
        }

        /// <inheritdoc/>
        public void Write(DateTime value)
        {
            ThrowIfDisposed();
            _owner.EmitDate(_rowNumber, _columnIndex, value);
            _columnIndex++;
        }

        /// <inheritdoc/>
        public void Write(DateTime? value)
        {
            if (value is null)
            {
                WriteNull();
                return;
            }
            Write(value.Value);
        }

        /// <inheritdoc/>
        /// <remarks>
        /// Shares the <see cref="DateTime"/> date-serial cell format (midnight), so it round-trips as a date.
        /// </remarks>
        public void Write(DateOnly value)
        {
            ThrowIfDisposed();
            _owner.EmitDate(_rowNumber, _columnIndex, value.ToDateTime(TimeOnly.MinValue));
            _columnIndex++;
        }

        /// <inheritdoc/>
        public void Write(DateOnly? value)
        {
            if (value is null)
            {
                WriteNull();
                return;
            }
            Write(value.Value);
        }

        /// <inheritdoc/>
        /// <remarks>Written as an Excel time serial (fraction of a 24h day) in a plain number cell.</remarks>
        public void Write(TimeOnly value)
        {
            ThrowIfDisposed();
            _owner.EmitNumber(_rowNumber, _columnIndex, value.Ticks / (double)TimeSpan.TicksPerDay);
            _columnIndex++;
        }

        /// <inheritdoc/>
        public void Write(TimeOnly? value)
        {
            if (value is null)
            {
                WriteNull();
                return;
            }
            Write(value.Value);
        }

        /// <inheritdoc/>
        public void Write<T>(T value)
            where T : IUtf8SpanFormattable
        {
            ThrowIfDisposed();
            _owner.EmitNumber(_rowNumber, _columnIndex, XlsbRowWriter.ToDouble(value));
            _columnIndex++;
        }

        /// <inheritdoc/>
        public void Write<T>(T? value)
            where T : struct, IUtf8SpanFormattable
        {
            if (value is null)
            {
                WriteNull();
                return;
            }
            Write(value.Value);
        }

        /// <inheritdoc/>
        /// <exception cref="InvalidOperationException">Skipping would advance past BIFF8's 256-column limit.</exception>
        public void Skip(int count = 1)
        {
            ThrowIfDisposed();
            ArgumentOutOfRangeException.ThrowIfNegative(count);
            if (count > 0 && _owner.HasPendingStyle)
            {
                WriterStateGuard.ThrowPendingStyle("skipping");
            }
            if (count > 0 && _columnIndex > XlsSheetWriter.MaxColumn - count + 1)
            {
                throw new InvalidOperationException($"BIFF8 worksheets are limited to {XlsSheetWriter.MaxColumn + 1} columns.");
            }
            _columnIndex += count;
        }

        private void ThrowIfDisposed()
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
        }

        /// <inheritdoc/>
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            bool wasPending = _owner.ClearPendingStyle();
            _owner.NotifyRowEnded();
            if (wasPending)
            {
                WriterStateGuard.ThrowPendingStyle("disposing the row");
            }
        }

        /// <inheritdoc/>
        /// <remarks>XLS buffers everything in memory, so there is no real async work — this wraps the sync path.</remarks>
        public ValueTask DisposeAsync()
        {
            Dispose();
            return ValueTask.CompletedTask;
        }
    }
}
