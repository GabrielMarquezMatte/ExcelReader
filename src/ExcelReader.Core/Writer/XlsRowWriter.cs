namespace ExcelReader.Core.Writer
{
    /// <summary>
    /// Writes the cells of a single row of a BIFF8 (.xls) sheet.
    /// </summary>
    public sealed class XlsRowWriter : IDisposable, IRowWriter
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

        // Reused across rows by XlsSheetWriter: rents one instance per sheet instead of one per row.
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
            }
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
                Skip(1);
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
                Skip(1);
                return;
            }
            Write(value.Value);
        }

        // DateOnly shares the DateTime date-serial cell format (midnight), so it round-trips as a date.
        /// <inheritdoc/>
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
                Skip(1);
                return;
            }
            Write(value.Value);
        }

        // TimeOnly is written as an Excel time serial (fraction of a 24h day) in a plain number cell.
        /// <inheritdoc/>
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
                Skip(1);
                return;
            }
            Write(value.Value);
        }

        /// <summary>Writes a numeric cell in the current column and advances to the next column.</summary>
        public void Write(int value)
        {
            ThrowIfDisposed();
            _owner.EmitNumber(_rowNumber, _columnIndex, value);
            _columnIndex++;
        }

        /// <summary>Writes a numeric cell, or leaves the column blank if <paramref name="value"/> is <see langword="null"/>, then advances to the next column.</summary>
        public void Write(int? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                _owner.EmitNumber(_rowNumber, _columnIndex, value.Value);
            }
            _columnIndex++;
        }

        /// <summary>Writes a numeric cell in the current column and advances to the next column.</summary>
        public void Write(long value)
        {
            ThrowIfDisposed();
            _owner.EmitNumber(_rowNumber, _columnIndex, value);
            _columnIndex++;
        }

        /// <summary>Writes a numeric cell, or leaves the column blank if <paramref name="value"/> is <see langword="null"/>, then advances to the next column.</summary>
        public void Write(long? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                _owner.EmitNumber(_rowNumber, _columnIndex, value.Value);
            }
            _columnIndex++;
        }

        /// <summary>Writes a numeric cell in the current column and advances to the next column.</summary>
        public void Write(double value)
        {
            ThrowIfDisposed();
            _owner.EmitNumber(_rowNumber, _columnIndex, value);
            _columnIndex++;
        }

        /// <summary>Writes a numeric cell, or leaves the column blank if <paramref name="value"/> is <see langword="null"/>, then advances to the next column.</summary>
        public void Write(double? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                _owner.EmitNumber(_rowNumber, _columnIndex, value.Value);
            }
            _columnIndex++;
        }

        /// <summary>Writes a numeric cell in the current column and advances to the next column.</summary>
        public void Write(decimal value)
        {
            ThrowIfDisposed();
            _owner.EmitNumber(_rowNumber, _columnIndex, (double)value);
            _columnIndex++;
        }

        /// <summary>Writes a numeric cell, or leaves the column blank if <paramref name="value"/> is <see langword="null"/>, then advances to the next column.</summary>
        public void Write(decimal? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                _owner.EmitNumber(_rowNumber, _columnIndex, (double)value.Value);
            }
            _columnIndex++;
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
                Skip(1);
                return;
            }
            Write(value.Value);
        }

        /// <inheritdoc/>
        public void Skip(int count = 1)
        {
            ThrowIfDisposed();
            ArgumentOutOfRangeException.ThrowIfNegative(count);
            _columnIndex += count;
        }

        private void ThrowIfDisposed()
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
        }

        /// <inheritdoc/>
        public void Dispose()
        {
            if (!_disposed)
            {
                _disposed = true;
                _owner.NotifyRowEnded();
            }
        }

        // XLS buffers everything in memory, so there is no real async work — wrap the sync path.
        /// <inheritdoc/>
        public ValueTask DisposeAsync()
        {
            Dispose();
            return ValueTask.CompletedTask;
        }
    }
}
