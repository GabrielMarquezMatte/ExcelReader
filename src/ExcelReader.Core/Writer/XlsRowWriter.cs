namespace ExcelReader.Core.Writer
{
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

        public void Write(string? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                _owner.EmitLabel(_rowNumber, _columnIndex, value);
            }
            _columnIndex++;
        }

        public void Write(bool value)
        {
            ThrowIfDisposed();
            _owner.EmitBool(_rowNumber, _columnIndex, value);
            _columnIndex++;
        }

        public void Write(bool? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                _owner.EmitBool(_rowNumber, _columnIndex, value.Value);
            }
            _columnIndex++;
        }

        public void Write(DateTime value)
        {
            ThrowIfDisposed();
            _owner.EmitDate(_rowNumber, _columnIndex, value);
            _columnIndex++;
        }

        public void Write(DateTime? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                _owner.EmitDate(_rowNumber, _columnIndex, value.Value);
            }
            _columnIndex++;
        }

        // DateOnly shares the DateTime date-serial cell format (midnight), so it round-trips as a date.
        public void Write(DateOnly value)
        {
            ThrowIfDisposed();
            _owner.EmitDate(_rowNumber, _columnIndex, value.ToDateTime(TimeOnly.MinValue));
            _columnIndex++;
        }

        public void Write(DateOnly? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                _owner.EmitDate(_rowNumber, _columnIndex, value.Value.ToDateTime(TimeOnly.MinValue));
            }
            _columnIndex++;
        }

        // TimeOnly is written as an Excel time serial (fraction of a 24h day) in a plain number cell.
        public void Write(TimeOnly value)
        {
            ThrowIfDisposed();
            _owner.EmitNumber(_rowNumber, _columnIndex, value.Ticks / (double)TimeSpan.TicksPerDay);
            _columnIndex++;
        }

        public void Write(TimeOnly? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                _owner.EmitNumber(_rowNumber, _columnIndex, value.Value.Ticks / (double)TimeSpan.TicksPerDay);
            }
            _columnIndex++;
        }

        public void Write(int value)
        {
            ThrowIfDisposed();
            _owner.EmitNumber(_rowNumber, _columnIndex, value);
            _columnIndex++;
        }

        public void Write(int? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                _owner.EmitNumber(_rowNumber, _columnIndex, value.Value);
            }
            _columnIndex++;
        }

        public void Write(long value)
        {
            ThrowIfDisposed();
            _owner.EmitNumber(_rowNumber, _columnIndex, value);
            _columnIndex++;
        }

        public void Write(long? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                _owner.EmitNumber(_rowNumber, _columnIndex, value.Value);
            }
            _columnIndex++;
        }

        public void Write(double value)
        {
            ThrowIfDisposed();
            _owner.EmitNumber(_rowNumber, _columnIndex, value);
            _columnIndex++;
        }

        public void Write(double? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                _owner.EmitNumber(_rowNumber, _columnIndex, value.Value);
            }
            _columnIndex++;
        }

        public void Write(decimal value)
        {
            ThrowIfDisposed();
            _owner.EmitNumber(_rowNumber, _columnIndex, (double)value);
            _columnIndex++;
        }

        public void Write(decimal? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                _owner.EmitNumber(_rowNumber, _columnIndex, (double)value.Value);
            }
            _columnIndex++;
        }

        public void Write<T>(T value)
            where T : IUtf8SpanFormattable
        {
            ThrowIfDisposed();
            _owner.EmitNumber(_rowNumber, _columnIndex, XlsbRowWriter.ToDouble(value));
            _columnIndex++;
        }

        public void Write<T>(T? value)
            where T : struct, IUtf8SpanFormattable
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                _owner.EmitNumber(_rowNumber, _columnIndex, XlsbRowWriter.ToDouble(value.Value));
            }
            _columnIndex++;
        }

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

        public void Dispose()
        {
            if (!_disposed)
            {
                _disposed = true;
                _owner.NotifyRowEnded();
            }
        }

        // XLS buffers everything in memory, so there is no real async work — wrap the sync path.
        public ValueTask DisposeAsync()
        {
            Dispose();
            return ValueTask.CompletedTask;
        }
    }
}
