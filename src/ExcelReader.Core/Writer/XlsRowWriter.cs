using System.Diagnostics.CodeAnalysis;
using System.Globalization;

namespace ExcelReader.Core.Writer
{
    public sealed class XlsRowWriter : IDisposable, IRowWriter
    {
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "XlsSheetWriter is borrowed; its lifetime is managed by the caller.")]
        private readonly XlsSheetWriter _owner;
        private readonly int _rowNumber;
        private int _columnIndex;
        private bool _disposed;

        internal XlsRowWriter(XlsSheetWriter owner, int rowNumber)
        {
            _owner = owner;
            _rowNumber = rowNumber;
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

        public void Write(float value)
        {
            ThrowIfDisposed();
            _owner.EmitNumber(_rowNumber, _columnIndex, value);
            _columnIndex++;
        }

        public void Write(float? value)
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
            where T : ISpanFormattable
        {
            ThrowIfDisposed();
            _owner.EmitNumber(_rowNumber, _columnIndex, ToDouble(value));
            _columnIndex++;
        }

        public void Write<T>(T? value)
            where T : struct, ISpanFormattable
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                _owner.EmitNumber(_rowNumber, _columnIndex, ToDouble(value.Value));
            }
            _columnIndex++;
        }

        public void Skip(int count = 1)
        {
            ThrowIfDisposed();
            _columnIndex += count;
        }

        // XLS Number cells store a binary double; round-tripping through the invariant text form
        // gives the exact value for the common integer/floating types.
        private static double ToDouble<T>(T value)
            where T : ISpanFormattable
        {
            switch (value)
            {
                case double d:
                    return d;
                case float f:
                    return f;
                case decimal m:
                    return (double)m;
                case int i:
                    return i;
                case long l:
                    return l;
                case short s:
                    return s;
                case byte b:
                    return b;
                case uint ui:
                    return ui;
                case ulong ul:
                    return ul;
                case ushort us:
                    return us;
                case sbyte sb:
                    return sb;
            }
            Span<char> buffer = stackalloc char[64];
            if (value.TryFormat(buffer, out int written, default, CultureInfo.InvariantCulture)
                && double.TryParse(buffer[..written], CultureInfo.InvariantCulture, out double result))
            {
                return result;
            }
            return Convert.ToDouble(value, CultureInfo.InvariantCulture);
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
