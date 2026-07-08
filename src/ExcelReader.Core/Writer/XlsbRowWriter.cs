using System.Diagnostics.CodeAnalysis;
using System.Globalization;

namespace ExcelReader.Core.Writer
{
    public sealed class XlsbRowWriter : IRowWriter
    {
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "XlsbSheetWriter is borrowed; its lifetime is managed by the caller.")]
        private readonly XlsbSheetWriter _owner;
        private int _columnIndex;
        private bool _disposed;

        internal XlsbRowWriter(XlsbSheetWriter owner)
        {
            _owner = owner;
        }

        private void ThrowIfDisposed()
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
        }

        public void Write(string? value)
        {
            ThrowIfDisposed();
            _owner.WriteStringCell(_columnIndex, value);
            _columnIndex++;
        }

        public void Write(bool value)
        {
            ThrowIfDisposed();
            _owner.WriteBoolCell(_columnIndex, value);
            _columnIndex++;
        }

        public void Write(bool? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                Write(value.Value);
                return;
            }
            _columnIndex++;
        }

        public void Write(DateTime value)
        {
            ThrowIfDisposed();
            _owner.WriteDateSerialCell(_columnIndex, value.ToOADate());
            _columnIndex++;
        }

        public void Write(DateTime? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                Write(value.Value);
                return;
            }
            _columnIndex++;
        }

        // DateOnly shares the DateTime date-serial cell format (midnight), so it round-trips as a date.
        public void Write(DateOnly value)
        {
            ThrowIfDisposed();
            Write(value.ToDateTime(TimeOnly.MinValue));
        }

        public void Write(DateOnly? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                Write(value.Value.ToDateTime(TimeOnly.MinValue));
                return;
            }
            _columnIndex++;
        }

        // TimeOnly is written as an Excel time serial (fraction of a 24h day) in a plain number cell.
        public void Write(TimeOnly value)
        {
            ThrowIfDisposed();
            WriteDouble(value.Ticks / (double)TimeSpan.TicksPerDay, style: 0);
        }

        public void Write(TimeOnly? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                WriteDouble(value.Value.Ticks / (double)TimeSpan.TicksPerDay, style: 0);
                return;
            }
            _columnIndex++;
        }

        public void Write(int value)
        {
            ThrowIfDisposed();
            WriteDouble(value, style: 0);
        }

        public void Write(int? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                WriteDouble(value.Value, style: 0);
                return;
            }
            _columnIndex++;
        }

        public void Write(long value)
        {
            ThrowIfDisposed();
            WriteDouble(value, style: 0);
        }

        public void Write(long? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                WriteDouble(value.Value, style: 0);
                return;
            }
            _columnIndex++;
        }

        public void Write(double value)
        {
            ThrowIfDisposed();
            WriteDouble(value, style: 0);
        }

        public void Write(double? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                WriteDouble(value.Value, style: 0);
                return;
            }
            _columnIndex++;
        }

        public void Write(decimal value)
        {
            ThrowIfDisposed();
            WriteDouble((double)value, style: 0);
        }

        public void Write(decimal? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                WriteDouble((double)value.Value, style: 0);
                return;
            }
            _columnIndex++;
        }

        public void Write<T>(T value)
            where T : ISpanFormattable
        {
            ThrowIfDisposed();
            WriteDouble(ToDouble(value), style: 0);
        }

        public void Write<T>(T? value)
            where T : struct, ISpanFormattable
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                Write(value.Value);
                return;
            }
            _columnIndex++;
        }

        public void Skip(int count = 1)
        {
            ThrowIfDisposed();
            ArgumentOutOfRangeException.ThrowIfNegative(count);
            _columnIndex += count;
        }

        private void WriteDouble(double value, int style)
        {
            _owner.WriteDoubleCell(_columnIndex, value, style);
            _columnIndex++;
        }

        internal static double ToDouble<T>(T value)
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
            if (value is IConvertible convertible)
            {
                return convertible.ToDouble(CultureInfo.InvariantCulture);
            }
            Span<char> chars = stackalloc char[64];
            return value.TryFormat(chars, out int written, default, CultureInfo.InvariantCulture) &&
                double.TryParse(chars[..written], CultureInfo.InvariantCulture, out double parsed)
                ? parsed
                : 0.0;
        }

        public ValueTask DisposeAsync()
        {
            if (_disposed)
            {
                return ValueTask.CompletedTask;
            }
            _disposed = true;
            _owner.NotifyRowEnded();
            return ValueTask.CompletedTask;
        }
    }
}
