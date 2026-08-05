using System.Globalization;
using System.Runtime.CompilerServices;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Writer
{
    /// <summary>Writes the cells of a single row into an .xlsb worksheet, one column at a time.</summary>
    public sealed class XlsbRowWriter : IRowWriter
    {
        private readonly XlsbSheetWriter _owner;
        private int _columnIndex;
        private bool _disposed;

        internal XlsbRowWriter(XlsbSheetWriter owner)
        {
            _owner = owner;
        }

        // Reused across rows by XlsbSheetWriter: rents one instance per sheet instead of one per row.
        internal void Reset()
        {
            _columnIndex = 0;
            _disposed = false;
        }

        private void ThrowIfDisposed()
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
        }

        /// <inheritdoc/>
        public void Write(string? value)
        {
            ThrowIfDisposed();
            _owner.WriteStringCell(_columnIndex, value);
            _columnIndex++;
        }

        /// <inheritdoc/>
        public void Write(bool value)
        {
            ThrowIfDisposed();
            _owner.WriteBoolCell(_columnIndex, value);
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
            _owner.WriteDateSerialCell(_columnIndex, value.ToOADate());
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

        /// <inheritdoc/>
        /// <remarks>
        /// Shares the <see cref="DateTime"/> date-serial cell format (midnight), so it round-trips as a date.
        /// </remarks>
        public void Write(DateOnly value)
        {
            ThrowIfDisposed();
            Write(value.ToDateTime(TimeOnly.MinValue));
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

        /// <inheritdoc/>
        /// <remarks>Written as an Excel time serial (fraction of a 24h day) in a plain number cell.</remarks>
        public void Write(TimeOnly value)
        {
            ThrowIfDisposed();
            WriteDouble(value.Ticks / (double)TimeSpan.TicksPerDay);
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

        /// <summary>Writes <paramref name="value"/> as a numeric cell in the current column, advancing to the next column.</summary>
        public void Write(int value)
        {
            ThrowIfDisposed();
            WriteDouble(value);
        }

        /// <summary>Writes <paramref name="value"/> as a numeric cell, or an empty cell if it has no value, advancing to the next column.</summary>
        public void Write(int? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                WriteDouble(value.Value);
                return;
            }
            _columnIndex++;
        }

        /// <summary>Writes <paramref name="value"/> as a numeric cell in the current column, advancing to the next column.</summary>
        public void Write(long value)
        {
            ThrowIfDisposed();
            WriteDouble(value);
        }

        /// <summary>Writes <paramref name="value"/> as a numeric cell, or an empty cell if it has no value, advancing to the next column.</summary>
        public void Write(long? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                WriteDouble(value.Value);
                return;
            }
            _columnIndex++;
        }

        /// <summary>Writes <paramref name="value"/> as a numeric cell in the current column, advancing to the next column.</summary>
        public void Write(double value)
        {
            ThrowIfDisposed();
            WriteDouble(value);
        }

        /// <summary>Writes <paramref name="value"/> as a numeric cell, or an empty cell if it has no value, advancing to the next column.</summary>
        public void Write(double? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                WriteDouble(value.Value);
                return;
            }
            _columnIndex++;
        }

        /// <summary>Writes <paramref name="value"/> as a numeric cell in the current column, advancing to the next column.</summary>
        public void Write(decimal value)
        {
            ThrowIfDisposed();
            WriteDouble((double)value);
        }

        /// <summary>Writes <paramref name="value"/> as a numeric cell, or an empty cell if it has no value, advancing to the next column.</summary>
        public void Write(decimal? value)
        {
            ThrowIfDisposed();
            if (value is not null)
            {
                WriteDouble((double)value.Value);
                return;
            }
            _columnIndex++;
        }

        /// <inheritdoc/>
        public void Write<T>(T value)
            where T : IUtf8SpanFormattable
        {
            ThrowIfDisposed();
            WriteDouble(ToDouble(value));
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
        /// <exception cref="ExcelLimitExceededException">Skipping would advance past Excel's 16,384-column limit.</exception>
        public void Skip(int count = 1)
        {
            ThrowIfDisposed();
            ArgumentOutOfRangeException.ThrowIfNegative(count);
            // XlsxRowWriter already bounds this; XLS/XLSB didn't, so Skip(int.MaxValue) advanced
            // _columnIndex unboundedly instead of failing fast.
            if (count > 0 && _columnIndex > ExcelLimits.MaxColumns - count)
            {
                throw new ExcelLimitExceededException("Columns", ExcelLimits.MaxColumns, (long)_columnIndex + count);
            }
            _columnIndex += count;
        }

        private void WriteDouble(double value)
        {
            _owner.WriteDoubleCell(_columnIndex, value);
            _columnIndex++;
        }

        // typeof(T) == typeof(...) folds to a JIT constant per generic instantiation, so the dead
        // branches are elided and — unlike a type-pattern switch on an unconstrained T — no boxing
        // occurs. Mirrors the pattern in CellFormatter.WriteValue<T>/CsvRowWriter.WriteUtf8Field.
        [SkipLocalsInit]
        internal static double ToDouble<T>(T value) where T : IUtf8SpanFormattable
        {
            if (typeof(T) == typeof(double))
            {
                return Unsafe.As<T, double>(ref value);
            }
            if (typeof(T) == typeof(float))
            {
                return Unsafe.As<T, float>(ref value);
            }
            if (typeof(T) == typeof(decimal))
            {
                return (double)Unsafe.As<T, decimal>(ref value);
            }
            if (typeof(T) == typeof(int))
            {
                return Unsafe.As<T, int>(ref value);
            }
            if (typeof(T) == typeof(long))
            {
                return Unsafe.As<T, long>(ref value);
            }
            if (typeof(T) == typeof(short))
            {
                return Unsafe.As<T, short>(ref value);
            }
            if (typeof(T) == typeof(byte))
            {
                return Unsafe.As<T, byte>(ref value);
            }
            if (typeof(T) == typeof(uint))
            {
                return Unsafe.As<T, uint>(ref value);
            }
            if (typeof(T) == typeof(ulong))
            {
                return Unsafe.As<T, ulong>(ref value);
            }
            if (typeof(T) == typeof(ushort))
            {
                return Unsafe.As<T, ushort>(ref value);
            }
            if (typeof(T) == typeof(sbyte))
            {
                return Unsafe.As<T, sbyte>(ref value);
            }
            if (value is IConvertible convertible)
            {
                return convertible.ToDouble(CultureInfo.InvariantCulture);
            }
            Span<byte> bytes = stackalloc byte[64];
            // Silently falling back to 0.0 here (as this used to) writes a wrong-but-plausible
            // value for a caller-supplied T that can't actually convert — banned by STYLEGUIDE.md:209.
            // This is a caller mistake (a Write<T> call with a T that doesn't represent a number), not a
            // file defect, so it throws ArgumentException, matching CellValueGuards' convention for the
            // same "caller passed something a spreadsheet cell can't hold" class of error.
            if (value.TryFormat(bytes, out int written, default, CultureInfo.InvariantCulture) &&
                double.TryParse(bytes[..written], CultureInfo.InvariantCulture, out double parsed))
            {
                return parsed;
            }
            throw new ArgumentException(
                $"Cannot convert value of type '{typeof(T)}' to a numeric cell value.", nameof(value));
        }

        /// <inheritdoc/>
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
