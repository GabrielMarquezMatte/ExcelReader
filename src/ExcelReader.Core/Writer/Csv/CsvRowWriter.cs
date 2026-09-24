using System.Buffers;
using System.Buffers.Text;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Unicode;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Csv
{
    /// <summary>
    /// Writes the fields of a single CSV row, quoting a field only when it contains the delimiter,
    /// quote character, carriage return, or line feed.
    /// </summary>
    public sealed class CsvRowWriter : IRowWriter
    {
        private const int StackFieldBytes = 64;

        private readonly CsvWriter _owner;
        private readonly BiffBuffer _buffer;
        private readonly byte _delimiter;
        private readonly byte _quote;
        private readonly SearchValues<byte> _specialBytes;
        private readonly SearchValues<char> _specialChars;
        private int _columnIndex;
        private bool _disposed;

        internal CsvRowWriter(CsvWriter owner, BiffBuffer buffer, byte delimiter, byte quote,
            SearchValues<byte> specialBytes, SearchValues<char> specialChars)
        {
            _owner = owner;
            _buffer = buffer;
            _delimiter = delimiter;
            _quote = quote;
            _specialBytes = specialBytes;
            _specialChars = specialChars;
        }

        internal void Reset()
        {
            _columnIndex = 0;
            _disposed = false;
        }

        private void ThrowIfDisposed()
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
        }

        private void BeginField()
        {
            if (_columnIndex > 0)
            {
                _buffer.WriteByte(_delimiter);
            }
            _columnIndex++;
        }

        /// <inheritdoc/>
        public void Write(string? value)
        {
            ThrowIfDisposed();
            BeginField();
            if (!string.IsNullOrEmpty(value))
            {
                WriteStringField(value);
            }
        }

        /// <inheritdoc/>
        public void WriteUtf8(ReadOnlySpan<byte> utf8)
        {
            if (!Utf8.IsValid(utf8))
            {
                Write(Encoding.UTF8.GetString(utf8));
                return;
            }
            ThrowIfDisposed();
            BeginField();
            if (!utf8.IsEmpty)
            {
                WriteFieldBytes(utf8);
            }
        }

        /// <inheritdoc/>
        /// <remarks>
        /// Writes lowercase "true"/"false" — the only spellings ColumnParserFactory's IsTruthy recognizes
        /// besides "1", so a written bool round-trips through ExcelParser&lt;T&gt; unchanged.
        /// </remarks>
        public void Write(bool value)
        {
            ThrowIfDisposed();
            BeginField();
            _buffer.Write(value ? "true"u8 : "false"u8);
        }

        /// <inheritdoc/>
        public void Write(bool? value)
        {
            if (value is null) { Skip(1); return; }
            Write(value.Value);
        }

        /// <inheritdoc/>
        /// <remarks>
        /// Writes the round-trip ISO 8601 ("O") format: unambiguous, culture-independent, and — for
        /// the common no-offset shape — exactly what ColumnParserFactory's CSV date fast path recognizes.
        /// </remarks>
        public void Write(DateTime value)
        {
            ThrowIfDisposed();
            BeginField();
            WriteUtf8Field(value, "O");
        }

        /// <inheritdoc/>
        public void Write(DateTime? value)
        {
            if (value is null) { Skip(1); return; }
            Write(value.Value);
        }

        /// <inheritdoc/>
        /// <remarks>
        /// Writes the value as ISO "yyyy-MM-dd" (the "O" round-trip form), which the CSV text-date
        /// parser reads back.
        /// </remarks>
        public void Write(DateOnly value)
        {
            ThrowIfDisposed();
            BeginField();
            WriteUtf8Field(value, "O");
        }

        /// <inheritdoc/>
        public void Write(DateOnly? value)
        {
            if (value is null) { Skip(1); return; }
            Write(value.Value);
        }

        /// <inheritdoc/>
        /// <remarks>
        /// Writes the value as an Excel-style time serial (fraction of a 24h day): a plain number the
        /// parser reconstructs via TryGetDouble, matching how the XLSX/XLSB row writers store it.
        /// </remarks>
        public void Write(TimeOnly value)
        {
            ThrowIfDisposed();
            BeginField();
            WriteUtf8Field(value.Ticks / (double)TimeSpan.TicksPerDay, default);
        }

        /// <inheritdoc/>
        public void Write(TimeOnly? value)
        {
            if (value is null) { Skip(1); return; }
            Write(value.Value);
        }

        /// <inheritdoc/>
        /// <remarks>
        /// Constrained to IUtf8SpanFormattable (not ISpanFormattable): CSV output is UTF-8, so numbers
        /// format straight to bytes with no char buffer or transcode. Every BCL numeric type implements it.
        /// </remarks>
        public void Write<T>(T value)
            where T : IUtf8SpanFormattable
        {
            ThrowIfDisposed();
            BeginField();
            WriteUtf8Field(value, default);
        }

        /// <inheritdoc/>
        public void Write<T>(T? value)
            where T : struct, IUtf8SpanFormattable
        {
            if (value is null) { Skip(1); return; }
            Write(value.Value);
        }

        /// <inheritdoc/>
        public void Skip(int count = 1)
        {
            ThrowIfDisposed();
            ArgumentOutOfRangeException.ThrowIfNegative(count);
            for (int i = 0; i < count; i++)
            {
                BeginField();
            }
        }

        [SkipLocalsInit]
        private void WriteUtf8Field<T>(T value, ReadOnlySpan<char> format) where T : IUtf8SpanFormattable
        {
            Span<byte> buf = stackalloc byte[StackFieldBytes];
            if (format.IsEmpty && typeof(T) == typeof(int) && Utf8Formatter.TryFormat(Unsafe.As<T, int>(ref value), buf, out var written))
            {
                WriteFieldBytes(buf[..written]);
                return;
            }
            if (format.IsEmpty && typeof(T) == typeof(long) && Utf8Formatter.TryFormat(Unsafe.As<T, long>(ref value), buf, out written))
            {
                WriteFieldBytes(buf[..written]);
                return;
            }
            if (format.IsEmpty && typeof(T) == typeof(double) && CellFormatter.TryFormatDouble(Unsafe.As<T, double>(ref value), buf, out written))
            {
                WriteFieldBytes(buf[..written]);
                return;
            }
            if (!value.TryFormat(buf, out written, format, CultureInfo.InvariantCulture))
            {
                WriteUtf8FieldSlow(value, format);
                return;
            }
            WriteFieldBytes(buf[..written]);
        }

        private void WriteUtf8FieldSlow<T>(T value, ReadOnlySpan<char> format)
            where T : IUtf8SpanFormattable
        {
            byte[] rented = ArrayPool<byte>.Shared.Rent(StackFieldBytes * 4);
            int written;
            while (!value.TryFormat(rented, out written, format, CultureInfo.InvariantCulture))
            {
                int larger = rented.Length * 2;
                ArrayPool<byte>.Shared.Return(rented);
                rented = ArrayPool<byte>.Shared.Rent(larger);
            }
            WriteFieldBytes(rented.AsSpan(0, written));
            ArrayPool<byte>.Shared.Return(rented);
        }

        private void WriteFieldBytes(ReadOnlySpan<byte> value)
        {
            int firstSpecial = value.IndexOfAny(_specialBytes);
            if (firstSpecial < 0)
            {
                _buffer.Write(value);
                return;
            }
            _buffer.WriteByte(_quote);
            int start = 0;
            for (int i = firstSpecial; i < value.Length; i++)
            {
                if (value[i] != _quote)
                {
                    continue;
                }
                _buffer.Write(value[start..i]);
                _buffer.WriteByte(_quote);
                _buffer.WriteByte(_quote);
                start = i + 1;
            }
            _buffer.Write(value[start..]);
            _buffer.WriteByte(_quote);
        }

        private void WriteStringField(ReadOnlySpan<char> value)
        {
            int firstSpecial = value.IndexOfAny(_specialChars);
            if (firstSpecial < 0)
            {
                _buffer.WriteUtf8(value);
                return;
            }
            char quote = (char)_quote;
            _buffer.WriteByte(_quote);
            int start = 0;
            for (int i = firstSpecial; i < value.Length; i++)
            {
                if (value[i] != quote)
                {
                    continue;
                }
                _buffer.WriteUtf8(value[start..i]);
                _buffer.WriteByte(_quote);
                _buffer.WriteByte(_quote);
                start = i + 1;
            }
            _buffer.WriteUtf8(value[start..]);
            _buffer.WriteByte(_quote);
        }

        /// <inheritdoc/>
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            _owner.EndRow();
        }

        /// <inheritdoc/>
        public ValueTask DisposeAsync()
        {
            if (_disposed)
            {
                return ValueTask.CompletedTask;
            }
            _disposed = true;
            return _owner.EndRowAsync();
        }
    }
}
