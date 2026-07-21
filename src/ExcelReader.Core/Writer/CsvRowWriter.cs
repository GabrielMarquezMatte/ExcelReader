using System.Buffers;
using System.Buffers.Text;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
{
    public sealed class CsvRowWriter : IRowWriter, IDisposable
    {
        // Numbers, dates, Guids, and every other BCL formattable fit an ASCII field well under this;
        // the rare overflow falls back to a rented buffer in WriteUtf8FieldSlow.
        private const int StackFieldBytes = 64;

        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "CsvWriter is borrowed; its lifetime is managed by the caller.")]
        private readonly CsvWriter _owner;
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "BiffBuffer is owned by CsvWriter; row writer borrows it.")]
        private readonly BiffBuffer _buffer;
        private readonly byte _delimiter;
        private readonly byte _quote;
        // Prebuilt (once per CsvWriter) so the per-field quote check is a vectorized scan rather than
        // a byte-at-a-time loop. Bytes for UTF-8 field output, chars for the string transcode path.
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

        // Reused across rows by CsvWriter: rents one instance per writer instead of one per row.
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

        public void Write(string? value)
        {
            ThrowIfDisposed();
            BeginField();
            if (!string.IsNullOrEmpty(value))
            {
                WriteStringField(value);
            }
        }

        // Lowercase "true"/"false" — the only spellings ColumnParserFactory's IsTruthy recognizes
        // besides "1", so a written bool round-trips through ExcelParser<T> unchanged.
        public void Write(bool value)
        {
            ThrowIfDisposed();
            BeginField();
            _buffer.Write(value ? "true"u8 : "false"u8);
        }

        public void Write(bool? value)
        {
            ThrowIfDisposed();
            BeginField();
            if (value is not null)
            {
                _buffer.Write(value.Value ? "true"u8 : "false"u8);
            }
        }

        // Round-trip ISO 8601 ("O"): unambiguous, culture-independent, and — for the common
        // no-offset shape — exactly what ColumnParserFactory's CSV date fast path recognizes.
        public void Write(DateTime value)
        {
            ThrowIfDisposed();
            BeginField();
            WriteUtf8Field(value, "O");
        }

        public void Write(DateTime? value)
        {
            ThrowIfDisposed();
            BeginField();
            if (value is not null)
            {
                WriteUtf8Field(value.Value, "O");
            }
        }

        // DateOnly as ISO "yyyy-MM-dd" (the "O" round-trip form), which the CSV text-date parser reads back.
        public void Write(DateOnly value)
        {
            ThrowIfDisposed();
            BeginField();
            WriteUtf8Field(value, "O");
        }

        public void Write(DateOnly? value)
        {
            ThrowIfDisposed();
            BeginField();
            if (value is not null)
            {
                WriteUtf8Field(value.Value, "O");
            }
        }

        // TimeOnly as an Excel-style time serial (fraction of a 24h day): a plain number the parser
        // reconstructs via TryGetDouble, matching how the XLSX/XLSB row writers store it.
        public void Write(TimeOnly value)
        {
            ThrowIfDisposed();
            BeginField();
            WriteUtf8Field(value.Ticks / (double)TimeSpan.TicksPerDay, default);
        }

        public void Write(TimeOnly? value)
        {
            ThrowIfDisposed();
            BeginField();
            if (value is not null)
            {
                WriteUtf8Field(value.Value.Ticks / (double)TimeSpan.TicksPerDay, default);
            }
        }

        // IUtf8SpanFormattable (not ISpanFormattable): CSV output is UTF-8, so numbers format straight
        // to bytes with no char buffer or transcode. Every BCL numeric type implements it.
        public void Write<T>(T value)
            where T : IUtf8SpanFormattable
        {
            ThrowIfDisposed();
            BeginField();
            WriteUtf8Field(value, default);
        }

        public void Write<T>(T? value)
            where T : struct, IUtf8SpanFormattable
        {
            ThrowIfDisposed();
            BeginField();
            if (value is not null)
            {
                WriteUtf8Field(value.Value, default);
            }
        }

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
            // Utf8Formatter is culture-free (no per-field NumberFormatInfo lookup) and matches the
            // InvariantCulture/default-format output below for these types. Guards are JIT constants,
            // so non-matching T compiles them away (mirrors Cell.TryParse's dispatch).
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
            if (format.IsEmpty && typeof(T) == typeof(double) && Utf8Formatter.TryFormat(Unsafe.As<T, double>(ref value), buf, out written))
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

        // Overflow path for a pathologically long formatted value; standard numeric/date/Guid output
        // never reaches here. Grows a pooled buffer until the value fits, then writes it.
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

        // Writes an already-UTF-8 field, quoting only if it contains the delimiter, quote, CR, or LF.
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

        // String fields transcode to UTF-8 exactly once; the special-char scan runs on the chars first
        // so the common (unquoted) case is a single vectorized scan plus one WriteUtf8.
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

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            _owner.EndRow();
        }

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
