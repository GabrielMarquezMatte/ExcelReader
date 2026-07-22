using System.Buffers;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
{
    // Minimal RFC4180 writer: no sheets, styles, or shared strings, so rows are buffered and
    // flushed straight to the stream instead of going through the ZIP/BIFF machinery the other
    // writers need. Delimiter/Quote mirror CsvReaderOptions so a written file needs no reader
    // configuration to round-trip.
    public sealed class CsvWriter : IDisposable, IAsyncDisposable
    {
        // ponytail: same 64 KB flush threshold as the XLSX XlsxSheetWriter — bounds memory on huge
        // files while turning many tiny row writes into a handful of big stream writes, and keeps
        // the pooled backing array under the LOH threshold instead of parking it there permanently.
        private const int FlushThreshold = 64 * 1024;

        private readonly Stream _stream;
        private readonly bool _leaveOpen;
        private readonly byte _delimiter;
        private readonly byte _quote;
        // Built once here and shared with the (reused) CsvRowWriter so per-field quote detection is a
        // vectorized scan; CR/LF join the delimiter/quote as the bytes that force a field to be quoted.
        private readonly SearchValues<byte> _specialBytes;
        private readonly SearchValues<char> _specialChars;
        private readonly BiffBuffer _buffer = new(4096);
        private CsvRowWriter? _rowWriter;
        private bool _rowActive;
        private bool _disposed;

        private CsvWriter(Stream stream, bool leaveOpen, CsvWriterOptions options)
        {
            _stream = stream;
            _leaveOpen = leaveOpen;
            _delimiter = options.Delimiter;
            _quote = options.Quote;
            // Via explicitly-typed spans so this binds to SearchValues.Create(ReadOnlySpan<T>) on both
            // target frameworks (net8.0 has no params overload) with no intermediate array allocation.
            ReadOnlySpan<byte> specialBytes = [_delimiter, _quote, (byte)'\r', (byte)'\n'];
            ReadOnlySpan<char> specialChars = [(char)_delimiter, (char)_quote, '\r', '\n'];
            _specialBytes = SearchValues.Create(specialBytes);
            _specialChars = SearchValues.Create(specialChars);
        }

        public static CsvWriter Create(Stream stream, bool leaveOpen = false, CsvWriterOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(stream);
            CsvWriterOptions effective = options ?? CsvWriterOptions.Default;
            ValidateOptions(effective);
            return new CsvWriter(stream, leaveOpen, effective);
        }

        private static void ValidateOptions(CsvWriterOptions options)
        {
            if (options.Delimiter == options.Quote)
            {
                throw new ArgumentException("Delimiter and Quote must be different bytes.", nameof(options));
            }
            if (options.Delimiter is (byte)'\r' or (byte)'\n')
            {
                throw new ArgumentException("Delimiter cannot be a carriage return or line feed.", nameof(options));
            }
            if (options.Quote is (byte)'\r' or (byte)'\n')
            {
                throw new ArgumentException("Quote cannot be a carriage return or line feed.", nameof(options));
            }
        }

        public CsvRowWriter StartRow()
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (_rowActive)
            {
                throw new InvalidOperationException("The previous CsvRowWriter must be disposed before starting a new row.");
            }
            _rowActive = true;
            _rowWriter ??= new CsvRowWriter(this, _buffer, _delimiter, _quote, _specialBytes, _specialChars);
            _rowWriter.Reset();
            return _rowWriter;
        }

        internal void EndRow()
        {
            _buffer.Write("\r\n"u8);
            _rowActive = false;
            if (_buffer.Length >= FlushThreshold)
            {
                _stream.Write(_buffer.Span);
                _buffer.Reset();
            }
        }

        internal ValueTask EndRowAsync(CancellationToken ct = default)
        {
            _buffer.Write("\r\n"u8);
            _rowActive = false;
            if (_buffer.Length >= FlushThreshold)
            {
                return FlushBufferAsync(ct);
            }
            return ValueTask.CompletedTask;
        }

        private async ValueTask FlushBufferAsync(CancellationToken ct)
        {
            await _stream.WriteAsync(_buffer.Memory, ct).ConfigureAwait(false);
            _buffer.Reset();
        }

        public void Flush()
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            FlushCore();
            _stream.Flush();
        }

        public async ValueTask FlushAsync(CancellationToken ct = default)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            ct.ThrowIfCancellationRequested();
            if (_buffer.Length > 0)
            {
                await FlushBufferAsync(ct).ConfigureAwait(false);
            }
            await _stream.FlushAsync(ct).ConfigureAwait(false);
        }

        private void FlushCore()
        {
            if (_buffer.Length > 0)
            {
                _stream.Write(_buffer.Span);
                _buffer.Reset();
            }
        }

        // Emits the terminating newline for a row left open when the writer is disposed (the caller
        // disposed the writer without disposing the last CsvRowWriter). Buffer-only, so it is safe on
        // both the sync and async dispose paths.
        private void TerminateOpenRow()
        {
            if (_rowActive)
            {
                _buffer.Write("\r\n"u8);
                _rowActive = false;
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            TerminateOpenRow();
            FlushCore();
            _buffer.Dispose();
            if (!_leaveOpen)
            {
                _stream.Dispose();
            }
        }

        public async ValueTask DisposeAsync()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            TerminateOpenRow();
            if (_buffer.Length > 0)
            {
                await _stream.WriteAsync(_buffer.Memory).ConfigureAwait(false);
            }
            _buffer.Dispose();
            if (!_leaveOpen)
            {
                await _stream.DisposeAsync().ConfigureAwait(false);
            }
        }
    }
}
