using System.Buffers;
using System.Text;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Csv
{
    /// <summary>
    /// A minimal RFC4180 writer for CSV files: rows are buffered and flushed straight to the
    /// underlying stream, with no sheets, styles, or shared strings — instead of going through the
    /// ZIP/BIFF machinery the other writers need.
    /// </summary>
    /// <remarks>
    /// Delimiter/Quote mirror <c>CsvReaderOptions</c> so a written file needs no reader configuration
    /// to round-trip.
    /// </remarks>
    internal sealed class CsvWriter : IDisposable, IAsyncDisposable
    {
        // ponytail: same 64 KB flush threshold as the XLSX XlsxSheetWriter — bounds memory on huge
        private const int FlushThreshold = 64 * 1024;

        private readonly Stream _stream;
        private readonly bool _leaveOpen;
        private readonly byte _delimiter;
        private readonly byte _quote;
        private readonly SearchValues<byte> _specialBytes;
        private readonly SearchValues<char> _specialChars;
        private readonly bool _lineFeedOnly;
        private readonly BiffBuffer _buffer = new(4096);
        private CsvRowWriter? _rowWriter;
        private bool _rowActive;
        private bool _faulted;
        private bool _disposed;

        private CsvWriter(Stream stream, bool leaveOpen, CsvWriterOptions options)
        {
            _stream = stream;
            _leaveOpen = leaveOpen;
            _delimiter = options.Delimiter;
            _quote = options.Quote;
            _lineFeedOnly = options.NewLine is CsvNewLine.LineFeed;
            ReadOnlySpan<byte> specialBytes = [_delimiter, _quote, (byte)'\r', (byte)'\n'];
            ReadOnlySpan<char> specialChars = [(char)_delimiter, (char)_quote, '\r', '\n'];
            _specialBytes = SearchValues.Create(specialBytes);
            _specialChars = SearchValues.Create(specialChars);
        }

        /// <summary>
        /// Creates a writer that emits CSV text to <paramref name="stream"/>.
        /// </summary>
        /// <param name="stream">The destination stream.</param>
        /// <param name="leaveOpen">If <see langword="true"/>, the stream is left open when the writer is disposed.</param>
        /// <param name="options">Dialect, encoding, and record-terminator options; defaults to <see cref="CsvWriterOptions.Default"/> when omitted.</param>
        /// <exception cref="ArgumentException">The delimiter and quote byte are the same, either is a carriage return or line feed, or <see cref="CsvWriterOptions.NewLine"/> is not a defined value.</exception>
        public static CsvWriter Create(Stream stream, bool leaveOpen = false, CsvWriterOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(stream);
            CsvWriterOptions effective = options ?? CsvWriterOptions.Default;
            ValidateOptions(effective);
            Encoding? target = Transcoded(effective.Encoding);
            if (effective.WriteByteOrderMark)
            {
                stream.Write((target ?? Encoding.UTF8).GetPreamble());
            }
            if (target is null)
            {
                return new CsvWriter(stream, leaveOpen, effective);
            }
            return new CsvWriter(Encoding.CreateTranscodingStream(stream, target, Encoding.UTF8, leaveOpen), leaveOpen: false, effective);
        }

        private static Encoding? Transcoded(Encoding? encoding)
        {
            return encoding is null || encoding.CodePage == Encoding.UTF8.CodePage ? null : encoding;
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
            if (options.NewLine is not (CsvNewLine.CarriageReturnLineFeed or CsvNewLine.LineFeed))
            {
                throw new ArgumentException($"NewLine must be a defined CsvNewLine value; got {options.NewLine}.", nameof(options));
            }
        }

        /// <summary>
        /// Begins writing the next row, returning a reused <see cref="CsvRowWriter"/> that must be
        /// disposed before another row can be started.
        /// </summary>
        /// <exception cref="ObjectDisposedException">The writer has already been disposed.</exception>
        /// <exception cref="InvalidOperationException">The previous row's <see cref="CsvRowWriter"/> has not been disposed yet.</exception>
        public CsvRowWriter StartRow()
        {
            ObjectDisposedException.ThrowIf(_disposed || _faulted, this);
            if (_rowActive)
            {
                throw new InvalidOperationException("The previous CsvRowWriter must be disposed before starting a new row.");
            }
            _rowActive = true;
            _rowWriter ??= new CsvRowWriter(this, _buffer, _delimiter, _quote, _specialBytes, _specialChars);
            _rowWriter.Reset();
            return _rowWriter;
        }

        private void WriteNewLine()
        {
            _buffer.Write(_lineFeedOnly ? "\n"u8 : "\r\n"u8);
        }

        internal void EndRow()
        {
            WriteNewLine();
            _rowActive = false;
            if (_buffer.Length >= FlushThreshold)
            {
                try
                {
                    FlushCore();
                }
                catch
                {
                    Fault();
                    throw;
                }
            }
        }

        internal ValueTask EndRowAsync(CancellationToken ct = default)
        {
            WriteNewLine();
            _rowActive = false;
            return _buffer.Length >= FlushThreshold ? FlushBufferOrFaultAsync(ct) : ValueTask.CompletedTask;
        }

        private async ValueTask FlushBufferOrFaultAsync(CancellationToken ct)
        {
            try
            {
                await FlushBufferAsync(ct).ConfigureAwait(false);
            }
            catch
            {
                Fault();
                throw;
            }
        }

        private async ValueTask FlushBufferAsync(CancellationToken ct)
        {
            await _stream.WriteAsync(_buffer.Memory, ct).ConfigureAwait(false);
            _buffer.Reset();
        }

        /// <summary>
        /// Writes any buffered rows to the underlying stream and flushes it.
        /// </summary>
        /// <exception cref="ObjectDisposedException">The writer has already been disposed, or an earlier write failed.</exception>
        public void Flush()
        {
            ObjectDisposedException.ThrowIf(_disposed || _faulted, this);
            try
            {
                FlushCore();
                _stream.Flush();
            }
            catch
            {
                Fault();
                throw;
            }
        }

        /// <summary>
        /// Writes any buffered rows to the underlying stream and flushes it asynchronously.
        /// </summary>
        /// <exception cref="ObjectDisposedException">The writer has already been disposed, or an earlier write failed.</exception>
        public async ValueTask FlushAsync(CancellationToken ct = default)
        {
            ObjectDisposedException.ThrowIf(_disposed || _faulted, this);
            ct.ThrowIfCancellationRequested();
            try
            {
                if (_buffer.Length > 0)
                {
                    await FlushBufferAsync(ct).ConfigureAwait(false);
                }
                await _stream.FlushAsync(ct).ConfigureAwait(false);
            }
            catch
            {
                Fault();
                throw;
            }
        }

        private void FlushCore()
        {
            if (_buffer.Length > 0)
            {
                _stream.Write(_buffer.Span);
                _buffer.Reset();
            }
        }

        // A failed write leaves an unknown prefix of the buffer in the stream, so writing it again could
        // silently corrupt the output: the writer becomes terminal and only Dispose remains meaningful.
        private void Fault()
        {
            if (_faulted)
            {
                return;
            }
            _faulted = true;
            _buffer.Dispose();
        }

        private void ReleaseStream()
        {
            if (_leaveOpen)
            {
                return;
            }
            if (_faulted)
            {
                FailureCleanup.Dispose(_stream);
                return;
            }
            _stream.Dispose();
        }

        private ValueTask ReleaseStreamAsync()
        {
            if (_leaveOpen)
            {
                return ValueTask.CompletedTask;
            }
            return _faulted ? FailureCleanup.DisposeAsync(_stream) : _stream.DisposeAsync();
        }

        private void TerminateOpenRow()
        {
            if (_rowActive)
            {
                WriteNewLine();
                _rowActive = false;
            }
        }

        /// <summary>
        /// Terminates any open row, flushes buffered output, and closes the underlying stream unless
        /// the writer was created with <c>leaveOpen: true</c>.
        /// </summary>
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            try
            {
                if (!_faulted)
                {
                    TerminateOpenRow();
                    FlushCore();
                }
            }
            catch
            {
                Fault();
                ReleaseStream();
                throw;
            }
            _buffer.Dispose();
            ReleaseStream();
        }

        /// <summary>
        /// Terminates any open row, flushes buffered output, and closes the underlying stream unless
        /// the writer was created with <c>leaveOpen: true</c>.
        /// </summary>
        public async ValueTask DisposeAsync()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            try
            {
                if (!_faulted)
                {
                    TerminateOpenRow();
                    if (_buffer.Length > 0)
                    {
                        await _stream.WriteAsync(_buffer.Memory).ConfigureAwait(false);
                    }
                }
            }
            catch
            {
                Fault();
                await ReleaseStreamAsync().ConfigureAwait(false);
                throw;
            }
            _buffer.Dispose();
            await ReleaseStreamAsync().ConfigureAwait(false);
        }
    }
}
