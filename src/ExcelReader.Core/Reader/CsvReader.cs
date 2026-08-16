using System.Diagnostics.CodeAnalysis;
using System.Text;

namespace ExcelReader.Core.Reader
{
    /// <summary>
    /// A forward-only reader over a delimited (CSV-style) text source, exposed as a single, unnamed
    /// sheet through the same <see cref="IExcelRowReader"/> surface as the XLSX/XLSB/XLS readers.
    /// </summary>
    /// <remarks>Unlike the XLSX/XLSB/XLS readers, there are no styles or shared strings to resolve.</remarks>
    public sealed partial class CsvReader : IExcelRowReader, IExcelRowReader<CsvReader.Enumerator>
    {
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP008:Don't assign member with injected and created disposables",
            Justification = "Holds either the caller's stream or a transcoding wrapper this reader creates and owns; both branches set _leaveOpen consistently with which one applies.")]
        private readonly Stream? _stream;
        private readonly bool _leaveOpen;
        private readonly CsvReaderOptions _options;
        private readonly ReadOnlyMemory<byte> _memory;
        private readonly long _startPosition = -1;
        private bool _enumeratedOnce;

        internal CsvReader(Stream stream, bool leaveOpen, CsvReaderOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(stream);
            _options = options ?? CsvReaderOptions.Default;
            ValidateOptions(_options);
            if (_options.Encoding is not null && _options.Encoding.CodePage != Encoding.UTF8.CodePage)
            {
                // The transcoding stream is a fresh wrapper we alone own; its own leaveOpen setting
                // already governs whether the underlying stream is closed with it.
                _stream = Encoding.CreateTranscodingStream(stream, _options.Encoding, Encoding.UTF8, leaveOpen);
                _leaveOpen = false;
            }
            else
            {
                _stream = stream;
                _leaveOpen = leaveOpen;
            }
            if (_stream.CanSeek)
            {
                _startPosition = _stream.Position;
            }
            _memory = default;
        }

        internal CsvReader(ReadOnlyMemory<byte> data, CsvReaderOptions? options = null)
        {
            _options = options ?? CsvReaderOptions.Default;
            ValidateOptions(_options);
            _stream = null;
            _leaveOpen = true;
            _memory = Transcode(data, _options.Encoding);
        }

        private static ReadOnlyMemory<byte> Transcode(ReadOnlyMemory<byte> data, Encoding? encoding)
        {
            if (encoding is null || encoding.CodePage == Encoding.UTF8.CodePage)
            {
                return data;
            }
            using MemoryStream source = XlsCompoundFile.AsStream(data);
            using Stream transcoding = Encoding.CreateTranscodingStream(source, encoding, Encoding.UTF8, leaveOpen: true);
            using MemoryStream target = new(data.Length);
            transcoding.CopyTo(target);
            return target.GetBuffer().AsMemory(0, (int)target.Length);
        }

        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Reader ownership transfers to the caller, who disposes it via await using / DisposeAsync.")]
        internal static ValueTask<CsvReader> CreateAsync(Stream stream, bool leaveOpen, CsvReaderOptions? options = null, CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            return new ValueTask<CsvReader>(new CsvReader(stream, leaveOpen, options));
        }

        private static void ValidateOptions(CsvReaderOptions options)
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

        /// <inheritdoc/>
        public bool IsDate1904 => false;

        /// <summary>Gets the sheet name. Always the empty string, since a CSV source has a single, unnamed sheet.</summary>
        public string SheetName => "";

        /// <summary>Gets the sheet count. Always 1, since a CSV source has a single, unnamed sheet.</summary>
        public int SheetCount => 1;

        /// <summary>Gets the sheet name at <paramref name="index"/>. Always the empty string, since a CSV source has a single, unnamed sheet.</summary>
        /// <param name="index">The zero-based sheet index. Must be 0.</param>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is not 0.</exception>
        public string SheetNameAt(int index)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(index);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(index, SheetCount);
            return "";
        }

        /// <summary>Checks whether <paramref name="name"/> matches the (empty) CSV sheet name, case-insensitively.</summary>
        /// <param name="name">The sheet name to look for.</param>
        /// <returns><see langword="true"/> if <paramref name="name"/> is empty; otherwise <see langword="false"/>.</returns>
        public bool TryMoveToSheet(ReadOnlySpan<char> name)
        {
            return name.Equals(SheetName, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>Validates that <paramref name="index"/> is 0, the only valid sheet index for a CSV source.</summary>
        /// <param name="index">The zero-based sheet index. Must be 0.</param>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is not 0.</exception>
        public void MoveToSheet(int index)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(index);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(index, SheetCount);
        }

        /// <summary>Gets an enumerator that reads records synchronously from the start of the source.</summary>
        [SuppressMessage("Performance", "HLQ006:GetEnumerator should return a value type",
            Justification = "Enumerator is a class so the same type can also expose MoveNextAsync for the async path.")]
        public Enumerator GetEnumerator()
        {
            ResetToStart();
            if (_stream is null)
            {
                return new Enumerator(_memory, _options);
            }
            return new Enumerator(_stream, _options);
        }

        IExcelRowEnumerator IExcelRowReader<IExcelRowEnumerator>.GetEnumerator()
        {
            return GetEnumerator();
        }

        /// <summary>Gets an enumerator that reads records asynchronously from the start of the source.</summary>
        [SuppressMessage("Performance", "HLQ006:GetAsyncEnumerator should return a value type",
            Justification = "Enumerator is a class so the same type can also expose MoveNextAsync for the async path.")]
        public Enumerator GetAsyncEnumerator()
        {
            return GetEnumerator();
        }

        IExcelRowEnumerator IExcelRowReader<IExcelRowEnumerator>.GetAsyncEnumerator()
        {
            return GetAsyncEnumerator();
        }

        /// <summary>Asynchronously creates an enumerator that reads records from the start of the source.</summary>
        /// <param name="ct">A token to cancel the setup operation.</param>
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Enumerator ownership transfers to the caller, who disposes it via await using / DisposeAsync.")]
        public ValueTask<Enumerator> GetAsyncEnumeratorAsync(CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            ResetToStart();
            if (_stream is null)
            {
                return new ValueTask<Enumerator>(new Enumerator(_memory, _options, ct));
            }
            return new ValueTask<Enumerator>(new Enumerator(_stream, _options, ct));
        }

        async ValueTask<IExcelRowEnumerator> IExcelRowReader<IExcelRowEnumerator>.GetAsyncEnumeratorAsync(CancellationToken ct)
        {
            return await GetAsyncEnumeratorAsync(ct).ConfigureAwait(false);
        }

        // Lets the same reader be enumerated more than once when the source stream supports seeking
        // (mirrors XlsxReader.GetEnumerator reopening its ZIP entry fresh on every call). Over a
        // non-seekable/transcoding stream there's no position to rewind to, so a second enumeration
        // would silently yield zero rows instead of replaying the file — fail loudly instead.
        private void ResetToStart()
        {
            if (_stream is null)
            {
                return;
            }
            if (_startPosition >= 0)
            {
                _stream.Position = _startPosition;
                return;
            }
            if (_enumeratedOnce)
            {
                throw new InvalidOperationException(
                    "This CsvReader is over a non-seekable stream and can only be enumerated once.");
            }
            _enumeratedOnce = true;
        }

        /// <inheritdoc/>
        public void Dispose()
        {
            if (!_leaveOpen && _stream is not null)
            {
                _stream.Dispose();
            }
        }

        /// <inheritdoc/>
        public ValueTask DisposeAsync()
        {
            if (_leaveOpen || _stream is null)
            {
                return ValueTask.CompletedTask;
            }
            return _stream.DisposeAsync();
        }
    }
}
