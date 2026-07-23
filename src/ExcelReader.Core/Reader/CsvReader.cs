using System.Diagnostics.CodeAnalysis;
using System.Text;

namespace ExcelReader.Core.Reader
{
    // Forward-only CSV reader. Unlike the XLSX/XLSB/XLS readers there are no styles or shared strings.
    // It exposes a single, unnamed sheet so it can be driven through the same format-agnostic
    // IExcelRowReader surface (row enumeration + trivial sheet navigation) as the Excel readers.
    public sealed partial class CsvReader : IExcelRowReader, IExcelRowReader<CsvReader.Enumerator>
    {
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP008:Don't assign member with injected and created disposables",
            Justification = "Holds either the caller's stream or a transcoding wrapper this reader creates and owns; both branches set _leaveOpen consistently with which one applies.")]
        private readonly Stream _stream;
        private readonly bool _leaveOpen;
        private readonly CsvReaderOptions _options;
        private readonly long _startPosition = -1;
        private bool _enumeratedOnce;

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP003:Dispose previous before re-assigning",
            Justification = "Readonly field, first and only assignment in this constructor.")]
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
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Reader ownership transfers to the caller, who disposes it via await using / DisposeAsync.")]
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

        public bool IsDate1904 => false;

        // CSV is a single, unnamed sheet. These satisfy the IExcelRowReader surface so a CSV reader can
        // be driven through the same format-agnostic loop as the Excel readers.
        public string SheetName => "";
        public int SheetCount => 1;

        public bool TryMoveToSheet(ReadOnlySpan<char> name)
        {
            return name.Equals(SheetName, StringComparison.OrdinalIgnoreCase);
        }

        public void MoveToSheet(int index)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(index);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(index, SheetCount);
        }

        [SuppressMessage("Performance", "HLQ006:GetEnumerator should return a value type",
            Justification = "Enumerator is a class so the same type can also expose MoveNextAsync for the async path.")]
        public Enumerator GetEnumerator()
        {
            ResetToStart();
            return new Enumerator(_stream, _options);
        }

        IExcelRowEnumerator IExcelRowReader<IExcelRowEnumerator>.GetEnumerator()
        {
            return GetEnumerator();
        }

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

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Enumerator ownership transfers to the caller, who disposes it via await using / DisposeAsync.")]
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Enumerator ownership transfers to the caller, who disposes it via await using / DisposeAsync.")]
        public ValueTask<Enumerator> GetAsyncEnumeratorAsync(CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            ResetToStart();
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

        public void Dispose()
        {
            if (!_leaveOpen)
            {
                _stream.Dispose();
            }
        }

        public ValueTask DisposeAsync()
        {
            return _leaveOpen ? ValueTask.CompletedTask : _stream.DisposeAsync();
        }
    }
}
