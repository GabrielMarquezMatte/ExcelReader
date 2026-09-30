using System.Text;
using ExcelReader.Core.Parser.ParallelCsv;
using ExcelReader.Core.Reader.Xls;
using Microsoft.Win32.SafeHandles;

namespace ExcelReader.Core.Reader.Csv
{
    /// <summary>
    /// A forward-only reader over a delimited (CSV-style) text source, exposed as a single, unnamed
    /// sheet through the same <see cref="IExcelRowReader"/> surface as the XLSX/XLSB/XLS readers.
    /// </summary>
    /// <remarks>Unlike the XLSX/XLSB/XLS readers, there are no styles or shared strings to resolve.</remarks>
    public sealed partial class CsvReader : IExcelRowReader, IExcelRowReader<CsvReader.Enumerator>
    {
        private readonly Stream? _stream;
        private readonly bool _leaveOpen;
        private readonly CsvReaderOptions _options;
        private readonly ReadOnlyMemory<byte> _memory;
        private readonly long _startPosition = -1;
        private bool _enumeratedOnce;
        private SafeFileHandle? _chunkHandle;

        internal CsvReader(Stream stream, bool leaveOpen, CsvReaderOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(stream);
            _options = options ?? CsvReaderOptions.Default;
            ValidateOptions(_options);
            if (_options.Encoding is not null && _options.Encoding.CodePage != Encoding.UTF8.CodePage)
            {
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

        internal CsvReaderOptions Options
        {
            get
            {
                return _options;
            }
        }

        internal bool TryGetChunkSource(out CsvChunkSource source)
        {
            if (_stream is null)
            {
                source = new CsvChunkSource(_memory);
                return true;
            }
            if (_stream is FileStream file && _startPosition >= 0)
            {
                SafeFileHandle handle = ChunkHandle(file);
                source = new CsvChunkSource(handle, RandomAccess.GetLength(handle), _startPosition);
                return true;
            }
            source = default;
            return false;
        }

        // A synchronous Windows handle serializes every read on its file object; parallel chunks need an overlapped one.
        // ponytail: reopens by path, safe because Excel's file opens deny write/delete sharing; switch to
        // ReOpenFile if a caller-supplied FileStream that allows delete ever reaches this.
        private SafeFileHandle ChunkHandle(FileStream file)
        {
            if (!OperatingSystem.IsWindows())
            {
                return file.SafeFileHandle;
            }
            try
            {
                return _chunkHandle ??= File.OpenHandle(file.Name, FileMode.Open, FileAccess.Read, FileShare.Read,
                    FileOptions.Asynchronous | FileOptions.RandomAccess);
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
            {
                return file.SafeFileHandle;
            }
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

        /// <summary>Gets the sheet's visibility. Always <see cref="ExcelSheetVisibility.Visible"/>: delimited text has no tab bar to hide from.</summary>
        public ExcelSheetVisibility SheetVisibility => ExcelSheetVisibility.Visible;

        /// <summary>Gets the visibility of the sheet at <paramref name="index"/>. Always <see cref="ExcelSheetVisibility.Visible"/>.</summary>
        /// <param name="index">The zero-based sheet index. Must be 0.</param>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is not 0.</exception>
        public ExcelSheetVisibility SheetVisibilityAt(int index)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(index);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(index, SheetCount);
            return ExcelSheetVisibility.Visible;
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
        /// <param name="ct">A token observed by every <c>MoveNextAsync</c> call.</param>
        public Enumerator GetAsyncEnumerator(CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            ResetToStart();
            if (_stream is null)
            {
                return new Enumerator(_memory, _options, ct);
            }
            return new Enumerator(_stream, _options, ct);
        }

        IExcelRowEnumerator IExcelRowReader<IExcelRowEnumerator>.GetAsyncEnumerator(CancellationToken ct)
        {
            return GetAsyncEnumerator(ct);
        }

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
            _chunkHandle?.Dispose();
            if (!_leaveOpen && _stream is not null)
            {
                _stream.Dispose();
            }
        }

        /// <inheritdoc/>
        public ValueTask DisposeAsync()
        {
            _chunkHandle?.Dispose();
            if (_leaveOpen || _stream is null)
            {
                return ValueTask.CompletedTask;
            }
            return _stream.DisposeAsync();
        }
    }
}
