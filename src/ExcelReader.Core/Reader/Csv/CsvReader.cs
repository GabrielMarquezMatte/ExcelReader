using System.Diagnostics.CodeAnalysis;
using System.Text;
using ExcelReader.Core.Parser.ParallelCsv;
using ExcelReader.Core.Reader.Internal;
using ExcelReader.Core.Reader.Sources;
using ExcelReader.Core.Reader.Xls;
using Microsoft.Win32.SafeHandles;

namespace ExcelReader.Core.Reader.Csv
{
    /// <summary>An open delimited-text (CSV-style) source, exposed as a workbook with a single, unnamed sheet.</summary>
    /// <remarks>Unlike the XLSX/XLSB/XLS readers, there are no styles or shared strings to resolve.</remarks>
    public sealed partial class CsvReader : IExcelWorkbook
    {
        [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed", Justification = "Disposed by ReleaseResources, which the lifetime runs after the last enumerator.")]
        private readonly ByteSource? _source;
        private readonly long _sourceStart;
        [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed", Justification = "An alias of the stream the ByteSource owns; ReleaseResources disposes the source.")]
        private readonly FileStream? _file;
        [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed", Justification = "Disposed by ReleaseResources, which the lifetime runs after the last enumerator.")]
        private readonly Stream? _stream;
        private readonly bool _leaveOpen;
        private readonly CsvReaderOptions _options;
        private readonly ReadOnlyMemory<byte> _memory;
        private readonly ExcelSheetList<CsvSheet> _sheetList;
        private int _enumeratedOnce;
        private SafeFileHandle? _chunkHandle;

        internal ReaderLifetime Lifetime { get; }

        internal CsvReader(Stream stream, bool leaveOpen, CsvReaderOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(stream);
            _options = options ?? CsvReaderOptions.Default;
            ValidateOptions(_options);
            Lifetime = new ReaderLifetime(ReleaseResources);
            if (stream.CanSeek)
            {
                _sourceStart = stream.Position;
                _file = stream as FileStream;
                _source = ByteSource.FromStream(stream, leaveOpen);
                _leaveOpen = true;
            }
            else if (NeedsTranscoding(_options.Encoding))
            {
                _stream = Encoding.CreateTranscodingStream(stream, _options.Encoding!, Encoding.UTF8, leaveOpen);
                _leaveOpen = false;
            }
            else
            {
                _stream = stream;
                _leaveOpen = leaveOpen;
            }
            _sheetList = new ExcelSheetList<CsvSheet>([new CsvSheet(this)]);
        }

        private static bool NeedsTranscoding(Encoding? encoding)
        {
            return encoding is not null && encoding.CodePage != Encoding.UTF8.CodePage;
        }

        internal CsvReader(ReadOnlyMemory<byte> data, CsvReaderOptions? options = null)
        {
            _options = options ?? CsvReaderOptions.Default;
            ValidateOptions(_options);
            Lifetime = new ReaderLifetime(ReleaseResources);
            _stream = null;
            _leaveOpen = true;
            _memory = Transcode(data, _options.Encoding);
            _sheetList = new ExcelSheetList<CsvSheet>([new CsvSheet(this)]);
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
            if (_source is null && _stream is null)
            {
                source = new CsvChunkSource(_memory);
                return true;
            }
            if (_file is not null && !NeedsTranscoding(_options.Encoding))
            {
                SafeFileHandle handle = ChunkHandle(_file);
                source = new CsvChunkSource(handle, RandomAccess.GetLength(handle), _sourceStart);
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
            if (!NeedsTranscoding(encoding))
            {
                return data;
            }
            using MemoryStream source = XlsCompoundFile.AsStream(data);
            using Stream transcoding = Encoding.CreateTranscodingStream(source, encoding!, Encoding.UTF8, leaveOpen: true);
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

        /// <summary>Gets the sheet count. Always 1, since a CSV source has a single, unnamed sheet.</summary>
        public int SheetCount => 1;

        /// <summary>Gets the source's sheets: always exactly one, unnamed. Reading the list opens nothing.</summary>
        /// <exception cref="ObjectDisposedException">The reader was disposed.</exception>
        public ExcelSheetList<CsvSheet> Sheets
        {
            get
            {
                Lifetime.ThrowIfClosed(this);
                return _sheetList;
            }
        }

        /// <summary>Checks whether <paramref name="name"/> is the (empty) name of the CSV sheet, ignoring case.</summary>
        /// <param name="name">The sheet name to look for.</param>
        /// <param name="sheet">The CSV sheet, when <paramref name="name"/> is empty.</param>
        /// <returns><see langword="true"/> if <paramref name="name"/> is empty; otherwise <see langword="false"/>.</returns>
        /// <exception cref="ObjectDisposedException">The reader was disposed.</exception>
        public bool TryGetSheet(ReadOnlySpan<char> name, out CsvSheet sheet)
        {
            Lifetime.ThrowIfClosed(this);
            bool found = name.IsEmpty;
            sheet = found ? _sheetList[0] : default;
            return found;
        }

        IExcelSheet IExcelWorkbook.SheetAt(int index)
        {
            return Sheets[index];
        }

        bool IExcelWorkbook.TryGetSheet(ReadOnlySpan<char> name, [MaybeNullWhen(false)] out IExcelSheet sheet)
        {
            bool found = TryGetSheet(name, out CsvSheet typed);
            sheet = found ? typed : null;
            return found;
        }

        /// <summary>Gets the source's only sheet.</summary>
        /// <exception cref="ObjectDisposedException">The workbook was disposed.</exception>
        public CsvSheet FirstSheet => Sheets[0];

        IExcelSheet IExcelWorkbook.FirstSheet => FirstSheet;

        internal Enumerator OpenSheet(CancellationToken ct = default)
        {
            Lifetime.Acquire(this);
            try
            {
                if (_source is not null)
                {
                    return new Enumerator(OpenSourceStream(_source), _options, ownsSource: true, Lifetime, ct);
                }
                if (_stream is null)
                {
                    return new Enumerator(_memory, _options, Lifetime, ct);
                }
                if (Interlocked.Exchange(ref _enumeratedOnce, 1) != 0)
                {
                    throw new InvalidOperationException(
                        "This CsvReader is over a non-seekable stream and can only be enumerated once.");
                }
                return new Enumerator(_stream, _options, ownsSource: false, Lifetime, ct);
            }
            catch
            {
                Lifetime.Release();
                throw;
            }
        }

        private Stream OpenSourceStream(ByteSource source)
        {
            bool transcode = NeedsTranscoding(_options.Encoding);
            ByteSourceStream raw = new(source, _sourceStart, source.Length - _sourceStart, buffered: transcode);
            return transcode
                ? Encoding.CreateTranscodingStream(raw, _options.Encoding!, Encoding.UTF8, leaveOpen: false)
                : raw;
        }

        /// <inheritdoc/>
        public void Dispose()
        {
            Lifetime.Close();
        }

        /// <inheritdoc/>
        public ValueTask DisposeAsync()
        {
            Lifetime.Close();
            return ValueTask.CompletedTask;
        }

        private void ReleaseResources()
        {
            _chunkHandle?.Dispose();
            _source?.Dispose();
            if (!_leaveOpen)
            {
                _stream?.Dispose();
            }
        }
    }
}
