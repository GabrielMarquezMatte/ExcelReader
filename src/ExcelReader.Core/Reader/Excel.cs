using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using ExcelReader.Core.Crypto;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Internal;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Parser.Internal;

namespace ExcelReader.Core.Reader
{
    /// <summary>
    /// Entry point for opening Excel and CSV workbooks. <see cref="Open(string,ExcelReaderOptions?)"/>/
    /// <see cref="Open(Stream,bool,ExcelReaderOptions?)"/> and their async counterparts auto-detect XLSX/XLSB/XLS
    /// from the file's signature and return a format-agnostic <see cref="IExcelRowReader"/>; the
    /// <c>From*</c>/<c>FromXls*</c>/<c>FromXlsb*</c>/<c>FromCsv*</c> methods open a specific, known format directly.
    /// </summary>
    public static partial class Excel
    {
        /// <summary>Opens an XLSX workbook from a file path, taking ownership of the file stream.</summary>
        /// <param name="path">The path to the XLSX file.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsxReader FromFile(string path, ExcelReaderOptions? options = null)
        {
            return From(File.OpenRead(path), leaveOpen: false, options);
        }

        /// <summary>Opens an XLSX workbook from an existing stream.</summary>
        /// <param name="stream">The stream containing the XLSX data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the reader is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsxReader From(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null)
        {
            if (TryDecryptCfbStream(stream, leaveOpen, options, out Stream decrypted))
            {
                return new XlsxReader(decrypted, leaveOpen: false, options);
            }
            return new XlsxReader(stream, leaveOpen, options);
        }

        /// <summary>
        /// Opens an XLSX workbook directly from an in-memory buffer. Reads the ZIP
        /// central directory and decompresses parts without a <see cref="ZipArchive"/>
        /// or intermediate <see cref="Stream"/> — every part is fully materialized up front, so the returned
        /// reader never suspends, even under <c>await foreach</c>.
        /// </summary>
        /// <param name="data">The whole XLSX file's bytes. Must outlive the returned reader.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>. <see cref="ExcelReaderOptions.PrefetchDecompression"/> is ignored on this path — there is nothing left to overlap.</param>
        public static XlsxReader From(ReadOnlyMemory<byte> data, ExcelReaderOptions? options = null)
        {
            ExcelReaderOptions effective = options ?? ExcelReaderOptions.Default;
            // Eager decryption, not a DecryptedPackageStream: this overload never suspends.
            if (data.Span.StartsWith(XlsCompoundFile.Signature) && EncryptedPackageOpener.IsEncryptedMemory(data, effective))
            {
                ReadOnlyMemory<byte> plain = EncryptedPackageOpener.DecryptToMemory(data, effective);
                return XlsxReader.CreateFromMemory(plain, effective);
            }
            return XlsxReader.CreateFromMemory(data, effective);
        }

        /// <summary>Opens an XLSX workbook from a file path, taking ownership of the file stream. Alias for <see cref="FromFile(string, ExcelReaderOptions?)"/>, keeping the format-named factory STYLEGUIDE.md requires for every format.</summary>
        /// <param name="path">The path to the XLSX file.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsxReader FromXlsxFile(string path, ExcelReaderOptions? options = null)
        {
            return FromFile(path, options);
        }

        /// <summary>Opens an XLSX workbook from an existing stream. Alias for <see cref="From(Stream, bool, ExcelReaderOptions?)"/>, keeping the format-named factory STYLEGUIDE.md requires for every format.</summary>
        /// <param name="stream">The stream containing the XLSX data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the reader is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsxReader FromXlsx(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null)
        {
            return From(stream, leaveOpen, options);
        }

        /// <summary>Opens an XLSX workbook directly from an in-memory buffer. Alias for <see cref="From(ReadOnlyMemory{byte}, ExcelReaderOptions?)"/>, keeping the format-named factory STYLEGUIDE.md requires for every format.</summary>
        /// <param name="data">The whole XLSX file's bytes. Must outlive the returned reader.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsxReader FromXlsx(ReadOnlyMemory<byte> data, ExcelReaderOptions? options = null)
        {
            return From(data, options);
        }

        /// <summary>Opens a legacy binary (XLS) workbook from a file path, taking ownership of the file stream.</summary>
        /// <param name="path">The path to the XLS file.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsReader FromXlsFile(string path, ExcelReaderOptions? options = null)
        {
            return new XlsReader(File.OpenRead(path), leaveOpen: false, options);
        }

        /// <summary>Opens a legacy binary (XLS) workbook from an existing stream.</summary>
        /// <param name="stream">The stream containing the XLS data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the reader is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsReader FromXls(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null)
        {
            return new XlsReader(stream, leaveOpen, options);
        }

        /// <summary>Opens a legacy binary (XLS) workbook directly from an in-memory buffer.</summary>
        /// <param name="data">The whole XLS file's bytes. Must outlive the returned reader and must not be mutated while it is in use.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsReader FromXls(ReadOnlyMemory<byte> data, ExcelReaderOptions? options = null)
        {
            return new XlsReader(data, options);
        }

        /// <summary>Opens an XLSB (Excel binary) workbook from a file path, taking ownership of the file stream.</summary>
        /// <param name="path">The path to the XLSB file.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsbReader FromXlsbFile(string path, ExcelReaderOptions? options = null)
        {
            return FromXlsb(File.OpenRead(path), leaveOpen: false, options);
        }

        /// <summary>Opens an XLSB (Excel binary) workbook from an existing stream.</summary>
        /// <param name="stream">The stream containing the XLSB data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the reader is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsbReader FromXlsb(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null)
        {
            if (TryDecryptCfbStream(stream, leaveOpen, options, out Stream decrypted))
            {
                return new XlsbReader(decrypted, leaveOpen: false, options);
            }
            return new XlsbReader(stream, leaveOpen, options);
        }

        /// <summary>
        /// Opens an XLSB workbook directly from an in-memory buffer. Reads the ZIP
        /// central directory and decompresses parts without a <see cref="ZipArchive"/>
        /// or intermediate <see cref="Stream"/> — every part is fully materialized up front, so the returned
        /// reader never suspends, even under <c>await foreach</c>.
        /// </summary>
        /// <param name="data">The whole XLSB file's bytes. Must outlive the returned reader.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>. <see cref="ExcelReaderOptions.PrefetchDecompression"/> is ignored on this path — there is nothing left to overlap.</param>
        public static XlsbReader FromXlsb(ReadOnlyMemory<byte> data, ExcelReaderOptions? options = null)
        {
            ExcelReaderOptions effective = options ?? ExcelReaderOptions.Default;
            // Eager decryption, not a DecryptedPackageStream: this overload never suspends.
            if (data.Span.StartsWith(XlsCompoundFile.Signature) && EncryptedPackageOpener.IsEncryptedMemory(data, effective))
            {
                ReadOnlyMemory<byte> plain = EncryptedPackageOpener.DecryptToMemory(data, effective);
                return XlsbReader.CreateFromMemory(plain, effective);
            }
            return XlsbReader.CreateFromMemory(data, effective);
        }

        /// <summary>Asynchronously opens an XLSX workbook from a file path, taking ownership of the file stream.</summary>
        /// <param name="path">The path to the XLSX file.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        public static ValueTask<XlsxReader> FromFileAsync(string path, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            FileStream stream = OpenAsyncFile(path);
            return FromAsync(stream, leaveOpen: false, options, ct);
        }

        /// <summary>Asynchronously opens an XLSX workbook from an existing stream.</summary>
        /// <param name="stream">The stream containing the XLSX data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the reader is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        public static ValueTask<XlsxReader> FromAsync(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            if (TryDecryptCfbStream(stream, leaveOpen, options, out Stream decrypted))
            {
                return XlsxReader.CreateAsync(decrypted, leaveOpen: false, options, ct);
            }
            return XlsxReader.CreateAsync(stream, leaveOpen, options, ct);
        }

        /// <summary>Asynchronously opens an XLSX workbook from a file path, taking ownership of the file stream. Alias for <see cref="FromFileAsync(string, ExcelReaderOptions?, CancellationToken)"/>, keeping the format-named factory STYLEGUIDE.md requires for every format.</summary>
        /// <param name="path">The path to the XLSX file.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        public static ValueTask<XlsxReader> FromXlsxFileAsync(string path, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            return FromFileAsync(path, options, ct);
        }

        /// <summary>Asynchronously opens an XLSX workbook from an existing stream. Alias for <see cref="FromAsync(Stream, bool, ExcelReaderOptions?, CancellationToken)"/>, keeping the format-named factory STYLEGUIDE.md requires for every format.</summary>
        /// <param name="stream">The stream containing the XLSX data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the reader is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        public static ValueTask<XlsxReader> FromXlsxAsync(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            return FromAsync(stream, leaveOpen, options, ct);
        }

        /// <summary>Asynchronously opens a legacy binary (XLS) workbook from a file path, taking ownership of the file stream.</summary>
        /// <param name="path">The path to the XLS file.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        public static ValueTask<XlsReader> FromXlsFileAsync(string path, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            FileStream stream = OpenAsyncFile(path);
            return XlsReader.CreateAsync(stream, leaveOpen: false, options, ct);
        }

        /// <summary>Asynchronously opens a legacy binary (XLS) workbook from an existing stream.</summary>
        /// <param name="stream">The stream containing the XLS data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the reader is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        public static ValueTask<XlsReader> FromXlsAsync(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            return XlsReader.CreateAsync(stream, leaveOpen, options, ct);
        }

        /// <summary>Asynchronously opens an XLSB (Excel binary) workbook from a file path, taking ownership of the file stream.</summary>
        /// <param name="path">The path to the XLSB file.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        public static ValueTask<XlsbReader> FromXlsbFileAsync(string path, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            FileStream stream = OpenAsyncFile(path);
            return FromXlsbAsync(stream, leaveOpen: false, options, ct);
        }

        /// <summary>Asynchronously opens an XLSB (Excel binary) workbook from an existing stream.</summary>
        /// <param name="stream">The stream containing the XLSB data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the reader is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        public static ValueTask<XlsbReader> FromXlsbAsync(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            if (TryDecryptCfbStream(stream, leaveOpen, options, out Stream decrypted))
            {
                return XlsbReader.CreateAsync(decrypted, leaveOpen: false, options, ct);
            }
            return XlsbReader.CreateAsync(stream, leaveOpen, options, ct);
        }

        /// <summary>Opens a CSV (or other delimited-text) source from a file path, taking ownership of the file stream.</summary>
        /// <param name="path">The path to the CSV file.</param>
        /// <param name="options">Delimiter, quote, encoding, and size-limit settings; <see cref="CsvReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static CsvReader FromCsvFile(string path, CsvReaderOptions? options = null)
        {
            return new CsvReader(File.OpenRead(path), leaveOpen: false, options);
        }

        /// <summary>Opens a CSV (or other delimited-text) source from an existing stream.</summary>
        /// <param name="stream">The stream containing the CSV data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the reader is disposed.</param>
        /// <param name="options">Delimiter, quote, encoding, and size-limit settings; <see cref="CsvReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static CsvReader FromCsv(Stream stream, bool leaveOpen = true, CsvReaderOptions? options = null)
        {
            return new CsvReader(stream, leaveOpen, options);
        }

        /// <summary>Opens a CSV (or other delimited-text) source directly from an in-memory buffer.</summary>
        /// <param name="data">The whole CSV source's bytes. Must outlive the returned reader and must not be mutated while it is in use.</param>
        /// <param name="options">Delimiter, quote, encoding, and size-limit settings; <see cref="CsvReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static CsvReader FromCsv(ReadOnlyMemory<byte> data, CsvReaderOptions? options = null)
        {
            return new CsvReader(data, options);
        }

        /// <summary>Asynchronously opens a CSV (or other delimited-text) source from a file path, taking ownership of the file stream.</summary>
        /// <param name="path">The path to the CSV file.</param>
        /// <param name="options">Delimiter, quote, encoding, and size-limit settings; <see cref="CsvReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        public static ValueTask<CsvReader> FromCsvFileAsync(string path, CsvReaderOptions? options = null, CancellationToken ct = default)
        {
            FileStream stream = OpenAsyncFile(path);
            return CsvReader.CreateAsync(stream, leaveOpen: false, options, ct);
        }

        /// <summary>Asynchronously opens a CSV (or other delimited-text) source from an existing stream.</summary>
        /// <param name="stream">The stream containing the CSV data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the reader is disposed.</param>
        /// <param name="options">Delimiter, quote, encoding, and size-limit settings; <see cref="CsvReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        public static ValueTask<CsvReader> FromCsvAsync(Stream stream, bool leaveOpen = true, CsvReaderOptions? options = null, CancellationToken ct = default)
        {
            return CsvReader.CreateAsync(stream, leaveOpen, options, ct);
        }

        /// <summary>
        /// Parses a CSV file into <typeparamref name="T"/> instances across several threads, yielding the
        /// same sequence, in the same order, that the sequential <see cref="Parser.ExcelParser{T}"/> produces.
        /// </summary>
        /// <typeparam name="T">The row model type to bind each CSV record to.</typeparam>
        /// <param name="path">The path of the CSV file to read.</param>
        /// <param name="degreeOfParallelism">The maximum number of parsing threads. <c>0</c> means <see cref="Environment.ProcessorCount"/>; <c>1</c> parses sequentially.</param>
        /// <param name="readerOptions">CSV dialect options. Defaults to <see cref="CsvReaderOptions.Default"/>.</param>
        /// <param name="config">Typed-parsing configuration. Defaults to a new <see cref="Parser.ExcelParserConfig"/>.</param>
        /// <param name="ct">A token to cancel enumeration.</param>
        /// <remarks>
        /// <para>
        /// Parsing and type conversion run in parallel; whatever the caller does per row does not. A caller
        /// whose per-row work dominates will see little gain from raising <paramref name="degreeOfParallelism"/>.
        /// </para>
        /// <para>
        /// Falls back to sequential parsing — same results, one thread — when the source is too small to
        /// partition usefully, or when <see cref="CsvReaderOptions.Encoding"/> is set to a non-UTF-8 encoding,
        /// which must be transcoded sequentially.
        /// </para>
        /// <para>
        /// <see cref="CsvReaderOptions.InternStrings"/> yields substantially less here than on the sequential
        /// path: each partition keeps its own dedup cache, so hit rates fall and memory multiplies. Results
        /// are unaffected.
        /// </para>
        /// </remarks>
        [RequiresUnreferencedCode("Typed parsing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Typed parsing binds property setters at runtime (MethodInfo.CreateDelegate / MakeGenericMethod).")]
        public static IAsyncEnumerable<T> ParseCsvParallelAsync<T>(
            string path,
            int degreeOfParallelism = 0,
            CsvReaderOptions? readerOptions = null,
            ExcelParserConfig? config = null,
            CancellationToken ct = default)
        {
            ArgumentException.ThrowIfNullOrEmpty(path);
            ArgumentOutOfRangeException.ThrowIfNegative(degreeOfParallelism);
            return ParallelCsvFactory.Create<T>(path, degreeOfParallelism, readerOptions, config, ct);
        }

        /// <summary>
        /// Parses an in-memory CSV buffer into <typeparamref name="T"/> instances across several threads,
        /// yielding the same sequence, in the same order, that the sequential <see cref="Parser.ExcelParser{T}"/> produces.
        /// </summary>
        /// <typeparam name="T">The row model type to bind each CSV record to.</typeparam>
        /// <param name="data">The CSV bytes. The caller keeps ownership; the buffer must not be mutated during enumeration.</param>
        /// <param name="degreeOfParallelism">The maximum number of parsing threads. <c>0</c> means <see cref="Environment.ProcessorCount"/>; <c>1</c> parses sequentially.</param>
        /// <param name="readerOptions">CSV dialect options. Defaults to <see cref="CsvReaderOptions.Default"/>.</param>
        /// <param name="config">Typed-parsing configuration. Defaults to a new <see cref="Parser.ExcelParserConfig"/>.</param>
        /// <param name="ct">A token to cancel enumeration.</param>
        /// <remarks>Carries the same fallback and <see cref="CsvReaderOptions.InternStrings"/> caveats as the path-based overload.</remarks>
        [RequiresUnreferencedCode("Typed parsing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Typed parsing binds property setters at runtime (MethodInfo.CreateDelegate / MakeGenericMethod).")]
        public static IAsyncEnumerable<T> ParseCsvParallelAsync<T>(
            ReadOnlyMemory<byte> data,
            int degreeOfParallelism = 0,
            CsvReaderOptions? readerOptions = null,
            ExcelParserConfig? config = null,
            CancellationToken ct = default)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(degreeOfParallelism);
            return ParallelCsvFactory.Create<T>(data, degreeOfParallelism, readerOptions, config, ct);
        }

        /// <summary>
        /// Parses a CSV stream into <typeparamref name="T"/> instances, in parallel where the stream can be
        /// partitioned, yielding the same sequence, in the same order, that the sequential
        /// <see cref="Parser.ExcelParser{T}"/> produces.
        /// </summary>
        /// <typeparam name="T">The row model type to bind each CSV record to.</typeparam>
        /// <param name="stream">The CSV stream, read from its current position. The caller keeps ownership and must not read from it concurrently.</param>
        /// <param name="degreeOfParallelism">The maximum number of parsing threads. <c>0</c> means <see cref="Environment.ProcessorCount"/>; <c>1</c> parses sequentially.</param>
        /// <param name="readerOptions">CSV dialect options. Defaults to <see cref="CsvReaderOptions.Default"/>.</param>
        /// <param name="config">Typed-parsing configuration. Defaults to a new <see cref="Parser.ExcelParserConfig"/>.</param>
        /// <param name="ct">A token to cancel enumeration.</param>
        /// <remarks>
        /// <para>
        /// Only a <see cref="FileStream"/>, or a <see cref="MemoryStream"/> whose buffer is publicly
        /// visible, can be partitioned; every other stream — including seekable ones — parses sequentially,
        /// with identical results. A stream exposes a single mutable position, and a seek can be
        /// arbitrarily expensive, so this overload partitions only what it can read positionally and
        /// cheaply rather than promising parallelism it cannot deliver.
        /// </para>
        /// <para>
        /// Passing a <see cref="FileStream"/> reads its <see cref="FileStream.SafeFileHandle"/>, which
        /// flushes that stream's internal buffer and disables its subsequent buffering optimizations. The
        /// stream's position is not moved. Prefer the path-based overload where a path is available.
        /// </para>
        /// <para>Carries the same <see cref="CsvReaderOptions.InternStrings"/> caveat as the other overloads.</para>
        /// </remarks>
        [RequiresUnreferencedCode("Typed parsing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Typed parsing binds property setters at runtime (MethodInfo.CreateDelegate / MakeGenericMethod).")]
        public static IAsyncEnumerable<T> ParseCsvParallelAsync<T>(
            Stream stream,
            int degreeOfParallelism = 0,
            CsvReaderOptions? readerOptions = null,
            ExcelParserConfig? config = null,
            CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ArgumentOutOfRangeException.ThrowIfNegative(degreeOfParallelism);
            return ParallelCsvFactory.Create<T>(stream, degreeOfParallelism, readerOptions, config, ct);
        }

        // Bounds bytes pulled from an untrusted stream/file, independent of
        // CsvSnifferOptions.MaxSampleLines, which only bounds lines within the sample.
        private const int CsvDialectSampleBytes = 64 * 1024;

        /// <summary>Reads a sample from the start of a seekable stream and infers its CSV dialect. The stream's position is restored before returning.</summary>
        /// <param name="stream">A seekable stream containing delimited-text data.</param>
        /// <param name="options">Candidate delimiters/quotes and the sample-line cap; <see cref="CsvSnifferOptions.Default"/> when <see langword="null"/>.</param>
        /// <exception cref="ArgumentNullException"><paramref name="stream"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="stream"/> does not support seeking.</exception>
        public static CsvDialect SniffCsvDialect(Stream stream, CsvSnifferOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(stream);
            RequireSeekableForSniff(stream);
            long start = stream.Position;
            byte[] buffer = ArrayPool<byte>.Shared.Rent(CsvDialectSampleBytes);
            try
            {
                int read = stream.ReadAtLeast(buffer, CsvDialectSampleBytes, throwOnEndOfStream: false);
                stream.Position = start;
                return CsvSniffer.Detect(buffer.AsSpan(0, read), options ?? CsvSnifferOptions.Default);
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(buffer);
            }
        }

        /// <summary>Infers the CSV dialect of an in-memory buffer, from a sample of its leading bytes.</summary>
        /// <param name="data">The delimited-text data.</param>
        /// <param name="options">Candidate delimiters/quotes and the sample-line cap; <see cref="CsvSnifferOptions.Default"/> when <see langword="null"/>.</param>
        public static CsvDialect SniffCsvDialect(ReadOnlyMemory<byte> data, CsvSnifferOptions? options = null)
        {
            ReadOnlySpan<byte> sample = data.Length > CsvDialectSampleBytes ? data.Span[..CsvDialectSampleBytes] : data.Span;
            return CsvSniffer.Detect(sample, options ?? CsvSnifferOptions.Default);
        }

        /// <summary>Reads a sample from the start of a file and infers its CSV dialect.</summary>
        /// <param name="path">The path to the delimited-text file.</param>
        /// <param name="options">Candidate delimiters/quotes and the sample-line cap; <see cref="CsvSnifferOptions.Default"/> when <see langword="null"/>.</param>
        /// <exception cref="ArgumentNullException"><paramref name="path"/> is <see langword="null"/>.</exception>
        public static CsvDialect SniffCsvDialectFromFile(string path, CsvSnifferOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(path);
            using FileStream stream = File.OpenRead(path);
            return SniffCsvDialect(stream, options);
        }

        /// <summary>Asynchronously reads a sample from the start of a seekable stream and infers its CSV dialect. The stream's position is restored before returning.</summary>
        /// <param name="stream">A seekable stream containing delimited-text data.</param>
        /// <param name="options">Candidate delimiters/quotes and the sample-line cap; <see cref="CsvSnifferOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the read.</param>
        /// <exception cref="ArgumentNullException"><paramref name="stream"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="stream"/> does not support seeking.</exception>
        public static async ValueTask<CsvDialect> SniffCsvDialectAsync(Stream stream, CsvSnifferOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(stream);
            RequireSeekableForSniff(stream);
            long start = stream.Position;
            byte[] buffer = ArrayPool<byte>.Shared.Rent(CsvDialectSampleBytes);
            try
            {
                int read = await stream.ReadAtLeastAsync(buffer, CsvDialectSampleBytes, throwOnEndOfStream: false, ct).ConfigureAwait(false);
                stream.Position = start;
                return CsvSniffer.Detect(buffer.AsSpan(0, read), options ?? CsvSnifferOptions.Default);
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(buffer);
            }
        }

        /// <summary>Asynchronously reads a sample from the start of a file and infers its CSV dialect.</summary>
        /// <param name="path">The path to the delimited-text file.</param>
        /// <param name="options">Candidate delimiters/quotes and the sample-line cap; <see cref="CsvSnifferOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the read.</param>
        /// <exception cref="ArgumentNullException"><paramref name="path"/> is <see langword="null"/>.</exception>
        public static async ValueTask<CsvDialect> SniffCsvDialectFromFileAsync(string path, CsvSnifferOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(path);
            FileStream stream = OpenAsyncFile(path);
            try
            {
                return await SniffCsvDialectAsync(stream, options, ct).ConfigureAwait(false);
            }
            finally
            {
                await stream.DisposeAsync().ConfigureAwait(false);
            }
        }

        private static void RequireSeekableForSniff(Stream stream)
        {
            if (!stream.CanSeek)
            {
                throw new ArgumentException(
                    "SniffCsvDialect requires a seekable stream so a sample can be read and the position restored. Use the ReadOnlyMemory<byte> overload for a non-seekable source.",
                    nameof(stream));
            }
        }

        /// <summary>
        /// Guesses a column schema for <paramref name="reader"/>'s current sheet by sampling it from
        /// the first row, without disturbing any enumerator the caller already holds.
        /// </summary>
        /// <param name="reader">The reader whose current sheet is sampled.</param>
        /// <param name="headerRow">1-based row number to take column names from; 0 means "no header",
        /// so every returned schema is addressable only by <see cref="ExcelColumnSchema.Index"/>.</param>
        /// <param name="sampleSize">How many rows after the header to inspect.</param>
        /// <returns>One <see cref="ExcelColumnSchema"/> per column, in column order.</returns>
        /// <remarks>
        /// This is a guess over a bounded sample, not a guarantee about the whole sheet — a column
        /// whose first <paramref name="sampleSize"/> rows are all integers is reported as
        /// <see cref="ExcelColumnType.Int64Column"/> even if row 10,000 holds text. Verify it fits
        /// before trusting it, and feed the result into <see cref="Parser.ExcelFluentParser{T}"/> to
        /// build a real map.
        /// </remarks>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="headerRow"/> is negative, or
        /// <paramref name="sampleSize"/> is not positive.</exception>
        /// <exception cref="ArgumentException">The sheet has fewer rows than <paramref name="headerRow"/>.</exception>
        public static ExcelColumnSchema[] InferSchema(IExcelRowReader reader, int headerRow = 1, int sampleSize = 100)
        {
            ArgumentNullException.ThrowIfNull(reader);
            using IExcelRowEnumerator rows = reader.GetEnumerator();
            return SchemaInference.Infer(rows, reader.IsDate1904, headerRow, sampleSize);
        }

        // XLSX/XLSB are ZIP ("PK\x03\x04"); XLS is OLE2/CFB. XLSB is told apart from XLSX by
        // "xl/workbook.bin" in the ZIP central directory.
        private static ReadOnlySpan<byte> ZipSignature => [0x50, 0x4B, 0x03, 0x04];

        /// <summary>
        /// Opens a workbook from a file path, auto-detecting its format (XLSX/XLSB/XLS) from the file's signature
        /// and taking ownership of the file stream.
        /// </summary>
        /// <remarks>Pattern-match on the returned reader against its concrete type (<see cref="XlsxReader"/> / <see cref="XlsReader"/> / <see cref="XlsbReader"/>) to access format-specific members.</remarks>
        /// <param name="path">The path to the workbook file.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <returns>A format-agnostic <see cref="IExcelRowReader"/> backed by the concrete reader that matches the detected format.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="path"/> is <see langword="null"/>.</exception>
        /// <exception cref="InvalidDataException">The file's signature does not match a supported format.</exception>
        public static IExcelRowReader Open(string path, ExcelReaderOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(path);
            return OpenSeekable(File.OpenRead(path), leaveOpen: false, options);
        }

        /// <summary>
        /// Opens a workbook from an existing seekable stream, auto-detecting its format (XLSX/XLSB/XLS) from the
        /// stream's signature.
        /// </summary>
        /// <param name="stream">A seekable stream containing the workbook data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the reader is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <returns>A format-agnostic <see cref="IExcelRowReader"/> backed by the concrete reader that matches the detected format.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="stream"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="stream"/> does not support seeking.</exception>
        /// <exception cref="InvalidDataException">The stream's signature does not match a supported format.</exception>
        public static IExcelRowReader Open(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(stream);
            return OpenSeekable(stream, leaveOpen, options);
        }

        /// <summary>
        /// Opens a workbook from an in-memory buffer, auto-detecting its format (XLSX/XLSB/XLS) from its
        /// signature. XLSX/XLSB route through <see cref="ZipMemoryIndex"/> instead of
        /// a <see cref="ZipArchive"/>/<see cref="Stream"/>, so the returned reader never
        /// suspends, even under <c>await foreach</c>.
        /// </summary>
        /// <param name="data">The whole workbook file's bytes. Must outlive the returned reader.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <returns>A format-agnostic <see cref="IExcelRowReader"/> backed by the concrete reader that matches the detected format.</returns>
        /// <exception cref="InvalidDataException">The buffer's signature does not match a supported format.</exception>
        public static IExcelRowReader Open(ReadOnlyMemory<byte> data, ExcelReaderOptions? options = null)
        {
            ExcelReaderOptions effective = options ?? ExcelReaderOptions.Default;
            // Eager decryption, not a DecryptedPackageStream: this overload never suspends.
            if (data.Span.StartsWith(XlsCompoundFile.Signature) && EncryptedPackageOpener.IsEncryptedMemory(data, effective))
            {
                ReadOnlyMemory<byte> plain = EncryptedPackageOpener.DecryptToMemory(data, effective);
                return OpenFromPlainMemory(plain, effective);
            }
            ExcelFileFormat format = ClassifyMemory(data, effective, out ZipMemoryIndex? memZip);
            if (format is ExcelFileFormat.Unknown)
            {
                memZip?.Dispose();
                UnknownFormatException();
            }
            return format switch
            {
                ExcelFileFormat.Xls => new XlsReader(data, effective),
                ExcelFileFormat.Xlsb => XlsbReader.CreateFromMemory(memZip!, effective),
                ExcelFileFormat.Xlsx => XlsxReader.CreateFromMemory(memZip!, effective),
                _ => throw new System.Diagnostics.UnreachableException(),
            };
        }

        // A genuinely decrypted OOXML package is always Xlsb/Xlsx; anything else means decryption
        // produced something that isn't a workbook.
        private static IExcelRowReader OpenFromPlainMemory(ReadOnlyMemory<byte> plain, ExcelReaderOptions options)
        {
            ExcelFileFormat format = ClassifyMemory(plain, options, out ZipMemoryIndex? memZip);
            if (format is not (ExcelFileFormat.Xlsb or ExcelFileFormat.Xlsx))
            {
                memZip?.Dispose();
                UnknownFormatException();
            }
            return format switch
            {
                ExcelFileFormat.Xlsb => XlsbReader.CreateFromMemory(memZip!, options),
                ExcelFileFormat.Xlsx => XlsxReader.CreateFromMemory(memZip!, options),
                _ => throw new System.Diagnostics.UnreachableException(),
            };
        }

        // The ZIP central directory must be walked to tell XLSB from XLSX, so the resulting
        // ZipMemoryIndex is handed back for reuse instead of being parsed a second time.
        private static ExcelFileFormat ClassifyMemory(ReadOnlyMemory<byte> data, ExcelReaderOptions options, out ZipMemoryIndex? memZip)
        {
            memZip = null;
            ReadOnlySpan<byte> span = data.Span;
            ReadOnlySpan<byte> header = span.Length > 8 ? span[..8] : span;
            if (TryClassifyHeader(header, out ExcelFileFormat format))
            {
                return format;
            }
            if (header.StartsWith(XlsCompoundFile.Signature))
            {
                // Same two-stage CFB probe as DetectSeekable: only the OLE directory (not the 8-byte
                // signature) can tell a legacy .xls apart from an encrypted OOXML package.
                return EncryptedPackageOpener.IsEncryptedMemory(data, options)
                    ? ExcelFileFormat.EncryptedOoxml
                    : ExcelFileFormat.Xls;
            }
            memZip = ZipMemoryIndex.Create(data, options);
            return memZip.TryGetEntry("xl/workbook.bin"u8, out _) ? ExcelFileFormat.Xlsb : ExcelFileFormat.Xlsx;
        }

        /// <summary>
        /// Asynchronously opens a workbook from a file path, auto-detecting its format (XLSX/XLSB/XLS) from the
        /// file's signature and taking ownership of the file stream.
        /// </summary>
        /// <param name="path">The path to the workbook file.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        /// <returns>A format-agnostic <see cref="IExcelRowReader"/> backed by the concrete reader that matches the detected format.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="path"/> is <see langword="null"/>.</exception>
        /// <exception cref="InvalidDataException">The file's signature does not match a supported format.</exception>
        public static ValueTask<IExcelRowReader> OpenAsync(string path, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(path);
            FileStream stream = OpenAsyncFile(path);
            return OpenSeekableAsync(stream, leaveOpen: false, options, ct);
        }

        /// <summary>
        /// Asynchronously opens a workbook from an existing seekable stream, auto-detecting its format
        /// (XLSX/XLSB/XLS) from the stream's signature.
        /// </summary>
        /// <param name="stream">A seekable stream containing the workbook data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the reader is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        /// <returns>A format-agnostic <see cref="IExcelRowReader"/> backed by the concrete reader that matches the detected format.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="stream"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="stream"/> does not support seeking.</exception>
        /// <exception cref="InvalidDataException">The stream's signature does not match a supported format.</exception>
        public static ValueTask<IExcelRowReader> OpenAsync(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(stream);
            return OpenSeekableAsync(stream, leaveOpen, options, ct);
        }

        /// <summary>Detects a workbook's file format (XLSX/XLSB/XLS/unknown) from a seekable stream's signature, without consuming it.</summary>
        /// <param name="stream">A seekable stream containing the workbook data. The stream's position is restored after detection.</param>
        /// <returns>The detected <see cref="ExcelFileFormat"/>, or <see cref="ExcelFileFormat.Unknown"/> if the signature matches no supported format.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="stream"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="stream"/> does not support seeking.</exception>
        public static ExcelFileFormat DetectFileFormat(Stream stream)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ExcelFileFormat format = DetectSeekable(stream, out ZipArchive? zip);
            zip?.Dispose();
            return format;
        }
        /// <summary>
        /// Detects a workbook's file format (XLSX/XLSB/XLS/unknown) from an in-memory buffer's signature, without consuming it.
        /// </summary>
        /// <param name="data">A buffer containing the workbook data.</param>
        /// <returns>The detected <see cref="ExcelFileFormat"/>, or <see cref="ExcelFileFormat.Unknown"/> if the signature matches no supported format.</returns>
        public static ExcelFileFormat DetectFileFormat(ReadOnlyMemory<byte> data)
        {
            var format = ClassifyMemory(data, ExcelReaderOptions.Default, out ZipMemoryIndex? memZip);
            memZip?.Dispose();
            return format;
        }

        /// <summary>Asynchronously detects a workbook's file format (XLSX/XLSB/XLS/unknown) from a seekable stream's signature, without consuming it.</summary>
        /// <param name="stream">A seekable stream containing the workbook data. The stream's position is restored after detection.</param>
        /// <param name="ct">A token to cancel the detection operation.</param>
        /// <returns>The detected <see cref="ExcelFileFormat"/>, or <see cref="ExcelFileFormat.Unknown"/> if the signature matches no supported format.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="stream"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="stream"/> does not support seeking.</exception>
        public static async ValueTask<ExcelFileFormat> DetectFileFormatAsync(Stream stream, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(stream);
            (ExcelFileFormat format, ZipArchive? zip) = await DetectSeekableAsync(stream, ct).ConfigureAwait(false);
            if (zip is not null)
            {
                await ZipArchiveDisposal.DisposeAsync(zip).ConfigureAwait(false);
            }
            return format;
        }

        private static IExcelRowReader OpenSeekable(Stream stream, bool leaveOpen, ExcelReaderOptions? options)
        {
            ExcelFileFormat format;
            ZipArchive? zip = null;
            try
            {
                format = DetectSeekable(stream, out zip);
            }
            catch
            {
                zip?.Dispose();
                DisposeOnFailure(stream, leaveOpen);
                throw;
            }
            if (format is ExcelFileFormat.Unknown)
            {
                UnknownFormat(stream, leaveOpen);
            }
            if (format is ExcelFileFormat.EncryptedOoxml)
            {
                ExcelReaderOptions effective = options ?? ExcelReaderOptions.Default;
                Stream decrypted = EncryptedPackageOpener.Decrypt(stream, leaveOpen, effective);
                return OpenDecryptedZip(decrypted, effective);
            }
            return format switch
            {
                ExcelFileFormat.Xls => new XlsReader(stream, leaveOpen, options),
                ExcelFileFormat.Xlsb => new XlsbReader(stream, leaveOpen, zip!, options),
                ExcelFileFormat.Xlsx => new XlsxReader(stream, leaveOpen, zip!, options),
                _ => throw new System.Diagnostics.UnreachableException(),
            };
        }

        // `decrypted` is a brand-new stream nobody else references, so it's always handed to the
        // chosen reader with leaveOpen:false; disposing that reader cascades to the CFB container.
        private static IExcelRowReader OpenDecryptedZip(Stream decrypted, ExcelReaderOptions options)
        {
            ZipArchive? zip = null;
            try
            {
                ExcelFileFormat zipFormat = ClassifyZipStream(decrypted, start: 0, out ZipArchive zipPeek);
                zip = zipPeek;
                return zipFormat switch
                {
                    ExcelFileFormat.Xlsb => new XlsbReader(decrypted, leaveOpen: false, zip, options),
                    ExcelFileFormat.Xlsx => new XlsxReader(decrypted, leaveOpen: false, zip, options),
                    _ => throw new System.Diagnostics.UnreachableException(),
                };
            }
            catch
            {
                zip?.Dispose();
                decrypted.Dispose();
                throw;
            }
        }

        private static async ValueTask<IExcelRowReader> OpenSeekableAsync(Stream stream, bool leaveOpen, ExcelReaderOptions? options, CancellationToken ct)
        {
            ExcelFileFormat format;
            ZipArchive? zip = null;
            try
            {
                (format, zip) = await DetectSeekableAsync(stream, ct).ConfigureAwait(false);
            }
            catch
            {
                if (zip is not null)
                {
                    await ZipArchiveDisposal.DisposeAsync(zip).ConfigureAwait(false);
                }
                await DisposeOnFailureAsync(stream, leaveOpen).ConfigureAwait(false);
                throw;
            }
            if (format is ExcelFileFormat.Unknown)
            {
                await DisposeOnFailureAsync(stream, leaveOpen).ConfigureAwait(false);
                UnknownFormatException();
            }
            if (format is ExcelFileFormat.EncryptedOoxml)
            {
                ExcelReaderOptions effective = options ?? ExcelReaderOptions.Default;
                Stream decrypted = EncryptedPackageOpener.Decrypt(stream, leaveOpen, effective);
                return await OpenDecryptedZipAsync(decrypted, effective, ct).ConfigureAwait(false);
            }
            return format switch
            {
                ExcelFileFormat.Xls => await XlsReader.CreateAsync(stream, leaveOpen, options, ct).ConfigureAwait(false),
                ExcelFileFormat.Xlsb => await XlsbReader.CreateFromOpenZipAsync(stream, leaveOpen, zip!, options, ct).ConfigureAwait(false),
                ExcelFileFormat.Xlsx => await XlsxReader.CreateFromOpenZipAsync(stream, leaveOpen, zip!, options, ct).ConfigureAwait(false),
                _ => throw new System.Diagnostics.UnreachableException(),
            };
        }

        // Async twin of OpenDecryptedZip: the central-directory peek is synchronous, but reader
        // construction goes through CreateFromOpenZipAsync so worksheet reads stay fully async after.
        private static async ValueTask<IExcelRowReader> OpenDecryptedZipAsync(Stream decrypted, ExcelReaderOptions options, CancellationToken ct)
        {
            ZipArchive? zip = null;
            try
            {
                ExcelFileFormat zipFormat = ClassifyZipStream(decrypted, start: 0, out ZipArchive zipPeek);
                zip = zipPeek;
                return zipFormat switch
                {
                    ExcelFileFormat.Xlsb => await XlsbReader.CreateFromOpenZipAsync(decrypted, leaveOpen: false, zip, options, ct).ConfigureAwait(false),
                    ExcelFileFormat.Xlsx => await XlsxReader.CreateFromOpenZipAsync(decrypted, leaveOpen: false, zip, options, ct).ConfigureAwait(false),
                    _ => throw new System.Diagnostics.UnreachableException(),
                };
            }
            catch
            {
                if (zip is not null)
                {
                    await ZipArchiveDisposal.DisposeAsync(zip).ConfigureAwait(false);
                }
                await decrypted.DisposeAsync().ConfigureAwait(false);
                throw;
            }
        }

        private static void DisposeOnFailure(Stream stream, bool leaveOpen)
        {
            if (!leaveOpen)
            {
                stream.Dispose();
            }
        }

        private static ValueTask DisposeOnFailureAsync(Stream stream, bool leaveOpen)
        {
            return leaveOpen ? ValueTask.CompletedTask : stream.DisposeAsync();
        }

        [DoesNotReturn]
        private static void UnknownFormat(Stream stream, bool leaveOpen)
        {
            DisposeOnFailure(stream, leaveOpen);
            UnknownFormatException();
        }

        [DoesNotReturn]
        private static void UnknownFormatException()
        {
            throw new InvalidDataException("Unrecognized file format; expected an XLSX/XLSB (ZIP) or XLS (OLE2) workbook.");
        }

        // Returns true (with the final answer) only for Unknown; false means the caller must probe
        // further — a ZIP central directory (XLSB vs XLSX) or, for a CFB signature, the OLE directory
        // (legacy .xls vs encrypted OOXML). Callers distinguish the two "false" cases by re-checking
        // `sig` themselves, since `format` carries no signal here.
        private static bool TryClassifyHeader(ReadOnlySpan<byte> sig, out ExcelFileFormat format)
        {
            if (sig.StartsWith(XlsCompoundFile.Signature) || sig.StartsWith(ZipSignature))
            {
                format = default;
                return false;
            }
            format = ExcelFileFormat.Unknown;
            return true;
        }

        // Peeks the central directory to distinguish XLSB from XLSX; kept open so the caller can hand
        // the archive straight to the chosen reader instead of re-parsing it.
        private static ExcelFileFormat ClassifyZipStream(Stream stream, long start, out ZipArchive zip)
        {
            var zipPeek = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
            zip = zipPeek; // assigned before GetEntry can throw, so a caller-side catch can dispose it
            bool isXlsb = zipPeek.GetEntry("xl/workbook.bin") is not null;
            stream.Position = start;
            return isXlsb ? ExcelFileFormat.Xlsb : ExcelFileFormat.Xlsx;
        }

        // `zip` receives the archive opened to peek the central directory (null for Xls/Unknown) so the
        // caller can hand it straight to the chosen reader instead of re-parsing it.
        [SkipLocalsInit]
        private static ExcelFileFormat DetectSeekable(Stream stream, out ZipArchive? zip)
        {
            zip = null;
            RequireSeekable(stream);
            long start = stream.Position;
            Span<byte> header = stackalloc byte[8];
            int read = stream.ReadAtLeast(header, header.Length, throwOnEndOfStream: false);
            stream.Position = start;
            ReadOnlySpan<byte> sig = header[..read];
            if (TryClassifyHeader(sig, out ExcelFileFormat format))
            {
                return format;
            }
            if (sig.StartsWith(XlsCompoundFile.Signature))
            {
                return EncryptedPackageOpener.IsEncryptedContainer(stream)
                    ? ExcelFileFormat.EncryptedOoxml
                    : ExcelFileFormat.Xls;
            }
            ExcelFileFormat zipFormat = ClassifyZipStream(stream, start, out ZipArchive zipPeek);
            zip = zipPeek;
            return zipFormat;
        }

        private static async ValueTask<(ExcelFileFormat Format, ZipArchive? Zip)> DetectSeekableAsync(Stream stream, CancellationToken ct)
        {
            RequireSeekable(stream);
            long start = stream.Position;
            byte[] header = new byte[8];
            int read = await stream.ReadAtLeastAsync(header, header.Length, throwOnEndOfStream: false, ct).ConfigureAwait(false);
            stream.Position = start;
            ReadOnlySpan<byte> sig = header.AsSpan(0, read);
            if (TryClassifyHeader(sig, out ExcelFileFormat format))
            {
                return (format, null);
            }
            if (sig.StartsWith(XlsCompoundFile.Signature))
            {
                ExcelFileFormat cfbFormat = EncryptedPackageOpener.IsEncryptedContainer(stream)
                    ? ExcelFileFormat.EncryptedOoxml
                    : ExcelFileFormat.Xls;
                return (cfbFormat, null);
            }
            ExcelFileFormat zipFormat = ClassifyZipStream(stream, start, out ZipArchive zip);
            return (zipFormat, zip);
        }

        // If the caller supplied a password and the stream is a CFB container, decrypt it up front so
        // the ZIP-based reader constructor gets a plaintext ZIP instead of a confusing "not a ZIP" error.
        private static bool TryDecryptCfbStream(Stream stream, bool leaveOpen, ExcelReaderOptions? options, out Stream decrypted)
        {
            if (options?.Password is not null && stream.CanSeek && HasCfbSignature(stream))
            {
                decrypted = EncryptedPackageOpener.Decrypt(stream, leaveOpen, options);
                return true;
            }
            decrypted = stream;
            return false;
        }

        private static bool HasCfbSignature(Stream stream)
        {
            long start = stream.Position;
            Span<byte> header = stackalloc byte[8];
            int read = stream.ReadAtLeast(header, header.Length, throwOnEndOfStream: false);
            stream.Position = start;
            return header[..read].StartsWith(XlsCompoundFile.Signature);
        }

        private static void RequireSeekable(Stream stream)
        {
            if (!stream.CanSeek)
            {
                throw new ArgumentException(
                    "Open requires a seekable stream so the format signature can be detected. Buffer the stream first, or call From/FromXls/FromXlsb directly.",
                    nameof(stream));
            }
        }

        private static FileStream OpenAsyncFile(string path)
        {
            return new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 65536,
                                  options: FileOptions.Asynchronous | FileOptions.SequentialScan);
        }
    }
}
