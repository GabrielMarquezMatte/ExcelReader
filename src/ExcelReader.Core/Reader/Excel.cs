using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using ExcelReader.Core.Crypto;
using ExcelReader.Core.Enums;
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
        public static XlsxReader FromXlsxFile(string path, ExcelReaderOptions? options = null)
        {
            return FromXlsx(File.OpenRead(path), leaveOpen: false, options);
        }

        /// <summary>Opens an XLSX workbook from an existing stream.</summary>
        /// <param name="stream">The stream containing the XLSX data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the reader is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsxReader FromXlsx(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null)
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
        public static XlsxReader FromXlsx(ReadOnlyMemory<byte> data, ExcelReaderOptions? options = null)
        {
            ExcelReaderOptions effective = options ?? ExcelReaderOptions.Default;
            if (data.Span.StartsWith(XlsCompoundFile.Signature) && EncryptedPackageOpener.IsEncryptedMemory(data, effective))
            {
                ReadOnlyMemory<byte> plain = EncryptedPackageOpener.DecryptToMemory(data, effective);
                return XlsxReader.CreateFromMemory(plain, effective);
            }
            return XlsxReader.CreateFromMemory(data, effective);
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
        public static ValueTask<XlsxReader> FromXlsxFileAsync(string path, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            FileStream stream = OpenAsyncFile(path);
            return FromXlsxAsync(stream, leaveOpen: false, options, ct);
        }

        /// <summary>Asynchronously opens an XLSX workbook from an existing stream.</summary>
        /// <param name="stream">The stream containing the XLSX data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the reader is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        public static ValueTask<XlsxReader> FromXlsxAsync(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            if (TryDecryptCfbStream(stream, leaveOpen, options, out Stream decrypted))
            {
                return XlsxReader.CreateAsync(decrypted, leaveOpen: false, options, ct);
            }
            return XlsxReader.CreateAsync(stream, leaveOpen, options, ct);
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
        /// <param name="options">Parallelism and dialect options. Defaults to <see cref="CsvParallelOptions.Default"/>. <see cref="CsvParallelOptions.HeaderRow"/> does not apply here and must be left at <c>0</c>; <paramref name="config"/> locates the header.</param>
        /// <param name="config">Typed-parsing configuration. Defaults to a new <see cref="Parser.ExcelParserConfig"/>.</param>
        /// <param name="ct">A token to cancel enumeration.</param>
        /// <remarks>
        /// <para>
        /// Parsing and type conversion run in parallel; whatever the caller does per row does not. A caller
        /// whose per-row work dominates will see little gain from raising <see cref="CsvParallelOptions.DegreeOfParallelism"/>.
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
            CsvParallelOptions? options = null,
            ExcelParserConfig? config = null,
            CancellationToken ct = default)
        {
            ArgumentException.ThrowIfNullOrEmpty(path);
            CsvParallelOptions parallel = ValidatedForTypedParse(options);
            return ParallelCsvFactory.Create<T>(path, parallel.DegreeOfParallelism, parallel.Reader, config, ct);
        }

        /// <summary>
        /// Parses an in-memory CSV buffer into <typeparamref name="T"/> instances across several threads,
        /// yielding the same sequence, in the same order, that the sequential <see cref="Parser.ExcelParser{T}"/> produces.
        /// </summary>
        /// <typeparam name="T">The row model type to bind each CSV record to.</typeparam>
        /// <param name="data">The CSV bytes. The caller keeps ownership; the buffer must not be mutated during enumeration.</param>
        /// <param name="options">Parallelism and dialect options. Defaults to <see cref="CsvParallelOptions.Default"/>. <see cref="CsvParallelOptions.HeaderRow"/> does not apply here and must be left at <c>0</c>; <paramref name="config"/> locates the header.</param>
        /// <param name="config">Typed-parsing configuration. Defaults to a new <see cref="Parser.ExcelParserConfig"/>.</param>
        /// <param name="ct">A token to cancel enumeration.</param>
        /// <remarks>Carries the same fallback and <see cref="CsvReaderOptions.InternStrings"/> caveats as the path-based overload.</remarks>
        [RequiresUnreferencedCode("Typed parsing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Typed parsing binds property setters at runtime (MethodInfo.CreateDelegate / MakeGenericMethod).")]
        public static IAsyncEnumerable<T> ParseCsvParallelAsync<T>(
            ReadOnlyMemory<byte> data,
            CsvParallelOptions? options = null,
            ExcelParserConfig? config = null,
            CancellationToken ct = default)
        {
            CsvParallelOptions parallel = ValidatedForTypedParse(options);
            return ParallelCsvFactory.Create<T>(data, parallel.DegreeOfParallelism, parallel.Reader, config, ct);
        }

        /// <summary>
        /// Parses a CSV stream into <typeparamref name="T"/> instances, in parallel where the stream can be
        /// partitioned, yielding the same sequence, in the same order, that the sequential
        /// <see cref="Parser.ExcelParser{T}"/> produces.
        /// </summary>
        /// <typeparam name="T">The row model type to bind each CSV record to.</typeparam>
        /// <param name="stream">The CSV stream, read from its current position. The caller keeps ownership and must not read from it concurrently.</param>
        /// <param name="options">Parallelism and dialect options. Defaults to <see cref="CsvParallelOptions.Default"/>. <see cref="CsvParallelOptions.HeaderRow"/> does not apply here and must be left at <c>0</c>; <paramref name="config"/> locates the header.</param>
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
            CsvParallelOptions? options = null,
            ExcelParserConfig? config = null,
            CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(stream);
            CsvParallelOptions parallel = ValidatedForTypedParse(options);
            return ParallelCsvFactory.Create<T>(stream, parallel.DegreeOfParallelism, parallel.Reader, config, ct);
        }

        /// <summary>
        /// Reads an in-memory CSV buffer across several threads, parses every record into a
        /// <typeparamref name="TRecord"/> and folds it into a <typeparamref name="TAccumulator"/>, one
        /// instance per partition, merged in buffer order.
        /// </summary>
        /// <typeparam name="TAccumulator">The accumulator type. See <see cref="ICsvAccumulator{TSelf, TModel}"/> for the contract it must honor.</typeparam>
        /// <typeparam name="TRecord">The record type. A record whose <see cref="ICsvRecord{TSelf}.TryParse"/> returns <see langword="false"/> is skipped.</typeparam>
        /// <param name="data">The CSV bytes. The caller keeps ownership; the buffer must not be mutated during processing.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>The accumulator holding every record.</returns>
        /// <remarks>
        /// <para>
        /// Partitions start at guessed record boundaries; one whose guess landed inside a quoted field
        /// spanning lines is read again into a new accumulator, as <see cref="ICsvAccumulator{TSelf, TModel}"/> describes.
        /// </para>
        /// <para>
        /// Falls back to one sequential pass, with a single accumulator and no call to
        /// <see cref="ICsvAccumulator{TSelf, TModel}.Merge"/>, when the source is too small to partition usefully
        /// or when <see cref="CsvReaderOptions.Encoding"/> is set to a non-UTF-8 encoding.
        /// </para>
        /// </remarks>
        public static Task<TAccumulator> AggregateCsvParallelAsync<TAccumulator, TRecord>(
            ReadOnlyMemory<byte> data,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
            where TAccumulator : ICsvAccumulator<TAccumulator, TRecord>, new()
            where TRecord : ICsvRecord<TRecord>, allows ref struct
        {
            return ParallelCsvProcessor.RunAsync(data, RecordAggregation<TAccumulator, TRecord>.Instance, null, Validated(options), ct);
        }

        /// <summary>
        /// Reads a CSV file across several threads, parses every record into a <typeparamref name="TRecord"/>
        /// and folds it into a <typeparamref name="TAccumulator"/>, one instance per partition, merged in file order.
        /// </summary>
        /// <typeparam name="TAccumulator">The accumulator type. See <see cref="ICsvAccumulator{TSelf, TModel}"/> for the contract it must honor.</typeparam>
        /// <typeparam name="TRecord">The record type. A record whose <see cref="ICsvRecord{TSelf}.TryParse"/> returns <see langword="false"/> is skipped.</typeparam>
        /// <param name="path">The path of the CSV file to read.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>The accumulator holding every record.</returns>
        /// <remarks>Carries the same contract and fallbacks as the in-memory overload.</remarks>
        public static Task<TAccumulator> AggregateCsvParallelAsync<TAccumulator, TRecord>(
            string path,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
            where TAccumulator : ICsvAccumulator<TAccumulator, TRecord>, new()
            where TRecord : ICsvRecord<TRecord>, allows ref struct
        {
            ArgumentException.ThrowIfNullOrEmpty(path);
            return ParallelCsvProcessor.RunAsync(path, RecordAggregation<TAccumulator, TRecord>.Instance, null, Validated(options), ct);
        }

        /// <summary>
        /// Reads a CSV stream, in parallel where the stream can be partitioned, parses every record into a
        /// <typeparamref name="TRecord"/> and folds it into a <typeparamref name="TAccumulator"/>, one instance
        /// per partition, merged in stream order.
        /// </summary>
        /// <typeparam name="TAccumulator">The accumulator type. See <see cref="ICsvAccumulator{TSelf, TModel}"/> for the contract it must honor.</typeparam>
        /// <typeparam name="TRecord">The record type. A record whose <see cref="ICsvRecord{TSelf}.TryParse"/> returns <see langword="false"/> is skipped.</typeparam>
        /// <param name="stream">The CSV stream, read from its current position. The caller keeps ownership and must not read from it concurrently.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>The accumulator holding every record.</returns>
        /// <remarks>
        /// Carries the same contract and fallbacks as the in-memory overload, and partitions only a
        /// <see cref="FileStream"/> or a <see cref="MemoryStream"/> whose buffer is publicly visible;
        /// every other stream is read sequentially.
        /// </remarks>
        public static Task<TAccumulator> AggregateCsvParallelAsync<TAccumulator, TRecord>(
            Stream stream,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
            where TAccumulator : ICsvAccumulator<TAccumulator, TRecord>, new()
            where TRecord : ICsvRecord<TRecord>, allows ref struct
        {
            ArgumentNullException.ThrowIfNull(stream);
            return ParallelCsvProcessor.RunAsync(stream, RecordAggregation<TAccumulator, TRecord>.Instance, null, Validated(options), ct);
        }

        /// <summary>
        /// Reads a CSV file across several threads, parses every record into a <typeparamref name="TRecord"/>
        /// and hands it to <paramref name="body"/>.
        /// </summary>
        /// <typeparam name="TRecord">The record type. A record whose <see cref="ICsvRecord{TSelf}.TryParse"/> returns <see langword="false"/> is skipped.</typeparam>
        /// <param name="path">The path of the CSV file to read.</param>
        /// <param name="body">Receives each record. See the remarks for the contract it must honor.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>A task that completes when every record has been delivered.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="body"/> is <see langword="null"/>.</exception>
        /// <remarks>
        /// <para>
        /// <paramref name="body"/> runs concurrently on worker threads and must be safe to call from
        /// several at once; the caller owns any synchronization. Records arrive in no particular order.
        /// </para>
        /// <para>
        /// <paramref name="body"/> may be invoked more than once for the same record. A partition whose start
        /// was guessed wrongly is read again from the confirmed offset, and the discarded speculative pass may
        /// already have delivered a whole partition's worth of records — 1 MiB to 64 MiB of them, not just a
        /// few near the seam. Those records may also be misparsed: a wrong start shifts every field boundary,
        /// so a quoted field holding a delimiter can split into values that appear nowhere in the source. An
        /// exception <paramref name="body"/> throws during a discarded pass is discarded with it.
        /// </para>
        /// <para>
        /// A start is only guessed wrongly when it falls inside a quoted field, so a source in which the quote
        /// character never appears delivers every record exactly once. For that guarantee on any source, use
        /// <see cref="AggregateCsvParallelAsync{TAccumulator, TRecord}(string, CsvParallelOptions?, CancellationToken)"/>,
        /// where the discarded partition's accumulator is thrown away.
        /// </para>
        /// <para>
        /// A <see cref="ReadOnlySpan{T}"/> column points into a pooled buffer the worker reuses once
        /// <paramref name="body"/> returns. Copy anything that must outlive the call.
        /// </para>
        /// <para>
        /// Falls back to one sequential pass, delivering records once each in source order on a single
        /// thread, when the source is too small to partition usefully or when
        /// <see cref="CsvReaderOptions.Encoding"/> is set to a non-UTF-8 encoding.
        /// </para>
        /// </remarks>
        public static Task ForEachCsvParallelAsync<TRecord>(
            string path,
            CsvRecordAction<TRecord> body,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
            where TRecord : ICsvRecord<TRecord>, allows ref struct
        {
            ArgumentException.ThrowIfNullOrEmpty(path);
            ArgumentNullException.ThrowIfNull(body);
            return ParallelCsvProcessor.RunAsync(path, RecordCallback<TRecord>.For(body), null, Validated(options), ct);
        }

        /// <summary>
        /// Reads an in-memory CSV buffer across several threads, parses every record into a
        /// <typeparamref name="TRecord"/> and hands it to <paramref name="body"/>.
        /// </summary>
        /// <typeparam name="TRecord">The record type. A record whose <see cref="ICsvRecord{TSelf}.TryParse"/> returns <see langword="false"/> is skipped.</typeparam>
        /// <param name="data">The CSV bytes. The caller keeps ownership; the buffer must not be mutated during processing.</param>
        /// <param name="body">Receives each record. See the remarks for the contract it must honor.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>A task that completes when every record has been delivered.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="body"/> is <see langword="null"/>.</exception>
        /// <remarks>Carries the same contract and fallbacks as the path-based overload.</remarks>
        public static Task ForEachCsvParallelAsync<TRecord>(
            ReadOnlyMemory<byte> data,
            CsvRecordAction<TRecord> body,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
            where TRecord : ICsvRecord<TRecord>, allows ref struct
        {
            ArgumentNullException.ThrowIfNull(body);
            return ParallelCsvProcessor.RunAsync(data, RecordCallback<TRecord>.For(body), null, Validated(options), ct);
        }

        /// <summary>
        /// Reads a CSV stream, in parallel where the stream can be partitioned, parses every record into a
        /// <typeparamref name="TRecord"/> and hands it to <paramref name="body"/>.
        /// </summary>
        /// <typeparam name="TRecord">The record type. A record whose <see cref="ICsvRecord{TSelf}.TryParse"/> returns <see langword="false"/> is skipped.</typeparam>
        /// <param name="stream">The CSV stream, read from its current position. The caller keeps ownership and must not read from it concurrently.</param>
        /// <param name="body">Receives each record. See the remarks for the contract it must honor.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>A task that completes when every record has been delivered.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="stream"/> or <paramref name="body"/> is <see langword="null"/>.</exception>
        /// <remarks>
        /// Carries the same contract and fallbacks as the path-based overload, and partitions only a
        /// <see cref="FileStream"/> or a <see cref="MemoryStream"/> whose buffer is publicly visible;
        /// every other stream is read sequentially.
        /// </remarks>
        public static Task ForEachCsvParallelAsync<TRecord>(
            Stream stream,
            CsvRecordAction<TRecord> body,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
            where TRecord : ICsvRecord<TRecord>, allows ref struct
        {
            ArgumentNullException.ThrowIfNull(stream);
            ArgumentNullException.ThrowIfNull(body);
            return ParallelCsvProcessor.RunAsync(stream, RecordCallback<TRecord>.For(body), null, Validated(options), ct);
        }

        /// <summary>
        /// Reads a CSV file across several threads, binds every record to a <typeparamref name="TModel"/>
        /// through <paramref name="map"/> and hands it to <paramref name="body"/>.
        /// </summary>
        /// <typeparam name="TModel">The model type. May be a <see langword="ref struct"/> with <see cref="ReadOnlySpan{T}"/> of <see cref="byte"/> properties.</typeparam>
        /// <param name="path">The path of the CSV file to read.</param>
        /// <param name="map">How columns bind to <typeparamref name="TModel"/>.</param>
        /// <param name="body">Receives each record. See the remarks for the contract it must honor.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>A task that completes when every record has been delivered.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="map"/> or <paramref name="body"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="map"/> binds columns by header name and <see cref="CsvParallelOptions.HeaderRow"/> is 0.</exception>
        /// <remarks>
        /// <para>
        /// Carries the same concurrency, repeat-invocation and span-lifetime contract as the
        /// <see cref="ICsvRecord{TSelf}"/> overloads. A property that fails to parse keeps its default
        /// unless <see cref="Parser.ExcelParserConfig.ThrowOnParseFailure"/> is set; a missing
        /// <c>[ExcelRequired]</c> column or value throws <see cref="Parser.ExcelParseException"/>; empty
        /// records are skipped.
        /// </para>
        /// <para>
        /// <see cref="Parser.ExcelParseException.Row"/> is exact when the source is read sequentially and
        /// 0 when it is partitioned, since a partition does not know how many records precede it.
        /// </para>
        /// </remarks>
        public static Task ForEachCsvParallelAsync<TModel>(
            string path,
            CsvModelMap<TModel> map,
            CsvRecordAction<TModel> body,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
            where TModel : allows ref struct
        {
            ArgumentException.ThrowIfNullOrEmpty(path);
            CsvParallelOptions validated = Validated(options);
            return ParallelCsvProcessor.RunAsync(path, MappedCallback<TModel>.Unbound, BoundCallback(map, body, validated), validated, ct);
        }

        /// <summary>
        /// Reads an in-memory CSV buffer across several threads, binds every record to a
        /// <typeparamref name="TModel"/> through <paramref name="map"/> and hands it to <paramref name="body"/>.
        /// </summary>
        /// <typeparam name="TModel">The model type. May be a <see langword="ref struct"/> with <see cref="ReadOnlySpan{T}"/> of <see cref="byte"/> properties.</typeparam>
        /// <param name="data">The CSV bytes. The caller keeps ownership; the buffer must not be mutated during processing.</param>
        /// <param name="map">How columns bind to <typeparamref name="TModel"/>.</param>
        /// <param name="body">Receives each record. See the remarks for the contract it must honor.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>A task that completes when every record has been delivered.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="map"/> or <paramref name="body"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="map"/> binds columns by header name and <see cref="CsvParallelOptions.HeaderRow"/> is 0.</exception>
        /// <remarks>Carries the same contract and fallbacks as the path-based overload.</remarks>
        public static Task ForEachCsvParallelAsync<TModel>(
            ReadOnlyMemory<byte> data,
            CsvModelMap<TModel> map,
            CsvRecordAction<TModel> body,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
            where TModel : allows ref struct
        {
            CsvParallelOptions validated = Validated(options);
            return ParallelCsvProcessor.RunAsync(data, MappedCallback<TModel>.Unbound, BoundCallback(map, body, validated), validated, ct);
        }

        /// <summary>
        /// Reads a CSV stream, in parallel where the stream can be partitioned, binds every record to a
        /// <typeparamref name="TModel"/> through <paramref name="map"/> and hands it to <paramref name="body"/>.
        /// </summary>
        /// <typeparam name="TModel">The model type. May be a <see langword="ref struct"/> with <see cref="ReadOnlySpan{T}"/> of <see cref="byte"/> properties.</typeparam>
        /// <param name="stream">The CSV stream, read from its current position. The caller keeps ownership and must not read from it concurrently.</param>
        /// <param name="map">How columns bind to <typeparamref name="TModel"/>.</param>
        /// <param name="body">Receives each record. See the remarks for the contract it must honor.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>A task that completes when every record has been delivered.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="stream"/>, <paramref name="map"/> or <paramref name="body"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="map"/> binds columns by header name and <see cref="CsvParallelOptions.HeaderRow"/> is 0.</exception>
        /// <remarks>
        /// Carries the same contract and fallbacks as the path-based overload, and the stream restrictions
        /// of the <see cref="ICsvRecord{TSelf}"/> stream overload.
        /// </remarks>
        public static Task ForEachCsvParallelAsync<TModel>(
            Stream stream,
            CsvModelMap<TModel> map,
            CsvRecordAction<TModel> body,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
            where TModel : allows ref struct
        {
            ArgumentNullException.ThrowIfNull(stream);
            CsvParallelOptions validated = Validated(options);
            return ParallelCsvProcessor.RunAsync(stream, MappedCallback<TModel>.Unbound, BoundCallback(map, body, validated), validated, ct);
        }

        /// <summary>
        /// Reads an in-memory CSV buffer across several threads, binds every record to a
        /// <typeparamref name="TModel"/> through <paramref name="map"/> and folds it into a
        /// <typeparamref name="TAccumulator"/>, one instance per partition, merged in buffer order.
        /// </summary>
        /// <typeparam name="TAccumulator">The accumulator type. See <see cref="ICsvAccumulator{TSelf, TModel}"/> for the contract it must honor.</typeparam>
        /// <typeparam name="TModel">The model type.</typeparam>
        /// <param name="data">The CSV bytes. The caller keeps ownership; the buffer must not be mutated during processing.</param>
        /// <param name="map">How columns bind to <typeparamref name="TModel"/>.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>The accumulator holding every record.</returns>
        /// <exception cref="ArgumentException"><paramref name="map"/> binds columns by header name and <see cref="CsvParallelOptions.HeaderRow"/> is 0.</exception>
        /// <remarks>
        /// <para>
        /// The header is bound once, before any record is folded. A property that fails to parse keeps its
        /// default unless <see cref="ExcelParserConfig.ThrowOnParseFailure"/> is set; a missing
        /// <c>[ExcelRequired]</c> column or value throws <see cref="ExcelParseException"/>; empty records are skipped.
        /// </para>
        /// <para>
        /// <see cref="ExcelParseException.Row"/> is exact when the source is read sequentially and 0 when it is
        /// partitioned, since a partition does not know how many records precede it.
        /// </para>
        /// <para>Otherwise carries the same contract and fallbacks as the <see cref="ICsvRecord{TSelf}"/> overloads.</para>
        /// </remarks>
        public static Task<TAccumulator> AggregateCsvParallelAsync<TAccumulator, TModel>(
            ReadOnlyMemory<byte> data,
            CsvModelMap<TModel> map,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
            where TAccumulator : ICsvAccumulator<TAccumulator, TModel>, new()
            where TModel : allows ref struct
        {
            CsvParallelOptions validated = Validated(options);
            return ParallelCsvProcessor.RunAsync(data, MappedAggregation<TAccumulator, TModel>.Unbound, Bound<TAccumulator, TModel>(map, validated), validated, ct);
        }

        /// <summary>
        /// Reads a CSV file across several threads, binds every record to a <typeparamref name="TModel"/>
        /// through <paramref name="map"/> and folds it into a <typeparamref name="TAccumulator"/>, one instance
        /// per partition, merged in file order.
        /// </summary>
        /// <typeparam name="TAccumulator">The accumulator type. See <see cref="ICsvAccumulator{TSelf, TModel}"/> for the contract it must honor.</typeparam>
        /// <typeparam name="TModel">The model type.</typeparam>
        /// <param name="path">The path of the CSV file to read.</param>
        /// <param name="map">How columns bind to <typeparamref name="TModel"/>.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>The accumulator holding every record.</returns>
        /// <exception cref="ArgumentException"><paramref name="map"/> binds columns by header name and <see cref="CsvParallelOptions.HeaderRow"/> is 0.</exception>
        /// <remarks>Carries the same contract and fallbacks as the in-memory overload.</remarks>
        public static Task<TAccumulator> AggregateCsvParallelAsync<TAccumulator, TModel>(
            string path,
            CsvModelMap<TModel> map,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
            where TAccumulator : ICsvAccumulator<TAccumulator, TModel>, new()
            where TModel : allows ref struct
        {
            ArgumentException.ThrowIfNullOrEmpty(path);
            CsvParallelOptions validated = Validated(options);
            return ParallelCsvProcessor.RunAsync(path, MappedAggregation<TAccumulator, TModel>.Unbound, Bound<TAccumulator, TModel>(map, validated), validated, ct);
        }

        /// <summary>
        /// Reads a CSV stream, in parallel where the stream can be partitioned, binds every record to a
        /// <typeparamref name="TModel"/> through <paramref name="map"/> and folds it into a
        /// <typeparamref name="TAccumulator"/>, one instance per partition, merged in stream order.
        /// </summary>
        /// <typeparam name="TAccumulator">The accumulator type. See <see cref="ICsvAccumulator{TSelf, TModel}"/> for the contract it must honor.</typeparam>
        /// <typeparam name="TModel">The model type.</typeparam>
        /// <param name="stream">The CSV stream, read from its current position. The caller keeps ownership and must not read from it concurrently.</param>
        /// <param name="map">How columns bind to <typeparamref name="TModel"/>.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>The accumulator holding every record.</returns>
        /// <exception cref="ArgumentException"><paramref name="map"/> binds columns by header name and <see cref="CsvParallelOptions.HeaderRow"/> is 0.</exception>
        /// <remarks>Carries the same contract and fallbacks as the in-memory overload, and the stream restrictions of the <see cref="ICsvRecord{TSelf}"/> stream overload.</remarks>
        public static Task<TAccumulator> AggregateCsvParallelAsync<TAccumulator, TModel>(
            Stream stream,
            CsvModelMap<TModel> map,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
            where TAccumulator : ICsvAccumulator<TAccumulator, TModel>, new()
            where TModel : allows ref struct
        {
            ArgumentNullException.ThrowIfNull(stream);
            CsvParallelOptions validated = Validated(options);
            return ParallelCsvProcessor.RunAsync(stream, MappedAggregation<TAccumulator, TModel>.Unbound, Bound<TAccumulator, TModel>(map, validated), validated, ct);
        }

        private static CsvAccumulateFactory<TAccumulator> Bound<TAccumulator, TModel>(CsvModelMap<TModel> map, CsvParallelOptions options)
            where TAccumulator : ICsvAccumulator<TAccumulator, TModel>, new()
            where TModel : allows ref struct
        {
            ArgumentNullException.ThrowIfNull(map);
            if (options.HeaderRow == 0 && !map.Info.IsIndexBased)
            {
                throw new ArgumentException("A map that binds columns by header name needs CsvParallelOptions.HeaderRow of at least 1.", nameof(map));
            }
            return MappedAggregation<TAccumulator, TModel>.Binder(map, options.HeaderRow);
        }

        private static CsvAccumulateFactory<byte> BoundCallback<TModel>(
            CsvModelMap<TModel> map, CsvRecordAction<TModel> body, CsvParallelOptions options)
            where TModel : allows ref struct
        {
            ArgumentNullException.ThrowIfNull(map);
            ArgumentNullException.ThrowIfNull(body);
            if (options.HeaderRow == 0 && !map.Info.IsIndexBased)
            {
                throw new ArgumentException("A map that binds columns by header name needs CsvParallelOptions.HeaderRow of at least 1.", nameof(map));
            }
            return MappedCallback<TModel>.Binder(map, options.HeaderRow, body);
        }

        /// <summary>
        /// Reads an in-memory CSV buffer across several threads and folds every record with the functions of
        /// <paramref name="aggregation"/>, one accumulator per partition, combined in buffer order.
        /// </summary>
        /// <typeparam name="TState">The accumulator type.</typeparam>
        /// <param name="data">The CSV bytes. The caller keeps ownership; the buffer must not be mutated during processing.</param>
        /// <param name="aggregation">The seed, accumulate and combine functions. They carry the same contract as <see cref="ICsvAccumulator{TSelf, TModel}"/>.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>The combined accumulator.</returns>
        /// <remarks>Carries the same fallbacks as the <see cref="ICsvAccumulator{TSelf, TModel}"/> overloads.</remarks>
        public static Task<TState> AggregateCsvParallelAsync<TState>(
            ReadOnlyMemory<byte> data,
            CsvAggregation<TState> aggregation,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
        {
            return ParallelCsvProcessor.RunAsync(data, Validated(aggregation), null, Validated(options), ct);
        }

        /// <summary>
        /// Reads a CSV file across several threads and folds every record with the functions of
        /// <paramref name="aggregation"/>, one accumulator per partition, combined in file order.
        /// </summary>
        /// <typeparam name="TState">The accumulator type.</typeparam>
        /// <param name="path">The path of the CSV file to read.</param>
        /// <param name="aggregation">The seed, accumulate and combine functions. They carry the same contract as <see cref="ICsvAccumulator{TSelf, TModel}"/>.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>The combined accumulator.</returns>
        /// <remarks>Carries the same fallbacks as the <see cref="ICsvAccumulator{TSelf, TModel}"/> overloads.</remarks>
        public static Task<TState> AggregateCsvParallelAsync<TState>(
            string path,
            CsvAggregation<TState> aggregation,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
        {
            ArgumentException.ThrowIfNullOrEmpty(path);
            return ParallelCsvProcessor.RunAsync(path, Validated(aggregation), null, Validated(options), ct);
        }

        /// <summary>
        /// Reads a CSV stream, in parallel where the stream can be partitioned, and folds every record with the
        /// functions of <paramref name="aggregation"/>, one accumulator per partition, combined in stream order.
        /// </summary>
        /// <typeparam name="TState">The accumulator type.</typeparam>
        /// <param name="stream">The CSV stream, read from its current position. The caller keeps ownership and must not read from it concurrently.</param>
        /// <param name="aggregation">The seed, accumulate and combine functions. They carry the same contract as <see cref="ICsvAccumulator{TSelf, TModel}"/>.</param>
        /// <param name="options">Parallelism, dialect and header options. Defaults to <see cref="CsvParallelOptions.Default"/>.</param>
        /// <param name="ct">A token to cancel processing.</param>
        /// <returns>The combined accumulator.</returns>
        /// <remarks>Carries the same fallbacks and stream restrictions as the <see cref="ICsvAccumulator{TSelf, TModel}"/> stream overload.</remarks>
        public static Task<TState> AggregateCsvParallelAsync<TState>(
            Stream stream,
            CsvAggregation<TState> aggregation,
            CsvParallelOptions? options = null,
            CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(stream);
            return ParallelCsvProcessor.RunAsync(stream, Validated(aggregation), null, Validated(options), ct);
        }

        private static CsvAggregation<TState> Validated<TState>(CsvAggregation<TState> aggregation)
        {
            ArgumentNullException.ThrowIfNull(aggregation);
            ArgumentNullException.ThrowIfNull(aggregation.Seed, nameof(aggregation));
            ArgumentNullException.ThrowIfNull(aggregation.Accumulate, nameof(aggregation));
            ArgumentNullException.ThrowIfNull(aggregation.Combine, nameof(aggregation));
            return aggregation;
        }

        private static CsvParallelOptions Validated(CsvParallelOptions? options)
        {
            options ??= CsvParallelOptions.Default;
            ArgumentOutOfRangeException.ThrowIfNegative(options.DegreeOfParallelism, nameof(options));
            ArgumentOutOfRangeException.ThrowIfNegative(options.HeaderRow, nameof(options));
            ArgumentNullException.ThrowIfNull(options.Reader, nameof(options));
            return options;
        }

        private static CsvParallelOptions ValidatedForTypedParse(CsvParallelOptions? options)
        {
            CsvParallelOptions validated = Validated(options);
            if (validated.HeaderRow != 0)
            {
                throw new ArgumentException(
                    $"{nameof(CsvParallelOptions)}.{nameof(CsvParallelOptions.HeaderRow)} does not apply to typed parsing; set {nameof(ExcelParserConfig)}.{nameof(ExcelParserConfig.HeaderRow)} instead.",
                    nameof(options));
            }
            return validated;
        }

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
                await zip.DisposeAsync().ConfigureAwait(false);
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
                    await zip.DisposeAsync().ConfigureAwait(false);
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
                    await zip.DisposeAsync().ConfigureAwait(false);
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

        private static ExcelFileFormat ClassifyZipStream(Stream stream, long start, out ZipArchive zip)
        {
            var zipPeek = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
            zip = zipPeek;
            bool isXlsb = zipPeek.GetEntry("xl/workbook.bin") is not null;
            stream.Position = start;
            return isXlsb ? ExcelFileFormat.Xlsb : ExcelFileFormat.Xlsx;
        }

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
            byte[] header = ArrayPool<byte>.Shared.Rent(8);
            ExcelFileFormat headerFormat;
            bool isCfb;
            try
            {
                int read = await stream.ReadAtLeastAsync(header.AsMemory(0, 8), 8, throwOnEndOfStream: false, ct).ConfigureAwait(false);
                stream.Position = start;
                ReadOnlySpan<byte> sig = header.AsSpan(0, read);
                if (TryClassifyHeader(sig, out headerFormat))
                {
                    return (headerFormat, null);
                }
                isCfb = sig.StartsWith(XlsCompoundFile.Signature);
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(header);
            }
            if (isCfb)
            {
                ExcelFileFormat cfbFormat = EncryptedPackageOpener.IsEncryptedContainer(stream)
                    ? ExcelFileFormat.EncryptedOoxml
                    : ExcelFileFormat.Xls;
                return (cfbFormat, null);
            }
            ExcelFileFormat zipFormat = ClassifyZipStream(stream, start, out ZipArchive zip);
            return (zipFormat, zip);
        }

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
