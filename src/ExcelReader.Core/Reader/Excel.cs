using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Internal;

namespace ExcelReader.Core.Reader
{
    /// <summary>
    /// Entry point for opening Excel and CSV workbooks. <see cref="Open(string,ExcelReaderOptions?)"/>/
    /// <see cref="Open(Stream,bool,ExcelReaderOptions?)"/> and their async counterparts auto-detect XLSX/XLSB/XLS
    /// from the file's signature and return a format-agnostic <see cref="IExcelRowReader"/>; the
    /// <c>From*</c>/<c>FromXls*</c>/<c>FromXlsb*</c>/<c>FromCsv*</c> methods open a specific, known format directly.
    /// </summary>
    public static class Excel
    {
        /// <summary>Opens an XLSX workbook from a file path, taking ownership of the file stream.</summary>
        /// <param name="path">The path to the XLSX file.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsxReader FromFile(string path, ExcelReaderOptions? options = null)
        {
            return new XlsxReader(File.OpenRead(path), leaveOpen: false, options);
        }

        /// <summary>Opens an XLSX workbook from an existing stream.</summary>
        /// <param name="stream">The stream containing the XLSX data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the reader is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsxReader From(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null)
        {
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
            return XlsxReader.CreateFromMemory(data, options);
        }

        /// <summary>Opens an XLSX workbook from a file path, taking ownership of the file stream. Alias for <see cref="FromFile(string, ExcelReaderOptions?)"/>, for callers who grep for a format-named factory.</summary>
        /// <param name="path">The path to the XLSX file.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsxReader FromXlsxFile(string path, ExcelReaderOptions? options = null)
        {
            return FromFile(path, options);
        }

        /// <summary>Opens an XLSX workbook from an existing stream. Alias for <see cref="From(Stream, bool, ExcelReaderOptions?)"/>, for callers who grep for a format-named factory.</summary>
        /// <param name="stream">The stream containing the XLSX data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the reader is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsxReader FromXlsx(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null)
        {
            return From(stream, leaveOpen, options);
        }

        /// <summary>Opens an XLSX workbook directly from an in-memory buffer. Alias for <see cref="From(ReadOnlyMemory{byte}, ExcelReaderOptions?)"/>, for callers who grep for a format-named factory.</summary>
        /// <param name="data">The whole XLSX file's bytes. Must outlive the returned reader.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsxReader FromXlsx(ReadOnlyMemory<byte> data, ExcelReaderOptions? options = null)
        {
            return From(data, options);
        }

        /// <summary>Opens a legacy binary (XLS) workbook from a file path, taking ownership of the file stream.</summary>
        /// <param name="path">The path to the XLS file.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Stream ownership transfers to XlsReader, which streams from it and disposes it on Dispose (and on construction failure).")]
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
            return new XlsbReader(File.OpenRead(path), leaveOpen: false, options);
        }

        /// <summary>Opens an XLSB (Excel binary) workbook from an existing stream.</summary>
        /// <param name="stream">The stream containing the XLSB data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the reader is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsbReader FromXlsb(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null)
        {
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
            return XlsbReader.CreateFromMemory(data, options);
        }

        /// <summary>Asynchronously opens an XLSX workbook from a file path, taking ownership of the file stream.</summary>
        /// <param name="path">The path to the XLSX file.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream ownership transfers to CreateAsync, which disposes it on failure and via the reader on success.")]
        public static ValueTask<XlsxReader> FromFileAsync(string path, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            FileStream stream = OpenAsyncFile(path);
            return XlsxReader.CreateAsync(stream, leaveOpen: false, options, ct);
        }

        /// <summary>Asynchronously opens an XLSX workbook from an existing stream.</summary>
        /// <param name="stream">The stream containing the XLSX data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the reader is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        public static ValueTask<XlsxReader> FromAsync(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            return XlsxReader.CreateAsync(stream, leaveOpen, options, ct);
        }

        /// <summary>Asynchronously opens an XLSX workbook from a file path, taking ownership of the file stream. Alias for <see cref="FromFileAsync(string, ExcelReaderOptions?, CancellationToken)"/>, for callers who grep for a format-named factory.</summary>
        /// <param name="path">The path to the XLSX file.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        public static ValueTask<XlsxReader> FromXlsxFileAsync(string path, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            return FromFileAsync(path, options, ct);
        }

        /// <summary>Asynchronously opens an XLSX workbook from an existing stream. Alias for <see cref="FromAsync(Stream, bool, ExcelReaderOptions?, CancellationToken)"/>, for callers who grep for a format-named factory.</summary>
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
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Stream ownership transfers to CreateAsync, which disposes it on failure and via the reader on success.")]
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
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream ownership transfers to CreateAsync, which disposes it on failure and via the reader on success.")]
        public static ValueTask<XlsbReader> FromXlsbFileAsync(string path, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            FileStream stream = OpenAsyncFile(path);
            return XlsbReader.CreateAsync(stream, leaveOpen: false, options, ct);
        }

        /// <summary>Asynchronously opens an XLSB (Excel binary) workbook from an existing stream.</summary>
        /// <param name="stream">The stream containing the XLSB data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the reader is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        public static ValueTask<XlsbReader> FromXlsbAsync(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
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
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream ownership transfers to CreateAsync, which disposes it on failure and via the reader on success.")]
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

        // The leading bytes that distinguish container formats: XLSX and XLSB are ZIP ("PK\x03\x04"),
        // XLS is an OLE2/CFB compound document. XLSB is distinguished from XLSX by the presence of
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
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Stream ownership transfers to OpenSeekable, which disposes it on failure and via the reader on success.")]
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
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "memZip (when non-null) is handed to the chosen reader on success, which takes ownership; on Unknown format it's disposed immediately above.")]
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "memZip (when non-null) is handed to the chosen reader on success, which takes ownership; on Unknown format it's disposed immediately above.")]
        public static IExcelRowReader Open(ReadOnlyMemory<byte> data, ExcelReaderOptions? options = null)
        {
            ExcelReaderOptions effective = options ?? ExcelReaderOptions.Default;
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

        // The ZIP central directory must be walked to tell XLSB from XLSX (same "peek to distinguish"
        // shape as DetectSeekable for streams), so the resulting ZipMemoryIndex is handed back for reuse
        // by the chosen reader's CreateFromMemory instead of being parsed a second time.
        private static ExcelFileFormat ClassifyMemory(ReadOnlyMemory<byte> data, ExcelReaderOptions options, out ZipMemoryIndex? memZip)
        {
            memZip = null;
            ReadOnlySpan<byte> span = data.Span;
            ReadOnlySpan<byte> header = span.Length > 8 ? span[..8] : span;
            if (TryClassifyHeader(header, out ExcelFileFormat format))
            {
                return format;
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
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Stream ownership transfers to OpenSeekableAsync, which disposes it on failure and via the reader on success.")]
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

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "zip (when non-null) is handed to the chosen reader on success, which takes ownership; on failure it's disposed in the catch below.")]
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "zip (when non-null) is handed to the chosen reader on success, which takes ownership; on failure it's disposed in the catch below.")]
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
            // zip (when non-null) is handed to the chosen reader, which takes ownership of it — this
            // is the same archive DetectSeekable already opened to peek "xl/workbook.bin", so the
            // central directory isn't parsed a second time.
            return format switch
            {
                ExcelFileFormat.Xls => new XlsReader(stream, leaveOpen, options),
                ExcelFileFormat.Xlsb => new XlsbReader(stream, leaveOpen, zip!, options),
                ExcelFileFormat.Xlsx => new XlsxReader(stream, leaveOpen, zip!, options),
                _ => throw new System.Diagnostics.UnreachableException(),
            };
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
            return format switch
            {
                ExcelFileFormat.Xls => await XlsReader.CreateAsync(stream, leaveOpen, options, ct).ConfigureAwait(false),
                ExcelFileFormat.Xlsb => await XlsbReader.CreateFromOpenZipAsync(stream, leaveOpen, zip!, options, ct).ConfigureAwait(false),
                ExcelFileFormat.Xlsx => await XlsxReader.CreateFromOpenZipAsync(stream, leaveOpen, zip!, options, ct).ConfigureAwait(false),
                _ => throw new System.Diagnostics.UnreachableException(),
            };
        }

        private static void DisposeOnFailure(Stream stream, bool leaveOpen)
        {
            if (!leaveOpen)
            {
                stream.Dispose();
            }
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "Open and OpenAsync transfer ownership when leaveOpen is false.")]
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

        // Detection peeks the 8-byte signature then rewinds. For ZIP streams, opens a temporary
        // ZipArchive to distinguish XLSB ("xl/workbook.bin" present) from XLSX (XML workbook).
        // Both XLSX and XLSB readers need a seekable source anyway (ZipArchive seeks the central
        // directory), so requiring seek here costs nothing and keeps the peek cheap.
        // Classifies the leading signature bytes shared by DetectSeekable/DetectSeekableAsync. Returns
        // true (with the final answer) for XLS/Unknown; false means "it's a ZIP" and the caller must
        // still peek the central directory to tell XLSB from XLSX - the one step that genuinely
        // differs between the sync (stackalloc) and async (heap buffer, awaited zip dispose) paths.
        private static bool TryClassifyHeader(ReadOnlySpan<byte> sig, out ExcelFileFormat format)
        {
            if (sig.StartsWith(XlsCompoundFile.Signature))
            {
                format = ExcelFileFormat.Xls;
                return true;
            }
            if (!sig.StartsWith(ZipSignature))
            {
                format = ExcelFileFormat.Unknown;
                return true;
            }
            format = default;
            return false;
        }

        // `zip` receives the archive opened to peek the central directory (null for Xls/Unknown, which
        // never need one) so the caller can hand it straight to the chosen reader instead of letting
        // that reader re-parse the central directory from scratch.
        [SkipLocalsInit]
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Ownership transfers to the caller via the out parameter, which disposes it on failure or hands it to the chosen reader on success.")]
        private static ExcelFileFormat DetectSeekable(Stream stream, out ZipArchive? zip)
        {
            zip = null;
            RequireSeekable(stream);
            long start = stream.Position;
            Span<byte> header = stackalloc byte[8];
            int read = stream.ReadAtLeast(header, header.Length, throwOnEndOfStream: false);
            stream.Position = start;
            if (TryClassifyHeader(header[..read], out ExcelFileFormat format))
            {
                return format;
            }
            // Peek the central directory to distinguish XLSB from XLSX. Assign `zip` immediately so a
            // caller-side catch can dispose it even if GetEntry below were to throw.
            var zipPeek = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
            zip = zipPeek;
            bool isXlsb = zipPeek.GetEntry("xl/workbook.bin") is not null;
            stream.Position = start;
            return isXlsb ? ExcelFileFormat.Xlsb : ExcelFileFormat.Xlsx;
        }

        private static async ValueTask<(ExcelFileFormat Format, ZipArchive? Zip)> DetectSeekableAsync(Stream stream, CancellationToken ct)
        {
            RequireSeekable(stream);
            long start = stream.Position;
            byte[] header = new byte[8];
            int read = await stream.ReadAtLeastAsync(header, header.Length, throwOnEndOfStream: false, ct).ConfigureAwait(false);
            stream.Position = start;
            if (TryClassifyHeader(header.AsSpan(0, read), out ExcelFileFormat format))
            {
                return (format, null);
            }
            // Peek the central directory to distinguish XLSB from XLSX; kept open (not disposed here)
            // so the caller can hand it straight to the chosen reader.
            var zipPeek = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
            bool isXlsb = zipPeek.GetEntry("xl/workbook.bin") is not null;
            stream.Position = start;
            return (isXlsb ? ExcelFileFormat.Xlsb : ExcelFileFormat.Xlsx, zipPeek);
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
