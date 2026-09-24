using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using ExcelReader.Core.Crypto;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader.Csv;
using ExcelReader.Core.Reader.Schema;
using ExcelReader.Core.Reader.Xls;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Core.Reader.Zip;

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
            FileStream stream = File.OpenRead(path);
            try
            {
                return new CsvReader(stream, leaveOpen: false, CsvDialectResolver.Resolve(stream, options ?? CsvReaderOptions.Default));
            }
            catch
            {
                stream.Dispose();
                throw;
            }
        }

        /// <summary>Opens a CSV (or other delimited-text) source from an existing stream.</summary>
        /// <param name="stream">The stream containing the CSV data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the reader is disposed.</param>
        /// <param name="options">Delimiter, quote, encoding, and size-limit settings; <see cref="CsvReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static CsvReader FromCsv(Stream stream, bool leaveOpen = true, CsvReaderOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(stream);
            CsvReaderOptions effective = CsvDialectResolver.Resolve(stream, options ?? CsvReaderOptions.Default);
            return new CsvReader(stream, leaveOpen, effective);
        }

        /// <summary>Opens a CSV (or other delimited-text) source directly from an in-memory buffer.</summary>
        /// <param name="data">The whole CSV source's bytes. Must outlive the returned reader and must not be mutated while it is in use.</param>
        /// <param name="options">Delimiter, quote, encoding, and size-limit settings; <see cref="CsvReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static CsvReader FromCsv(ReadOnlyMemory<byte> data, CsvReaderOptions? options = null)
        {
            CsvReaderOptions effective = CsvDialectResolver.Resolve(data, options ?? CsvReaderOptions.Default);
            return new CsvReader(data, effective);
        }

        /// <summary>Asynchronously opens a CSV (or other delimited-text) source from a file path, taking ownership of the file stream.</summary>
        /// <param name="path">The path to the CSV file.</param>
        /// <param name="options">Delimiter, quote, encoding, and size-limit settings; <see cref="CsvReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        public static async ValueTask<CsvReader> FromCsvFileAsync(string path, CsvReaderOptions? options = null, CancellationToken ct = default)
        {
            FileStream stream = OpenAsyncFile(path);
            try
            {
                CsvReaderOptions effective = await CsvDialectResolver.ResolveAsync(stream, options ?? CsvReaderOptions.Default, ct).ConfigureAwait(false);
                return await CsvReader.CreateAsync(stream, leaveOpen: false, effective, ct).ConfigureAwait(false);
            }
            catch
            {
                await stream.DisposeAsync().ConfigureAwait(false);
                throw;
            }
        }

        /// <summary>Asynchronously opens a CSV (or other delimited-text) source from an existing stream.</summary>
        /// <param name="stream">The stream containing the CSV data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the reader is disposed.</param>
        /// <param name="options">Delimiter, quote, encoding, and size-limit settings; <see cref="CsvReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        public static async ValueTask<CsvReader> FromCsvAsync(Stream stream, bool leaveOpen = true, CsvReaderOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(stream);
            CsvReaderOptions effective = await CsvDialectResolver.ResolveAsync(stream, options ?? CsvReaderOptions.Default, ct).ConfigureAwait(false);
            return await CsvReader.CreateAsync(stream, leaveOpen, effective, ct).ConfigureAwait(false);
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
        /// before trusting it, and feed the result into <see cref="ExcelParser.Build{T}"/> to
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

        /// <summary>
        /// Opens a workbook of a known format from a file path, taking ownership of the file stream.
        /// </summary>
        /// <param name="path">The path to the workbook file.</param>
        /// <param name="format">The format to open <paramref name="path"/> as. <see cref="ExcelFileFormat.Unknown"/>
        /// auto-detects XLSX/XLSB/XLS from the signature, exactly as <see cref="Open(string,ExcelReaderOptions?)"/> does.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when
        /// <see langword="null"/>. <see cref="ExcelReaderOptions.Csv"/> supplies the dialect when
        /// <paramref name="format"/> is <see cref="ExcelFileFormat.Csv"/>.</param>
        /// <returns>A format-agnostic <see cref="IExcelRowReader"/>.</returns>
        /// <remarks>CSV carries no signature, so auto-detection never reports it; opening delimited text
        /// means naming <see cref="ExcelFileFormat.Csv"/> here.</remarks>
        /// <exception cref="ArgumentNullException"><paramref name="path"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="format"/> is <see cref="ExcelFileFormat.EncryptedOoxml"/> or not a defined value.</exception>
        /// <exception cref="InvalidDataException"><paramref name="format"/> is <see cref="ExcelFileFormat.Unknown"/> and the file's signature matches no supported format.</exception>
        public static IExcelRowReader Open(string path, ExcelFileFormat format, ExcelReaderOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(path);
            RequireOpenableFormat(format);
            return format switch
            {
                ExcelFileFormat.Csv => FromCsvFile(path, CsvOf(options)),
                ExcelFileFormat.Xlsx => FromXlsxFile(path, options),
                ExcelFileFormat.Xlsb => FromXlsbFile(path, options),
                ExcelFileFormat.Xls => FromXlsFile(path, options),
                _ => Open(path, options),
            };
        }

        /// <summary>
        /// Opens a workbook of a known format from an existing stream.
        /// </summary>
        /// <param name="stream">The stream containing the workbook data. Must be seekable when
        /// <paramref name="format"/> is <see cref="ExcelFileFormat.Unknown"/>.</param>
        /// <param name="format">The format to read <paramref name="stream"/> as. <see cref="ExcelFileFormat.Unknown"/>
        /// auto-detects XLSX/XLSB/XLS from the signature.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the reader is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when
        /// <see langword="null"/>. <see cref="ExcelReaderOptions.Csv"/> supplies the dialect when
        /// <paramref name="format"/> is <see cref="ExcelFileFormat.Csv"/>.</param>
        /// <returns>A format-agnostic <see cref="IExcelRowReader"/>.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="stream"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="format"/> is <see cref="ExcelFileFormat.EncryptedOoxml"/> or not a defined value.</exception>
        /// <exception cref="ArgumentException"><paramref name="format"/> is <see cref="ExcelFileFormat.Unknown"/> and <paramref name="stream"/> does not support seeking.</exception>
        /// <exception cref="InvalidDataException"><paramref name="format"/> is <see cref="ExcelFileFormat.Unknown"/> and the stream's signature matches no supported format.</exception>
        public static IExcelRowReader Open(Stream stream, ExcelFileFormat format, bool leaveOpen = true, ExcelReaderOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(stream);
            RequireOpenableFormat(format);
            return format switch
            {
                ExcelFileFormat.Csv => FromCsv(stream, leaveOpen, CsvOf(options)),
                ExcelFileFormat.Xlsx => FromXlsx(stream, leaveOpen, options),
                ExcelFileFormat.Xlsb => FromXlsb(stream, leaveOpen, options),
                ExcelFileFormat.Xls => FromXls(stream, leaveOpen, options),
                _ => Open(stream, leaveOpen, options),
            };
        }

        /// <summary>
        /// Opens a workbook of a known format from an in-memory buffer.
        /// </summary>
        /// <param name="data">The whole source's bytes. Must outlive the returned reader.</param>
        /// <param name="format">The format to read <paramref name="data"/> as. <see cref="ExcelFileFormat.Unknown"/>
        /// auto-detects XLSX/XLSB/XLS from the signature.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when
        /// <see langword="null"/>. <see cref="ExcelReaderOptions.Csv"/> supplies the dialect when
        /// <paramref name="format"/> is <see cref="ExcelFileFormat.Csv"/>.</param>
        /// <returns>A format-agnostic <see cref="IExcelRowReader"/>.</returns>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="format"/> is <see cref="ExcelFileFormat.EncryptedOoxml"/> or not a defined value.</exception>
        /// <exception cref="InvalidDataException"><paramref name="format"/> is <see cref="ExcelFileFormat.Unknown"/> and the buffer's signature matches no supported format.</exception>
        public static IExcelRowReader Open(ReadOnlyMemory<byte> data, ExcelFileFormat format, ExcelReaderOptions? options = null)
        {
            RequireOpenableFormat(format);
            return format switch
            {
                ExcelFileFormat.Csv => FromCsv(data, CsvOf(options)),
                ExcelFileFormat.Xlsx => FromXlsx(data, options),
                ExcelFileFormat.Xlsb => FromXlsb(data, options),
                ExcelFileFormat.Xls => FromXls(data, options),
                _ => Open(data, options),
            };
        }

        /// <summary>
        /// Asynchronously opens a workbook of a known format from a file path, taking ownership of the file stream.
        /// </summary>
        /// <param name="path">The path to the workbook file.</param>
        /// <param name="format">The format to open <paramref name="path"/> as. <see cref="ExcelFileFormat.Unknown"/>
        /// auto-detects XLSX/XLSB/XLS from the signature.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when
        /// <see langword="null"/>. <see cref="ExcelReaderOptions.Csv"/> supplies the dialect when
        /// <paramref name="format"/> is <see cref="ExcelFileFormat.Csv"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        /// <returns>A format-agnostic <see cref="IExcelRowReader"/>.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="path"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="format"/> is <see cref="ExcelFileFormat.EncryptedOoxml"/> or not a defined value.</exception>
        /// <exception cref="InvalidDataException"><paramref name="format"/> is <see cref="ExcelFileFormat.Unknown"/> and the file's signature matches no supported format.</exception>
        public static ValueTask<IExcelRowReader> OpenAsync(string path, ExcelFileFormat format, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(path);
            RequireOpenableFormat(format);
            return format switch
            {
                ExcelFileFormat.Csv => AsRowReaderAsync(FromCsvFileAsync(path, CsvOf(options), ct)),
                ExcelFileFormat.Xlsx => AsRowReaderAsync(FromXlsxFileAsync(path, options, ct)),
                ExcelFileFormat.Xlsb => AsRowReaderAsync(FromXlsbFileAsync(path, options, ct)),
                ExcelFileFormat.Xls => AsRowReaderAsync(FromXlsFileAsync(path, options, ct)),
                _ => OpenAsync(path, options, ct),
            };
        }

        /// <summary>
        /// Asynchronously opens a workbook of a known format from an existing stream.
        /// </summary>
        /// <param name="stream">The stream containing the workbook data. Must be seekable when
        /// <paramref name="format"/> is <see cref="ExcelFileFormat.Unknown"/>.</param>
        /// <param name="format">The format to read <paramref name="stream"/> as. <see cref="ExcelFileFormat.Unknown"/>
        /// auto-detects XLSX/XLSB/XLS from the signature.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the reader is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when
        /// <see langword="null"/>. <see cref="ExcelReaderOptions.Csv"/> supplies the dialect when
        /// <paramref name="format"/> is <see cref="ExcelFileFormat.Csv"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        /// <returns>A format-agnostic <see cref="IExcelRowReader"/>.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="stream"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="format"/> is <see cref="ExcelFileFormat.EncryptedOoxml"/> or not a defined value.</exception>
        /// <exception cref="ArgumentException"><paramref name="format"/> is <see cref="ExcelFileFormat.Unknown"/> and <paramref name="stream"/> does not support seeking.</exception>
        /// <exception cref="InvalidDataException"><paramref name="format"/> is <see cref="ExcelFileFormat.Unknown"/> and the stream's signature matches no supported format.</exception>
        public static ValueTask<IExcelRowReader> OpenAsync(Stream stream, ExcelFileFormat format, bool leaveOpen = true, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(stream);
            RequireOpenableFormat(format);
            return format switch
            {
                ExcelFileFormat.Csv => AsRowReaderAsync(FromCsvAsync(stream, leaveOpen, CsvOf(options), ct)),
                ExcelFileFormat.Xlsx => AsRowReaderAsync(FromXlsxAsync(stream, leaveOpen, options, ct)),
                ExcelFileFormat.Xlsb => AsRowReaderAsync(FromXlsbAsync(stream, leaveOpen, options, ct)),
                ExcelFileFormat.Xls => AsRowReaderAsync(FromXlsAsync(stream, leaveOpen, options, ct)),
                _ => OpenAsync(stream, leaveOpen, options, ct),
            };
        }

        private static CsvReaderOptions CsvOf(ExcelReaderOptions? options)
        {
            return options?.Csv ?? CsvReaderOptions.Default;
        }

        private static async ValueTask<IExcelRowReader> AsRowReaderAsync<TReader>(ValueTask<TReader> opening)
            where TReader : IExcelRowReader
        {
            return await opening.ConfigureAwait(false);
        }

        private static void RequireOpenableFormat(ExcelFileFormat format)
        {
            if (format is ExcelFileFormat.EncryptedOoxml)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(format),
                    format,
                    "EncryptedOoxml is a detection result, not a format to open; pass Unknown, Xlsx or Xlsb with ExcelReaderOptions.Password set.");
            }
            if (format is < ExcelFileFormat.Unknown or > ExcelFileFormat.Csv)
            {
                throw new ArgumentOutOfRangeException(nameof(format), format, "Not a defined ExcelFileFormat value.");
            }
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

        internal static FileStream OpenAsyncFile(string path)
        {
            return new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 65536,
                                  options: FileOptions.Asynchronous | FileOptions.SequentialScan);
        }
    }
}
