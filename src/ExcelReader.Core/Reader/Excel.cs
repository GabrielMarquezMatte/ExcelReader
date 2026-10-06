using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using ExcelReader.Core.Crypto;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader.Csv;
using ExcelReader.Core.Reader.Schema;
using ExcelReader.Core.Reader.Sources;
using ExcelReader.Core.Reader.Xls;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Core.Reader.Zip;

namespace ExcelReader.Core.Reader
{
    /// <summary>
    /// Entry point for opening Excel and CSV workbooks. <see cref="Open(string,ExcelReaderOptions?)"/>/
    /// <see cref="Open(Stream,bool,ExcelReaderOptions?)"/> and their async counterparts auto-detect XLSX/XLSB/XLS
    /// from the file's signature and return a format-agnostic <see cref="IExcelWorkbook"/>; the
    /// <c>From*</c>/<c>FromXls*</c>/<c>FromXlsb*</c>/<c>FromCsv*</c> methods open a specific, known format directly.
    /// </summary>
    public static partial class Excel
    {
        /// <summary>Opens an XLSX workbook from a file path, taking ownership of the file stream.</summary>
        /// <param name="path">The path to the XLSX file.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsxWorkbook FromXlsxFile(string path, ExcelReaderOptions? options = null)
        {
            return FromXlsx(File.OpenRead(path), leaveOpen: false, options);
        }

        /// <summary>Opens an XLSX workbook from an existing stream.</summary>
        /// <param name="stream">The stream containing the XLSX data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the workbook is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsxWorkbook FromXlsx(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null)
        {
            if (TryDecryptCfbStream(stream, leaveOpen, options, out Stream decrypted))
            {
                return new XlsxWorkbook(decrypted, leaveOpen: false, options);
            }
            return new XlsxWorkbook(stream, leaveOpen, options);
        }

        /// <summary>
        /// Opens an XLSX workbook directly from an in-memory buffer. Reads the ZIP
        /// central directory and decompresses parts without a <c>ZipArchive</c>
        /// or intermediate <see cref="Stream"/> — every part is fully materialized up front, so the returned
        /// reader never suspends, even under <c>await foreach</c>.
        /// </summary>
        /// <param name="data">The whole XLSX file's bytes. Must outlive the returned reader.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>. <see cref="ExcelReaderOptions.PrefetchDecompression"/> is ignored on this path — there is nothing left to overlap.</param>
        public static XlsxWorkbook FromXlsx(ReadOnlyMemory<byte> data, ExcelReaderOptions? options = null)
        {
            ExcelReaderOptions effective = options ?? ExcelReaderOptions.Default;
            if (data.Span.StartsWith(XlsCompoundFile.Signature) && EncryptedPackageOpener.IsEncryptedMemory(data, effective))
            {
                ReadOnlyMemory<byte> plain = EncryptedPackageOpener.DecryptToMemory(data, effective);
                return XlsxWorkbook.CreateFromMemory(plain, effective);
            }
            return XlsxWorkbook.CreateFromMemory(data, effective);
        }

        /// <summary>Opens a legacy binary (XLS) workbook from a file path, taking ownership of the file stream.</summary>
        /// <param name="path">The path to the XLS file.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsWorkbook FromXlsFile(string path, ExcelReaderOptions? options = null)
        {
            return new XlsWorkbook(File.OpenRead(path), leaveOpen: false, options);
        }

        /// <summary>Opens a legacy binary (XLS) workbook from an existing stream.</summary>
        /// <param name="stream">The stream containing the XLS data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the workbook is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsWorkbook FromXls(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null)
        {
            return new XlsWorkbook(stream, leaveOpen, options);
        }

        /// <summary>Opens a legacy binary (XLS) workbook directly from an in-memory buffer.</summary>
        /// <param name="data">The whole XLS file's bytes. Must outlive the returned reader and must not be mutated while it is in use.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsWorkbook FromXls(ReadOnlyMemory<byte> data, ExcelReaderOptions? options = null)
        {
            return new XlsWorkbook(data, options);
        }

        /// <summary>Opens an XLSB (Excel binary) workbook from a file path, taking ownership of the file stream.</summary>
        /// <param name="path">The path to the XLSB file.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsbWorkbook FromXlsbFile(string path, ExcelReaderOptions? options = null)
        {
            return FromXlsb(File.OpenRead(path), leaveOpen: false, options);
        }

        /// <summary>Opens an XLSB (Excel binary) workbook from an existing stream.</summary>
        /// <param name="stream">The stream containing the XLSB data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the workbook is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        public static XlsbWorkbook FromXlsb(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null)
        {
            if (TryDecryptCfbStream(stream, leaveOpen, options, out Stream decrypted))
            {
                return new XlsbWorkbook(decrypted, leaveOpen: false, options);
            }
            return new XlsbWorkbook(stream, leaveOpen, options);
        }

        /// <summary>
        /// Opens an XLSB workbook directly from an in-memory buffer. Reads the ZIP
        /// central directory and decompresses parts without a <c>ZipArchive</c>
        /// or intermediate <see cref="Stream"/> — every part is fully materialized up front, so the returned
        /// reader never suspends, even under <c>await foreach</c>.
        /// </summary>
        /// <param name="data">The whole XLSB file's bytes. Must outlive the returned reader.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>. <see cref="ExcelReaderOptions.PrefetchDecompression"/> is ignored on this path — there is nothing left to overlap.</param>
        public static XlsbWorkbook FromXlsb(ReadOnlyMemory<byte> data, ExcelReaderOptions? options = null)
        {
            ExcelReaderOptions effective = options ?? ExcelReaderOptions.Default;
            if (data.Span.StartsWith(XlsCompoundFile.Signature) && EncryptedPackageOpener.IsEncryptedMemory(data, effective))
            {
                ReadOnlyMemory<byte> plain = EncryptedPackageOpener.DecryptToMemory(data, effective);
                return XlsbWorkbook.CreateFromMemory(plain, effective);
            }
            return XlsbWorkbook.CreateFromMemory(data, effective);
        }

        /// <summary>Asynchronously opens an XLSX workbook from a file path, taking ownership of the file stream.</summary>
        /// <param name="path">The path to the XLSX file.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        public static ValueTask<XlsxWorkbook> FromXlsxFileAsync(string path, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            FileStream stream = OpenAsyncFile(path);
            return FromXlsxAsync(stream, leaveOpen: false, options, ct);
        }

        /// <summary>Asynchronously opens an XLSX workbook from an existing stream.</summary>
        /// <param name="stream">The stream containing the XLSX data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the workbook is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        public static ValueTask<XlsxWorkbook> FromXlsxAsync(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            if (TryDecryptCfbStream(stream, leaveOpen, options, out Stream decrypted))
            {
                return XlsxWorkbook.CreateAsync(decrypted, leaveOpen: false, options, ct);
            }
            return XlsxWorkbook.CreateAsync(stream, leaveOpen, options, ct);
        }

        /// <summary>Asynchronously opens a legacy binary (XLS) workbook from a file path, taking ownership of the file stream.</summary>
        /// <param name="path">The path to the XLS file.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        public static ValueTask<XlsWorkbook> FromXlsFileAsync(string path, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            FileStream stream = OpenAsyncFile(path);
            return XlsWorkbook.CreateAsync(stream, leaveOpen: false, options, ct);
        }

        /// <summary>Asynchronously opens a legacy binary (XLS) workbook from an existing stream.</summary>
        /// <param name="stream">The stream containing the XLS data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the workbook is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        public static ValueTask<XlsWorkbook> FromXlsAsync(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            return XlsWorkbook.CreateAsync(stream, leaveOpen, options, ct);
        }

        /// <summary>Asynchronously opens an XLSB (Excel binary) workbook from a file path, taking ownership of the file stream.</summary>
        /// <param name="path">The path to the XLSB file.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        public static ValueTask<XlsbWorkbook> FromXlsbFileAsync(string path, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            FileStream stream = OpenAsyncFile(path);
            return FromXlsbAsync(stream, leaveOpen: false, options, ct);
        }

        /// <summary>Asynchronously opens an XLSB (Excel binary) workbook from an existing stream.</summary>
        /// <param name="stream">The stream containing the XLSB data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the workbook is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        public static ValueTask<XlsbWorkbook> FromXlsbAsync(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            if (TryDecryptCfbStream(stream, leaveOpen, options, out Stream decrypted))
            {
                return XlsbWorkbook.CreateAsync(decrypted, leaveOpen: false, options, ct);
            }
            return XlsbWorkbook.CreateAsync(stream, leaveOpen, options, ct);
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
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the workbook is disposed.</param>
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
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the workbook is disposed.</param>
        /// <param name="options">Delimiter, quote, encoding, and size-limit settings; <see cref="CsvReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        public static async ValueTask<CsvReader> FromCsvAsync(Stream stream, bool leaveOpen = true, CsvReaderOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(stream);
            CsvReaderOptions effective = await CsvDialectResolver.ResolveAsync(stream, options ?? CsvReaderOptions.Default, ct).ConfigureAwait(false);
            return await CsvReader.CreateAsync(stream, leaveOpen, effective, ct).ConfigureAwait(false);
        }

        /// <summary>Guesses a column schema for <paramref name="sheet"/> by sampling it from the first row.</summary>
        /// <param name="sheet">The sheet to sample. A new enumerator is opened and disposed; no other read of the sheet is disturbed.</param>
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
        /// <exception cref="ArgumentNullException"><paramref name="sheet"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="headerRow"/> is negative, or
        /// <paramref name="sampleSize"/> is not positive.</exception>
        /// <exception cref="ArgumentException">The sheet has fewer rows than <paramref name="headerRow"/>.</exception>
        public static ExcelColumnSchema[] InferSchema(IExcelSheet sheet, int headerRow, int sampleSize)
        {
            return InferSchema(sheet, headerRow, sampleSize, parseText: false);
        }

        /// <summary>
        /// Guesses a column schema by sampling the sheet's rows, as
        /// <see cref="InferSchema(IExcelSheet, int, int)"/> does, and can also type cells that hold text.
        /// </summary>
        /// <param name="sheet">The sheet to sample. A new enumerator is opened and disposed; no other read of the sheet is disturbed.</param>
        /// <param name="headerRow">1-based row number to take column names from; 0 means "no header",
        /// so every returned schema is addressable only by <see cref="ExcelColumnSchema.Index"/>.</param>
        /// <param name="sampleSize">How many rows after the header to inspect.</param>
        /// <param name="parseText">When <see langword="true"/>, a text cell (every CSV field, or a number
        /// stored as text) counts as an integer, a decimal, <c>true</c>/<c>false</c>, or an ISO-8601 date or
        /// date-time when its text has exactly that shape. Numbers with a leading zero (<c>00123</c>), scientific
        /// notation (<c>12E4</c>), padded or culture-formatted numbers, and non-ISO dates stay text. A column only gets a non-text type when
        /// every sampled value converts to it.</param>
        /// <returns>One <see cref="ExcelColumnSchema"/> per column, in column order.</returns>
        /// <remarks>Still a guess over a bounded sample: a value past the sample can still fail to convert.
        /// See <see cref="InferSchema(IExcelSheet, int, int)"/>.</remarks>
        /// <exception cref="ArgumentNullException"><paramref name="sheet"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="headerRow"/> is negative, or
        /// <paramref name="sampleSize"/> is not positive.</exception>
        /// <exception cref="ArgumentException">The sheet has fewer rows than <paramref name="headerRow"/>.</exception>
        public static ExcelColumnSchema[] InferSchema(IExcelSheet sheet, int headerRow = 1, int sampleSize = 100, bool parseText = false)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            using IExcelRowEnumerator rows = sheet.GetEnumerator();
            return SchemaInference.Infer(rows, sheet.IsDate1904, headerRow, sampleSize, parseText);
        }

        private static ReadOnlySpan<byte> ZipSignature => [0x50, 0x4B, 0x03, 0x04];

        /// <summary>
        /// Opens a workbook from a file path, auto-detecting its format (XLSX/XLSB/XLS) from the file's signature
        /// and taking ownership of the file stream.
        /// </summary>
        /// <remarks>Pattern-match on the returned workbook against its concrete type (<see cref="XlsxWorkbook"/> / <see cref="XlsWorkbook"/> / <see cref="XlsbWorkbook"/>) to access format-specific members.</remarks>
        /// <param name="path">The path to the workbook file.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <returns>A format-agnostic <see cref="IExcelWorkbook"/> backed by the concrete workbook that matches the detected format.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="path"/> is <see langword="null"/>.</exception>
        /// <exception cref="InvalidDataException">The file's signature does not match a supported format.</exception>
        public static IExcelWorkbook Open(string path, ExcelReaderOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(path);
            return OpenSeekable(File.OpenRead(path), leaveOpen: false, options);
        }

        /// <summary>
        /// Opens a workbook from an existing seekable stream, auto-detecting its format (XLSX/XLSB/XLS) from the
        /// stream's signature.
        /// </summary>
        /// <param name="stream">A seekable stream containing the workbook data.</param>
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the workbook is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <returns>A format-agnostic <see cref="IExcelWorkbook"/> backed by the concrete workbook that matches the detected format.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="stream"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="stream"/> does not support seeking.</exception>
        /// <exception cref="InvalidDataException">The stream's signature does not match a supported format.</exception>
        public static IExcelWorkbook Open(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(stream);
            return OpenSeekable(stream, leaveOpen, options);
        }

        /// <summary>
        /// Opens a workbook from an in-memory buffer, auto-detecting its format (XLSX/XLSB/XLS) from its
        /// signature. XLSX/XLSB route through <see cref="ZipIndex"/> instead of
        /// a <c>ZipArchive</c>/<see cref="Stream"/>, so the returned reader never
        /// suspends, even under <c>await foreach</c>.
        /// </summary>
        /// <param name="data">The whole workbook file's bytes. Must outlive the returned reader.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <returns>A format-agnostic <see cref="IExcelWorkbook"/> backed by the concrete workbook that matches the detected format.</returns>
        /// <exception cref="InvalidDataException">The buffer's signature does not match a supported format.</exception>
        public static IExcelWorkbook Open(ReadOnlyMemory<byte> data, ExcelReaderOptions? options = null)
        {
            ExcelReaderOptions effective = options ?? ExcelReaderOptions.Default;
            if (data.Span.StartsWith(XlsCompoundFile.Signature) && EncryptedPackageOpener.IsEncryptedMemory(data, effective))
            {
                ReadOnlyMemory<byte> plain = EncryptedPackageOpener.DecryptToMemory(data, effective);
                return OpenFromPlainMemory(plain, effective);
            }
            ExcelFileFormat format = ClassifyMemory(data, effective, out ZipIndex? memZip);
            if (format is ExcelFileFormat.Unknown)
            {
                memZip?.Dispose();
                UnknownFormatException();
            }
            return format switch
            {
                ExcelFileFormat.Xls => new XlsWorkbook(data, effective),
                ExcelFileFormat.Xlsb => XlsbWorkbook.CreateFromIndex(memZip!, effective),
                ExcelFileFormat.Xlsx => XlsxWorkbook.CreateFromIndex(memZip!, effective),
                _ => throw new System.Diagnostics.UnreachableException(),
            };
        }

        private static IExcelWorkbook OpenFromPlainMemory(ReadOnlyMemory<byte> plain, ExcelReaderOptions options)
        {
            ExcelFileFormat format = ClassifyMemory(plain, options, out ZipIndex? memZip);
            if (format is not (ExcelFileFormat.Xlsb or ExcelFileFormat.Xlsx))
            {
                memZip?.Dispose();
                UnknownFormatException();
            }
            return format switch
            {
                ExcelFileFormat.Xlsb => XlsbWorkbook.CreateFromIndex(memZip!, options),
                ExcelFileFormat.Xlsx => XlsxWorkbook.CreateFromIndex(memZip!, options),
                _ => throw new System.Diagnostics.UnreachableException(),
            };
        }

        private static ExcelFileFormat ClassifyMemory(ReadOnlyMemory<byte> data, ExcelReaderOptions options, out ZipIndex? memZip)
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
            memZip = ZipIndex.Create(data, options);
            return ClassifyZip(memZip);
        }

        /// <summary>
        /// Asynchronously opens a workbook from a file path, auto-detecting its format (XLSX/XLSB/XLS) from the
        /// file's signature and taking ownership of the file stream.
        /// </summary>
        /// <param name="path">The path to the workbook file.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        /// <returns>A format-agnostic <see cref="IExcelWorkbook"/> backed by the concrete workbook that matches the detected format.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="path"/> is <see langword="null"/>.</exception>
        /// <exception cref="InvalidDataException">The file's signature does not match a supported format.</exception>
        public static ValueTask<IExcelWorkbook> OpenAsync(string path, ExcelReaderOptions? options = null, CancellationToken ct = default)
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
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the workbook is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when <see langword="null"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        /// <returns>A format-agnostic <see cref="IExcelWorkbook"/> backed by the concrete workbook that matches the detected format.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="stream"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="stream"/> does not support seeking.</exception>
        /// <exception cref="InvalidDataException">The stream's signature does not match a supported format.</exception>
        public static ValueTask<IExcelWorkbook> OpenAsync(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null, CancellationToken ct = default)
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
        /// <returns>A format-agnostic <see cref="IExcelWorkbook"/>.</returns>
        /// <remarks>CSV carries no signature, so auto-detection never reports it; opening delimited text
        /// means naming <see cref="ExcelFileFormat.Csv"/> here.</remarks>
        /// <exception cref="ArgumentNullException"><paramref name="path"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="format"/> is <see cref="ExcelFileFormat.EncryptedOoxml"/> or not a defined value.</exception>
        /// <exception cref="InvalidDataException"><paramref name="format"/> is <see cref="ExcelFileFormat.Unknown"/> and the file's signature matches no supported format.</exception>
        public static IExcelWorkbook Open(string path, ExcelFileFormat format, ExcelReaderOptions? options = null)
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
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the workbook is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when
        /// <see langword="null"/>. <see cref="ExcelReaderOptions.Csv"/> supplies the dialect when
        /// <paramref name="format"/> is <see cref="ExcelFileFormat.Csv"/>.</param>
        /// <returns>A format-agnostic <see cref="IExcelWorkbook"/>.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="stream"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="format"/> is <see cref="ExcelFileFormat.EncryptedOoxml"/> or not a defined value.</exception>
        /// <exception cref="ArgumentException"><paramref name="format"/> is <see cref="ExcelFileFormat.Unknown"/> and <paramref name="stream"/> does not support seeking.</exception>
        /// <exception cref="InvalidDataException"><paramref name="format"/> is <see cref="ExcelFileFormat.Unknown"/> and the stream's signature matches no supported format.</exception>
        public static IExcelWorkbook Open(Stream stream, ExcelFileFormat format, bool leaveOpen = true, ExcelReaderOptions? options = null)
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
        /// <returns>A format-agnostic <see cref="IExcelWorkbook"/>.</returns>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="format"/> is <see cref="ExcelFileFormat.EncryptedOoxml"/> or not a defined value.</exception>
        /// <exception cref="InvalidDataException"><paramref name="format"/> is <see cref="ExcelFileFormat.Unknown"/> and the buffer's signature matches no supported format.</exception>
        public static IExcelWorkbook Open(ReadOnlyMemory<byte> data, ExcelFileFormat format, ExcelReaderOptions? options = null)
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
        /// <returns>A format-agnostic <see cref="IExcelWorkbook"/>.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="path"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="format"/> is <see cref="ExcelFileFormat.EncryptedOoxml"/> or not a defined value.</exception>
        /// <exception cref="InvalidDataException"><paramref name="format"/> is <see cref="ExcelFileFormat.Unknown"/> and the file's signature matches no supported format.</exception>
        public static ValueTask<IExcelWorkbook> OpenAsync(string path, ExcelFileFormat format, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(path);
            RequireOpenableFormat(format);
            return format switch
            {
                ExcelFileFormat.Csv => AsWorkbookAsync(FromCsvFileAsync(path, CsvOf(options), ct)),
                ExcelFileFormat.Xlsx => AsWorkbookAsync(FromXlsxFileAsync(path, options, ct)),
                ExcelFileFormat.Xlsb => AsWorkbookAsync(FromXlsbFileAsync(path, options, ct)),
                ExcelFileFormat.Xls => AsWorkbookAsync(FromXlsFileAsync(path, options, ct)),
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
        /// <param name="leaveOpen">When <see langword="true"/> (the default), <paramref name="stream"/> is not disposed when the workbook is disposed.</param>
        /// <param name="options">Resource limits and behavior toggles; <see cref="ExcelReaderOptions.Default"/> when
        /// <see langword="null"/>. <see cref="ExcelReaderOptions.Csv"/> supplies the dialect when
        /// <paramref name="format"/> is <see cref="ExcelFileFormat.Csv"/>.</param>
        /// <param name="ct">A token to cancel the open operation.</param>
        /// <returns>A format-agnostic <see cref="IExcelWorkbook"/>.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="stream"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="format"/> is <see cref="ExcelFileFormat.EncryptedOoxml"/> or not a defined value.</exception>
        /// <exception cref="ArgumentException"><paramref name="format"/> is <see cref="ExcelFileFormat.Unknown"/> and <paramref name="stream"/> does not support seeking.</exception>
        /// <exception cref="InvalidDataException"><paramref name="format"/> is <see cref="ExcelFileFormat.Unknown"/> and the stream's signature matches no supported format.</exception>
        public static ValueTask<IExcelWorkbook> OpenAsync(Stream stream, ExcelFileFormat format, bool leaveOpen = true, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(stream);
            RequireOpenableFormat(format);
            return format switch
            {
                ExcelFileFormat.Csv => AsWorkbookAsync(FromCsvAsync(stream, leaveOpen, CsvOf(options), ct)),
                ExcelFileFormat.Xlsx => AsWorkbookAsync(FromXlsxAsync(stream, leaveOpen, options, ct)),
                ExcelFileFormat.Xlsb => AsWorkbookAsync(FromXlsbAsync(stream, leaveOpen, options, ct)),
                ExcelFileFormat.Xls => AsWorkbookAsync(FromXlsAsync(stream, leaveOpen, options, ct)),
                _ => OpenAsync(stream, leaveOpen, options, ct),
            };
        }

        private static CsvReaderOptions CsvOf(ExcelReaderOptions? options)
        {
            return options?.Csv ?? CsvReaderOptions.Default;
        }

        private static async ValueTask<IExcelWorkbook> AsWorkbookAsync<TWorkbook>(ValueTask<TWorkbook> pending)
            where TWorkbook : IExcelWorkbook
        {
            return await pending.ConfigureAwait(false);
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
            ExcelFileFormat format = DetectSeekable(stream, leaveOpen: true, ExcelReaderOptions.Default, out ZipIndex? zip);
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
            var format = ClassifyMemory(data, ExcelReaderOptions.Default, out ZipIndex? memZip);
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
            (ExcelFileFormat format, ZipIndex? zip) = await DetectSeekableAsync(stream, leaveOpen: true, ExcelReaderOptions.Default, ct).ConfigureAwait(false);
            zip?.Dispose();
            return format;
        }

        private static IExcelWorkbook OpenSeekable(Stream stream, bool leaveOpen, ExcelReaderOptions? options)
        {
            ExcelReaderOptions effective = options ?? ExcelReaderOptions.Default;
            ExcelFileFormat format;
            ZipIndex? zip;
            try
            {
                format = DetectSeekable(stream, leaveOpen, effective, out zip);
            }
            catch
            {
                DisposeOnFailure(stream, leaveOpen);
                throw;
            }
            if (format is ExcelFileFormat.Unknown)
            {
                UnknownFormat(stream, leaveOpen);
            }
            if (format is ExcelFileFormat.EncryptedOoxml)
            {
                Stream decrypted = EncryptedPackageOpener.Decrypt(stream, leaveOpen, effective);
                return OpenDecryptedZip(decrypted, effective);
            }
            return format switch
            {
                ExcelFileFormat.Xls => new XlsWorkbook(stream, leaveOpen, options),
                ExcelFileFormat.Xlsb => XlsbWorkbook.CreateFromIndex(zip!, effective),
                ExcelFileFormat.Xlsx => XlsxWorkbook.CreateFromIndex(zip!, effective),
                _ => throw new System.Diagnostics.UnreachableException(),
            };
        }

        private static IExcelWorkbook OpenDecryptedZip(Stream decrypted, ExcelReaderOptions options)
        {
            ZipIndex zip = ZipIndex.Create(ByteSource.FromStream(decrypted, leaveOpen: false), options);
            return ClassifyZip(zip) is ExcelFileFormat.Xlsb
                ? XlsbWorkbook.CreateFromIndex(zip, options)
                : XlsxWorkbook.CreateFromIndex(zip, options);
        }

        private static async ValueTask<IExcelWorkbook> OpenSeekableAsync(Stream stream, bool leaveOpen, ExcelReaderOptions? options, CancellationToken ct)
        {
            ExcelReaderOptions effective = options ?? ExcelReaderOptions.Default;
            ExcelFileFormat format;
            ZipIndex? zip;
            try
            {
                (format, zip) = await DetectSeekableAsync(stream, leaveOpen, effective, ct).ConfigureAwait(false);
            }
            catch
            {
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
                Stream decrypted = EncryptedPackageOpener.Decrypt(stream, leaveOpen, effective);
                return await OpenDecryptedZipAsync(decrypted, effective, ct).ConfigureAwait(false);
            }
            return format switch
            {
                ExcelFileFormat.Xls => await XlsWorkbook.CreateAsync(stream, leaveOpen, options, ct).ConfigureAwait(false),
                ExcelFileFormat.Xlsb => await XlsbWorkbook.CreateFromIndexAsync(zip!, effective, ct).ConfigureAwait(false),
                ExcelFileFormat.Xlsx => await XlsxWorkbook.CreateFromIndexAsync(zip!, effective, ct).ConfigureAwait(false),
                _ => throw new System.Diagnostics.UnreachableException(),
            };
        }

        private static async ValueTask<IExcelWorkbook> OpenDecryptedZipAsync(Stream decrypted, ExcelReaderOptions options, CancellationToken ct)
        {
            ByteSource source = await ByteSource.FromStreamAsync(decrypted, leaveOpen: false, ct).ConfigureAwait(false);
            ZipIndex zip = await ZipIndex.CreateAsync(source, options, ct).ConfigureAwait(false);
            return ClassifyZip(zip) is ExcelFileFormat.Xlsb
                ? await XlsbWorkbook.CreateFromIndexAsync(zip, options, ct).ConfigureAwait(false)
                : await XlsxWorkbook.CreateFromIndexAsync(zip, options, ct).ConfigureAwait(false);
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

        private static ExcelFileFormat ClassifyZip(ZipIndex zip)
        {
            return zip.TryGetEntry("xl/workbook.bin"u8, out _) ? ExcelFileFormat.Xlsb : ExcelFileFormat.Xlsx;
        }

        [SkipLocalsInit]
        private static ExcelFileFormat DetectSeekable(Stream stream, bool leaveOpen, ExcelReaderOptions options, out ZipIndex? zip)
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
            try
            {
                zip = ZipIndex.Create(ByteSource.FromStream(stream, leaveOpen), options);
            }
            finally
            {
                if (stream.CanSeek)
                {
                    stream.Position = start;
                }
            }
            return ClassifyZip(zip);
        }

        private static async ValueTask<(ExcelFileFormat Format, ZipIndex? Zip)> DetectSeekableAsync(
            Stream stream, bool leaveOpen, ExcelReaderOptions options, CancellationToken ct)
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
            ByteSource source = await ByteSource.FromStreamAsync(stream, leaveOpen, ct).ConfigureAwait(false);
            ZipIndex zip;
            try
            {
                zip = await ZipIndex.CreateAsync(source, options, ct).ConfigureAwait(false);
            }
            finally
            {
                if (stream.CanSeek)
                {
                    stream.Position = start;
                }
            }
            return (ClassifyZip(zip), zip);
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
