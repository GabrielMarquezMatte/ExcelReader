using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Reader
{
    public static class Excel
    {
        public static XlsxReader FromFile(string path, ExcelReaderOptions? options = null)
        {
            return new XlsxReader(File.OpenRead(path), leaveOpen: false, options);
        }

        public static XlsxReader From(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null)
        {
            return new XlsxReader(stream, leaveOpen, options);
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream ownership transfers to XlsReader, which streams from it and disposes it on Dispose (and on construction failure).")]
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Stream ownership transfers to XlsReader, which streams from it and disposes it on Dispose (and on construction failure).")]
        public static XlsReader FromXlsFile(string path, ExcelReaderOptions? options = null)
        {
            return new XlsReader(File.OpenRead(path), leaveOpen: false, options);
        }

        public static XlsReader FromXls(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null)
        {
            return new XlsReader(stream, leaveOpen, options);
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream ownership transfers to XlsbReader on success, disposed on failure.")]
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Stream ownership transfers to XlsbReader on success, disposed on failure.")]
        public static XlsbReader FromXlsbFile(string path, ExcelReaderOptions? options = null)
        {
            return new XlsbReader(File.OpenRead(path), leaveOpen: false, options);
        }

        public static XlsbReader FromXlsb(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null)
        {
            return new XlsbReader(stream, leaveOpen, options);
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream ownership transfers to CreateAsync, which disposes it on failure and via the reader on success.")]
        public static ValueTask<XlsxReader> FromFileAsync(string path, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            FileStream stream = OpenAsyncFile(path);
            return XlsxReader.CreateAsync(stream, leaveOpen: false, options, ct);
        }

        public static ValueTask<XlsxReader> FromAsync(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            return XlsxReader.CreateAsync(stream, leaveOpen, options, ct);
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream ownership transfers to CreateAsync, which disposes it on failure and is consumed into the reader on success.")]
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Stream ownership transfers to CreateAsync, which disposes it on failure and via the reader on success.")]
        public static ValueTask<XlsReader> FromXlsFileAsync(string path, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            FileStream stream = OpenAsyncFile(path);
            return XlsReader.CreateAsync(stream, leaveOpen: false, options, ct);
        }

        public static ValueTask<XlsReader> FromXlsAsync(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            return XlsReader.CreateAsync(stream, leaveOpen, options, ct);
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream ownership transfers to CreateAsync, which disposes it on failure and via the reader on success.")]
        public static ValueTask<XlsbReader> FromXlsbFileAsync(string path, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            FileStream stream = OpenAsyncFile(path);
            return XlsbReader.CreateAsync(stream, leaveOpen: false, options, ct);
        }

        public static ValueTask<XlsbReader> FromXlsbAsync(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            return XlsbReader.CreateAsync(stream, leaveOpen, options, ct);
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream ownership transfers to CsvReader, which disposes it on Dispose/DisposeAsync when leaveOpen is false.")]
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Stream ownership transfers to CsvReader, which disposes it on Dispose/DisposeAsync when leaveOpen is false.")]
        public static CsvReader FromCsvFile(string path, CsvReaderOptions? options = null)
        {
            return new CsvReader(File.OpenRead(path), leaveOpen: false, options);
        }

        public static CsvReader FromCsv(Stream stream, bool leaveOpen = true, CsvReaderOptions? options = null)
        {
            return new CsvReader(stream, leaveOpen, options);
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream ownership transfers to CreateAsync, which disposes it on failure and via the reader on success.")]
        public static ValueTask<CsvReader> FromCsvFileAsync(string path, CsvReaderOptions? options = null, CancellationToken ct = default)
        {
            FileStream stream = OpenAsyncFile(path);
            return CsvReader.CreateAsync(stream, leaveOpen: false, options, ct);
        }

        public static ValueTask<CsvReader> FromCsvAsync(Stream stream, bool leaveOpen = true, CsvReaderOptions? options = null, CancellationToken ct = default)
        {
            return CsvReader.CreateAsync(stream, leaveOpen, options, ct);
        }

        // The leading bytes that distinguish container formats: XLSX and XLSB are ZIP ("PK\x03\x04"),
        // XLS is an OLE2/CFB compound document. XLSB is distinguished from XLSX by the presence of
        // "xl/workbook.bin" in the ZIP central directory.
        private static ReadOnlySpan<byte> ZipSignature => [0x50, 0x4B, 0x03, 0x04];
        // Opens a workbook of either format, choosing the reader from the file's signature.
        // The returned reader iterates rows through its concrete type (XlsxReader / XlsReader)
        // pattern-match on the result to enumerate.
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream ownership transfers to OpenSeekable, which disposes it on failure and via the reader on success.")]
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Stream ownership transfers to OpenSeekable, which disposes it on failure and via the reader on success.")]
        public static IExcelRowReader Open(string path, ExcelReaderOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(path);
            return OpenSeekable(File.OpenRead(path), leaveOpen: false, options);
        }

        public static IExcelRowReader Open(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(stream);
            return OpenSeekable(stream, leaveOpen, options);
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream ownership transfers to OpenSeekableAsync, which disposes it on failure and via the reader on success.")]
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Stream ownership transfers to OpenSeekableAsync, which disposes it on failure and via the reader on success.")]
        public static ValueTask<IExcelRowReader> OpenAsync(string path, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(path);
            FileStream stream = OpenAsyncFile(path);
            return OpenSeekableAsync(stream, leaveOpen: false, options, ct);
        }

        public static ValueTask<IExcelRowReader> OpenAsync(Stream stream, bool leaveOpen = true, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(stream);
            return OpenSeekableAsync(stream, leaveOpen, options, ct);
        }

        public static ExcelFileFormat DetectFileFormat(Stream stream)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ExcelFileFormat format = DetectSeekable(stream, out ZipArchive? zip);
            zip?.Dispose();
            return format;
        }

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

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "Open and OpenAsync transfer ownership when leaveOpen is false.")]
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
