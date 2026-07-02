using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;

namespace ExcelReader.Core.Reader
{
    [SuppressMessage("Design", "CA1068:CancellationToken parameters must come last",
        Justification = "ExcelReaderOptions was added after existing CancellationToken parameters to preserve source compatibility.")]
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
        public static ValueTask<XlsxReader> FromFileAsync(string path, CancellationToken ct = default, ExcelReaderOptions? options = null)
        {
            FileStream stream = new(path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 65536,
                                    options: FileOptions.Asynchronous | FileOptions.SequentialScan);
            return XlsxReader.CreateAsync(stream, leaveOpen: false, options, ct);
        }

        public static ValueTask<XlsxReader> FromAsync(Stream stream, bool leaveOpen = true, CancellationToken ct = default, ExcelReaderOptions? options = null)
        {
            return XlsxReader.CreateAsync(stream, leaveOpen, options, ct);
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream ownership transfers to CreateAsync, which disposes it on failure and is consumed into the reader on success.")]
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Stream ownership transfers to CreateAsync, which disposes it on failure and via the reader on success.")]
        public static ValueTask<XlsReader> FromXlsFileAsync(string path, CancellationToken ct = default, ExcelReaderOptions? options = null)
        {
            FileStream stream = new(path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 65536,
                                    options: FileOptions.Asynchronous | FileOptions.SequentialScan);
            return XlsReader.CreateAsync(stream, leaveOpen: false, options, ct);
        }

        public static ValueTask<XlsReader> FromXlsAsync(Stream stream, bool leaveOpen = true, CancellationToken ct = default, ExcelReaderOptions? options = null)
        {
            return XlsReader.CreateAsync(stream, leaveOpen, options, ct);
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream ownership transfers to CreateAsync, which disposes it on failure and via the reader on success.")]
        public static ValueTask<XlsbReader> FromXlsbFileAsync(string path, CancellationToken ct = default, ExcelReaderOptions? options = null)
        {
            FileStream stream = new(path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 65536,
                                    options: FileOptions.Asynchronous | FileOptions.SequentialScan);
            return XlsbReader.CreateAsync(stream, leaveOpen: false, options, ct);
        }

        public static ValueTask<XlsbReader> FromXlsbAsync(Stream stream, bool leaveOpen = true, CancellationToken ct = default, ExcelReaderOptions? options = null)
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
        public static ValueTask<CsvReader> FromCsvFileAsync(string path, CancellationToken ct = default, CsvReaderOptions? options = null)
        {
            FileStream stream = new(path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 65536,
                                    options: FileOptions.Asynchronous | FileOptions.SequentialScan);
            return CsvReader.CreateAsync(stream, leaveOpen: false, options, ct);
        }

        public static ValueTask<CsvReader> FromCsvAsync(Stream stream, bool leaveOpen = true, CancellationToken ct = default, CsvReaderOptions? options = null)
        {
            return CsvReader.CreateAsync(stream, leaveOpen, options, ct);
        }

        // The leading bytes that distinguish container formats: XLSX and XLSB are ZIP ("PK\x03\x04"),
        // XLS is an OLE2/CFB compound document. XLSB is distinguished from XLSX by the presence of
        // "xl/workbook.bin" in the ZIP central directory.
        private static ReadOnlySpan<byte> ZipSignature => [0x50, 0x4B, 0x03, 0x04];
        private static ReadOnlySpan<byte> OleSignature => [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

        private enum Format { Xlsx, Xls, Xlsb }

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
        public static ValueTask<IExcelRowReader> OpenAsync(string path, CancellationToken ct = default, ExcelReaderOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(path);
            FileStream stream = new(path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 65536,
                                    options: FileOptions.Asynchronous | FileOptions.SequentialScan);
            return OpenSeekableAsync(stream, leaveOpen: false, options, ct);
        }

        public static ValueTask<IExcelRowReader> OpenAsync(Stream stream, bool leaveOpen = true, CancellationToken ct = default, ExcelReaderOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(stream);
            return OpenSeekableAsync(stream, leaveOpen, options, ct);
        }

        private static IExcelRowReader OpenSeekable(Stream stream, bool leaveOpen, ExcelReaderOptions? options)
        {
            Format format;
            try
            {
                format = DetectSeekable(stream);
            }
            catch
            {
                if (!leaveOpen)
                {
                    stream.Dispose();
                }
                throw;
            }
            return format switch
            {
                Format.Xls => new XlsReader(stream, leaveOpen, options),
                Format.Xlsb => new XlsbReader(stream, leaveOpen, options),
                Format.Xlsx => new XlsxReader(stream, leaveOpen, options),
                _ => throw new System.Diagnostics.UnreachableException(),
            };
        }

        private static async ValueTask<IExcelRowReader> OpenSeekableAsync(Stream stream, bool leaveOpen, ExcelReaderOptions? options, CancellationToken ct)
        {
            Format format;
            try
            {
                format = await DetectSeekableAsync(stream, ct).ConfigureAwait(false);
            }
            catch
            {
                if (!leaveOpen)
                {
                    await stream.DisposeAsync().ConfigureAwait(false);
                }
                throw;
            }
            return format switch
            {
                Format.Xls => await XlsReader.CreateAsync(stream, leaveOpen, options, ct).ConfigureAwait(false),
                Format.Xlsb => await XlsbReader.CreateAsync(stream, leaveOpen, options, ct).ConfigureAwait(false),
                Format.Xlsx => await XlsxReader.CreateAsync(stream, leaveOpen, options, ct).ConfigureAwait(false),
                _ => throw new System.Diagnostics.UnreachableException(),
            };
        }

        // Detection peeks the 8-byte signature then rewinds. For ZIP streams, opens a temporary
        // ZipArchive to distinguish XLSB ("xl/workbook.bin" present) from XLSX (XML workbook).
        // Both XLSX and XLSB readers need a seekable source anyway (ZipArchive seeks the central
        // directory), so requiring seek here costs nothing and keeps the peek cheap.
        private static Format DetectSeekable(Stream stream)
        {
            RequireSeekable(stream);
            long start = stream.Position;
            Span<byte> header = stackalloc byte[8];
            int read = stream.ReadAtLeast(header, header.Length, throwOnEndOfStream: false);
            stream.Position = start;
            ReadOnlySpan<byte> sig = header[..read];
            if (sig.StartsWith(OleSignature))
            {
                return Format.Xls;
            }
            if (!sig.StartsWith(ZipSignature))
            {
                throw new InvalidDataException("Unrecognized file format; expected an XLSX/XLSB (ZIP) or XLS (OLE2) workbook.");
            }
            // Peek the central directory to distinguish XLSB from XLSX.
            using var zipPeek = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
            bool isXlsb = zipPeek.GetEntry("xl/workbook.bin") is not null;
            stream.Position = start;
            return isXlsb ? Format.Xlsb : Format.Xlsx;
        }

        private static async ValueTask<Format> DetectSeekableAsync(Stream stream, CancellationToken ct)
        {
            RequireSeekable(stream);
            long start = stream.Position;
            byte[] header = new byte[8];
            int read = await stream.ReadAtLeastAsync(header, header.Length, throwOnEndOfStream: false, ct).ConfigureAwait(false);
            stream.Position = start;
            ReadOnlySpan<byte> sig = header.AsSpan(0, read);
            if (sig.StartsWith(OleSignature))
            {
                return Format.Xls;
            }
            if (!sig.StartsWith(ZipSignature))
            {
                throw new InvalidDataException("Unrecognized file format; expected an XLSX/XLSB (ZIP) or XLS (OLE2) workbook.");
            }
            // Central directory read: open a temporary archive to peek entry names, then rewind.
            // Declared outside await using so the ZipArchive variable is accessible inside the block.
            var zipPeek = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
#if NET10_0_OR_GREATER
            await using (zipPeek.ConfigureAwait(false))
#else
            using (zipPeek)
#endif
            {
                bool isXlsb = zipPeek.GetEntry("xl/workbook.bin") is not null;
                stream.Position = start;
                return isXlsb ? Format.Xlsb : Format.Xlsx;
            }
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
    }
}
