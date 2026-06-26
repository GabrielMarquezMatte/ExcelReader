using System.Diagnostics.CodeAnalysis;

namespace ExcelReader.Core.Reader
{
    public static class Excel
    {
        public static XlsxReader FromFile(string path)
        {
            return new XlsxReader(File.OpenRead(path), leaveOpen: false);
        }

        public static XlsxReader From(Stream stream, bool leaveOpen = true)
        {
            return new XlsxReader(stream, leaveOpen);
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream ownership transfers to XlsReader, which streams from it and disposes it on Dispose (and on construction failure).")]
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Stream ownership transfers to XlsReader, which streams from it and disposes it on Dispose (and on construction failure).")]
        public static XlsReader FromXlsFile(string path)
        {
            return new XlsReader(File.OpenRead(path), leaveOpen: false);
        }

        public static XlsReader FromXls(Stream stream, bool leaveOpen = true)
        {
            return new XlsReader(stream, leaveOpen);
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream ownership transfers to CreateAsync, which disposes it on failure and via the reader on success.")]
        public static ValueTask<XlsxReader> FromFileAsync(string path, CancellationToken ct = default)
        {
            FileStream stream = new(path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 65536,
                                    options: FileOptions.Asynchronous | FileOptions.SequentialScan);
            return XlsxReader.CreateAsync(stream, leaveOpen: false, ct);
        }

        public static ValueTask<XlsxReader> FromAsync(Stream stream, bool leaveOpen = true, CancellationToken ct = default)
        {
            return XlsxReader.CreateAsync(stream, leaveOpen, ct);
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream ownership transfers to CreateAsync, which disposes it on failure and is consumed into the reader on success.")]
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Stream ownership transfers to CreateAsync, which disposes it on failure and via the reader on success.")]
        public static ValueTask<XlsReader> FromXlsFileAsync(string path, CancellationToken ct = default)
        {
            FileStream stream = new(path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 65536,
                                    options: FileOptions.Asynchronous | FileOptions.SequentialScan);
            return XlsReader.CreateAsync(stream, leaveOpen: false, ct);
        }

        public static ValueTask<XlsReader> FromXlsAsync(Stream stream, bool leaveOpen = true, CancellationToken ct = default)
        {
            return XlsReader.CreateAsync(stream, leaveOpen, ct);
        }

        // The leading bytes that distinguish the two container formats: XLSX is a ZIP
        // ("PK\x03\x04"), XLS is an OLE2/CFB compound document.
        private static ReadOnlySpan<byte> ZipSignature => [0x50, 0x4B, 0x03, 0x04];
        private static ReadOnlySpan<byte> OleSignature => [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

        private enum Format { Xlsx, Xls }

        // Opens a workbook of either format, choosing the reader from the file's signature.
        // The returned reader iterates rows through its concrete type (XlsxReader / XlsReader)
        // pattern-match on the result to enumerate.
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream ownership transfers to OpenSeekable, which disposes it on failure and via the reader on success.")]
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Stream ownership transfers to OpenSeekable, which disposes it on failure and via the reader on success.")]
        public static IExcelReader Open(string path)
        {
            ArgumentNullException.ThrowIfNull(path);
            return OpenSeekable(File.OpenRead(path), leaveOpen: false);
        }

        public static IExcelReader Open(Stream stream, bool leaveOpen = true)
        {
            ArgumentNullException.ThrowIfNull(stream);
            return OpenSeekable(stream, leaveOpen);
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream ownership transfers to OpenSeekableAsync, which disposes it on failure and via the reader on success.")]
        public static ValueTask<IExcelReader> OpenAsync(string path, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(path);
            FileStream stream = new(path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 65536,
                                    options: FileOptions.Asynchronous | FileOptions.SequentialScan);
            return OpenSeekableAsync(stream, leaveOpen: false, ct);
        }

        public static ValueTask<IExcelReader> OpenAsync(Stream stream, bool leaveOpen = true, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(stream);
            return OpenSeekableAsync(stream, leaveOpen, ct);
        }

        private static IExcelReader OpenSeekable(Stream stream, bool leaveOpen)
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
            if (format == Format.Xls)
            {
                return new XlsReader(stream, leaveOpen);
            }
            return new XlsxReader(stream, leaveOpen);
        }

        private static async ValueTask<IExcelReader> OpenSeekableAsync(Stream stream, bool leaveOpen, CancellationToken ct)
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
            if (format == Format.Xls)
            {
                return await XlsReader.CreateAsync(stream, leaveOpen, ct).ConfigureAwait(false);
            }
            return await XlsxReader.CreateAsync(stream, leaveOpen, ct).ConfigureAwait(false);
        }

        // Detection peeks the signature then rewinds, so the chosen reader sees the stream
        // at its original position. XLSX already needs a seekable source (ZipArchive seeks the
        // central directory), so requiring seek here costs nothing and keeps the peek cheap.
        private static Format DetectSeekable(Stream stream)
        {
            RequireSeekable(stream);
            long start = stream.Position;
            Span<byte> header = stackalloc byte[8];
            int read = stream.ReadAtLeast(header, header.Length, throwOnEndOfStream: false);
            stream.Position = start;
            return Detect(header[..read]);
        }

        private static async ValueTask<Format> DetectSeekableAsync(Stream stream, CancellationToken ct)
        {
            RequireSeekable(stream);
            long start = stream.Position;
            byte[] header = new byte[8];
            int read = await stream.ReadAtLeastAsync(header, header.Length, throwOnEndOfStream: false, ct).ConfigureAwait(false);
            stream.Position = start;
            return Detect(header.AsSpan(0, read));
        }

        private static void RequireSeekable(Stream stream)
        {
            if (!stream.CanSeek)
            {
                throw new ArgumentException(
                    "Open requires a seekable stream so the format signature can be detected. Buffer the stream first, or call From/FromXls directly.",
                    nameof(stream));
            }
        }

        private static Format Detect(ReadOnlySpan<byte> header)
        {
            if (header.StartsWith(ZipSignature))
            {
                return Format.Xlsx;
            }
            if (header.StartsWith(OleSignature))
            {
                return Format.Xls;
            }
            throw new InvalidDataException("Unrecognized file format; expected an XLSX (ZIP) or XLS (OLE2) workbook.");
        }
    }
}
