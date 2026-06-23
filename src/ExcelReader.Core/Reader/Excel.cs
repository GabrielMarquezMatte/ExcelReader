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
    }
}
