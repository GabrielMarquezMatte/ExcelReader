using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using System.Text;

namespace ExcelReader.Core.Writer.Internal
{
    // Shared "write one text part into a ZIP/OPC entry" and "flush the underlying stream" logic for
    // the ZIP-based workbook writers (WorkbookWriter / XlsbWorkbookWriter). XlsWorkbookWriter instead
    // assembles a single OLE container, so it has no ZIP entries to share this with.
    internal static class ZipEntryWriter
    {
        private static readonly UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Stream is disposed indirectly by StreamWriter with leaveOpen: false.")]
        internal static async ValueTask WriteTextAsync(ZipArchive zip, string entryName, string content, CompressionLevel compression, CancellationToken ct)
        {
            ZipArchiveEntry entry = zip.CreateEntry(entryName, compression);
#if NET10_0_OR_GREATER
            Stream stream = await entry.OpenAsync(ct).ConfigureAwait(false);
#else
            ct.ThrowIfCancellationRequested();
            Stream stream = entry.Open();
#endif
            StreamWriter writer = new(stream, Utf8NoBom, leaveOpen: false);
            await using (writer.ConfigureAwait(false))
            {
                await writer.WriteAsync(content).ConfigureAwait(false);
                await writer.FlushAsync(ct).ConfigureAwait(false);
            }
        }

        internal static ValueTask FlushAsync(Stream stream, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            return new ValueTask(stream.FlushAsync(ct));
        }
    }
}
