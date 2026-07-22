using System.Buffers;
using System.IO.Compression;
using System.Text;

namespace ExcelReader.Core.Writer.Internal
{
    // Shared "write one text part into a ZIP/OPC entry" and "flush the underlying stream" logic for
    // the ZIP-based workbook writers (XlsxWorkbookWriter / XlsbWorkbookWriter). XlsWorkbookWriter instead
    // assembles a single OLE container, so it has no ZIP entries to share this with.
    internal static class ZipEntryWriter
    {
        private static readonly UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);

        internal static async ValueTask WriteTextAsync(ZipArchive zip, string entryName, string content, CompressionLevel compression, CancellationToken ct)
        {
            // The caller already holds the full content as one string, so encode it in a single pass
            // instead of routing it through StreamWriter's small internal buffer.
            byte[] rented = ArrayPool<byte>.Shared.Rent(Utf8NoBom.GetByteCount(content));
            try
            {
                int written = Utf8NoBom.GetBytes(content, rented);
                await WriteBytesAsync(zip, entryName, rented.AsMemory(0, written), compression, ct).ConfigureAwait(false);
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(rented);
            }
        }

        internal static async ValueTask WriteBytesAsync(ZipArchive zip, string entryName, ReadOnlyMemory<byte> content, CompressionLevel compression, CancellationToken ct)
        {
            ZipArchiveEntry entry = zip.CreateEntry(entryName, compression);
#if NET10_0_OR_GREATER
            Stream stream = await entry.OpenAsync(ct).ConfigureAwait(false);
#else
            ct.ThrowIfCancellationRequested();
            Stream stream = entry.Open();
#endif
            await using (stream.ConfigureAwait(false))
            {
                await stream.WriteAsync(content, ct).ConfigureAwait(false);
            }
        }

        internal static ValueTask FlushAsync(Stream stream, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            return new ValueTask(stream.FlushAsync(ct));
        }
    }
}
