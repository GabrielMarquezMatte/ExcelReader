using System.Buffers;
using System.IO.Compression;
using System.Text;

namespace ExcelReader.Core.Writer.Internal
{
    internal static class ZipEntryWriter
    {
        private static readonly UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);

        internal static async ValueTask WriteTextAsync(ZipArchive zip, string entryName, string content, CompressionLevel compression, CancellationToken ct)
        {
            byte[] rented = ArrayPool<byte>.Shared.Rent(Utf8NoBom.GetMaxByteCount(content.Length));
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
            Stream stream = await entry.OpenAsync(ct).ConfigureAwait(false);
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


        internal static void WriteText(ZipArchive zip, string entryName, string content, CompressionLevel compression)
        {
            byte[] rented = ArrayPool<byte>.Shared.Rent(Utf8NoBom.GetMaxByteCount(content.Length));
            try
            {
                int written = Utf8NoBom.GetBytes(content, rented);
                WriteBytes(zip, entryName, rented.AsSpan(0, written), compression);
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(rented);
            }
        }

        internal static void WriteBytes(ZipArchive zip, string entryName, ReadOnlySpan<byte> content, CompressionLevel compression)
        {
            ZipArchiveEntry entry = zip.CreateEntry(entryName, compression);
            using Stream stream = entry.Open();
            stream.Write(content);
        }

        internal static void Flush(Stream stream)
        {
            stream.Flush();
        }
    }
}
