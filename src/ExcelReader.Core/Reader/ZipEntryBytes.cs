using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;

namespace ExcelReader.Core.Reader
{
    internal static class ZipEntryBytes
    {
        internal static byte[] Read(
            ZipArchive zip,
            string name,
            DecompressedByteCounter counter,
            string entryLimitName = "",
            long entryLimit = 0)
        {
            ZipArchiveEntry? entry = zip.GetEntry(name);
            return entry is null ? [] : Read(entry, counter, entryLimitName, entryLimit);
        }

        internal static byte[] Read(
            ZipArchiveEntry entry,
            DecompressedByteCounter counter,
            string entryLimitName = "",
            long entryLimit = 0)
        {
            using var stream = new LimitedReadStream(entry.Open(), counter, entryLimitName, entryLimit);
            using MemoryStream bytes = new();
            stream.CopyTo(bytes);
            return bytes.ToArray();
        }

        internal static async ValueTask<byte[]> ReadAsync(
            ZipArchive zip,
            string name,
            DecompressedByteCounter counter,
            CancellationToken ct,
            string entryLimitName = "",
            long entryLimit = 0)
        {
            ZipArchiveEntry? entry = zip.GetEntry(name);
            if (entry is null)
            {
                return [];
            }
            return await ReadAsync(entry, counter, ct, entryLimitName, entryLimit).ConfigureAwait(false);
        }

#if NET10_0_OR_GREATER
        [SuppressMessage("SharpSource", "SS059:MemoryStream can be disposed of asynchronously",
            Justification = "MemoryStream disposal is synchronous and keeps this helper simple.")]
        internal static async ValueTask<byte[]> ReadAsync(
            ZipArchiveEntry entry,
            DecompressedByteCounter counter,
            CancellationToken ct,
            string entryLimitName = "",
            long entryLimit = 0)
        {
            Stream opened = await entry.OpenAsync(ct).ConfigureAwait(false);
            var stream = new LimitedReadStream(opened, counter, entryLimitName, entryLimit);
            await using (stream.ConfigureAwait(false))
            {
                using MemoryStream bytes = new();
                await stream.CopyToAsync(bytes, ct).ConfigureAwait(false);
                return bytes.ToArray();
            }
        }
#else
        internal static ValueTask<byte[]> ReadAsync(
            ZipArchiveEntry entry,
            DecompressedByteCounter counter,
            CancellationToken ct,
            string entryLimitName = "",
            long entryLimit = 0)
        {
            ct.ThrowIfCancellationRequested();
            return new ValueTask<byte[]>(Read(entry, counter, entryLimitName, entryLimit));
        }
#endif
    }
}
