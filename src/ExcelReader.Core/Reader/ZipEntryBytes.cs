using System.Buffers;
using System.IO.Compression;

namespace ExcelReader.Core.Reader
{
    internal static class ZipEntryBytes
    {
        // Guards entry.Length — the attacker-controlled central-directory uncompressed size — against
        // both the workbook-wide remaining budget and any smaller per-entry limit, before the caller
        // sizes a destination buffer from it.
        private static void ThrowIfEntryLengthExceedsLimits(
            ZipArchiveEntry entry, DecompressedByteCounter counter, string entryLimitName, long entryLimit)
        {
            LimitChecks.ThrowIfEntryLengthExceeds(entry.Length, counter.Remaining,
                nameof(ExcelReaderOptions.MaxTotalDecompressedBytes));
            if (entryLimit > 0)
            {
                LimitChecks.ThrowIfEntryLengthExceeds(entry.Length, entryLimit, entryLimitName);
            }
        }

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
            ThrowIfEntryLengthExceedsLimits(entry, counter, entryLimitName, entryLimit);
            using var stream = new LimitedReadStream(entry.Open(), counter, entryLimitName, entryLimit);
            // ZipArchiveEntry.Length is the exact uncompressed size from the central directory, so the
            // destination can be sized once instead of growing/copying through an intermediate MemoryStream.
            byte[] bytes = new byte[checked((int)entry.Length)];
            stream.ReadExactly(bytes);
            return bytes;
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
        internal static async ValueTask<byte[]> ReadAsync(
            ZipArchiveEntry entry,
            DecompressedByteCounter counter,
            CancellationToken ct,
            string entryLimitName = "",
            long entryLimit = 0)
        {
            ThrowIfEntryLengthExceedsLimits(entry, counter, entryLimitName, entryLimit);
            Stream opened = await entry.OpenAsync(ct).ConfigureAwait(false);
            var stream = new LimitedReadStream(opened, counter, entryLimitName, entryLimit);
            await using (stream.ConfigureAwait(false))
            {
                byte[] bytes = new byte[checked((int)entry.Length)];
                await stream.ReadExactlyAsync(bytes, ct).ConfigureAwait(false);
                return bytes;
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
