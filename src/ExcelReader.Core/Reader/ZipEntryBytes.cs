using System.Buffers;
using System.IO.Compression;

namespace ExcelReader.Core.Reader
{
    // Reads a whole ZIP entry (workbook.xml/.rels/styles.xml and the XLSB equivalents) into a pooled
    // buffer instead of a GC array: these are transient parse inputs, read once and discarded within
    // the same constructor/factory method, so there is no reason to hand the caller anything that
    // outlives a `using` block. Returns ZipPart (defined in ZipMemoryIndex.cs) for the same reason the
    // in-memory ZIP path already does: a pooled array with an exact-length view, Dispose returns it.
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

        internal static ZipPart Read(
            ZipArchive zip,
            string name,
            DecompressedByteCounter counter,
            string entryLimitName = "",
            long entryLimit = 0)
        {
            ZipArchiveEntry? entry = zip.GetEntry(name);
            return entry is null ? default : Read(entry, counter, entryLimitName, entryLimit);
        }

        internal static ZipPart Read(
            ZipArchiveEntry entry,
            DecompressedByteCounter counter,
            string entryLimitName = "",
            long entryLimit = 0)
        {
            ThrowIfEntryLengthExceedsLimits(entry, counter, entryLimitName, entryLimit);
            using var stream = new LimitedReadStream(entry.Open(), counter, entryLimitName, entryLimit);
            // ZipArchiveEntry.Length is the exact uncompressed size from the central directory, so the
            // destination can be sized once instead of growing/copying through an intermediate MemoryStream.
            int length = checked((int)entry.Length);
            byte[] rented = ArrayPool<byte>.Shared.Rent(length);
            stream.ReadExactly(rented.AsSpan(0, length));
            return new ZipPart(rented.AsMemory(0, length), rented);
        }

        internal static async ValueTask<ZipPart> ReadAsync(
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
                return default;
            }
            return await ReadAsync(entry, counter, ct, entryLimitName, entryLimit).ConfigureAwait(false);
        }

#if NET10_0_OR_GREATER
        internal static async ValueTask<ZipPart> ReadAsync(
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
                int length = checked((int)entry.Length);
                byte[] rented = ArrayPool<byte>.Shared.Rent(length);
                await stream.ReadExactlyAsync(rented.AsMemory(0, length), ct).ConfigureAwait(false);
                return new ZipPart(rented.AsMemory(0, length), rented);
            }
        }
#else
        internal static ValueTask<ZipPart> ReadAsync(
            ZipArchiveEntry entry,
            DecompressedByteCounter counter,
            CancellationToken ct,
            string entryLimitName = "",
            long entryLimit = 0)
        {
            ct.ThrowIfCancellationRequested();
            return new ValueTask<ZipPart>(Read(entry, counter, entryLimitName, entryLimit));
        }
#endif

    }
}
