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
            try
            {
                ReadExactlyChecked(stream, rented.AsSpan(0, length));
            }
            catch
            {
                ArrayPool<byte>.Shared.Return(rented);
                throw;
            }
            return new ZipPart(rented.AsMemory(0, length), rented);
        }

        // PERF-3: an entry whose real decompressed data is *shorter* than its declared central-directory
        // size is a malformed file, not a raw BCL stream-plumbing exception — matches
        // ZipMemoryIndex.InflateToPart's equivalent rewrap on the in-memory path. There is no equivalent
        // over-delivery check here (InflateToPart has one): confirmed by direct experiment that
        // ZipArchiveEntry.Open()'s stream silently truncates at the entry's declared Length regardless of
        // how much more the underlying compressed data would actually produce — a trailing "is there
        // more?" read always reports EOF, so over-delivery is invisible on this BCL-backed path, not
        // fixable without bypassing ZipArchiveEntry.Open() entirely (a far larger change).
        private static void ReadExactlyChecked(Stream stream, Span<byte> destination)
        {
            try
            {
                stream.ReadExactly(destination);
            }
            catch (EndOfStreamException ex)
            {
                throw new InvalidDataException("The ZIP entry produced less data than its declared uncompressed size.", ex);
            }
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
                try
                {
                    await ReadExactlyCheckedAsync(stream, rented.AsMemory(0, length), ct).ConfigureAwait(false);
                }
                catch
                {
                    ArrayPool<byte>.Shared.Return(rented);
                    throw;
                }
                return new ZipPart(rented.AsMemory(0, length), rented);
            }
        }

        private static async ValueTask ReadExactlyCheckedAsync(Stream stream, Memory<byte> destination, CancellationToken ct)
        {
            try
            {
                await stream.ReadExactlyAsync(destination, ct).ConfigureAwait(false);
            }
            catch (EndOfStreamException ex)
            {
                throw new InvalidDataException("The ZIP entry produced less data than its declared uncompressed size.", ex);
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
