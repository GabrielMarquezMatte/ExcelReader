using System.Buffers;
using System.IO.Compression;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Core.Reader.Zip
{
    internal static class ZipEntryBytes
    {
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

    }
}
