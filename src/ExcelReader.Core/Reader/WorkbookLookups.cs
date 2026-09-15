using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using System.Text;

namespace ExcelReader.Core.Reader
{
    // Each reader keeps its own sheets/styles/shared arrays (the sheet tuple's non-name element differs
    // per format), so these take the array in rather than requiring a shared interface.
    internal static class WorkbookLookups
    {
        internal static bool TryFindSheetIndex<T>(T[] sheets, ReadOnlySpan<char> name, Func<T, string> nameOf, out int index)
        {
            for (int i = 0; i < sheets.Length; i++)
            {
                if (name.Equals(nameOf(sheets[i]), StringComparison.OrdinalIgnoreCase))
                {
                    index = i;
                    return true;
                }
            }
            index = -1;
            return false;
        }

        internal static void ValidateSheetIndex(int index, int count)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(index);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(index, count);
        }

        internal static bool IsDateStyle(bool[] styleIsDate, int style)
        {
            return (uint)style < (uint)styleIsDate.Length && styleIsDate[style];
        }

        internal static bool ResolveDateFlag(Dictionary<int, bool> customFormats, int numFmtId)
        {
            return customFormats.TryGetValue(numFmtId, out bool isDate) ? isDate : NumberFormat.IsBuiltinDate(numFmtId);
        }

        // Rejects an oversized shared-strings part before its entry stream is touched at all. The streaming
        // path never materializes one destination buffer sized from the (attacker-controlled)
        // central-directory length, but that declared length must still be checked against both caps up
        // front — validate before allocating, not before using.
        internal static void ThrowIfSharedEntryTooLarge(
            long declaredLength, DecompressedByteCounter counter, ExcelReaderOptions options)
        {
            LimitChecks.ThrowIfEntryLengthExceeds(declaredLength, counter.Remaining,
                nameof(ExcelReaderOptions.MaxTotalDecompressedBytes));
            if (options.MaxSharedStringBytes > 0)
            {
                LimitChecks.ThrowIfEntryLengthExceeds(declaredLength, options.MaxSharedStringBytes,
                    nameof(ExcelReaderOptions.MaxSharedStringBytes));
            }
        }

        // An array indexed by shared-string index, not a Dictionary<int,string>: the table's exact count
        // is known up front, so one array avoids the resize/rehash churn and LOH pressure an unsized
        // Dictionary pays at high cardinality. Capped so a workbook declaring an extreme count cannot
        // force one huge eager allocation; above the cap GetString() still returns the right value and
        // only loses dedup.
        private const int MaxCachedSharedStrings = 4_000_000; // ~32 MB of string? references

        internal static string?[] CreateSharedStringCache(int[] sharedOffsets)
        {
            int count = sharedOffsets.Length - 1;
            return count is > 0 and <= MaxCachedSharedStrings ? new string?[count] : [];
        }

        internal static (int Start, int Length, int ValidIndex) SharedAt(int[] sharedOffsets, int index)
        {
            if ((uint)index >= (uint)(sharedOffsets.Length - 1))
            {
                return (0, 0, -1);
            }
            return (sharedOffsets[index], sharedOffsets[index + 1] - sharedOffsets[index], index);
        }

        internal static ZipArchiveEntry GetWorksheetEntry(ZipArchive zip, (string Name, string Path)[] sheets, int current)
        {
            return zip.GetEntry(sheets[current].Path)
                ?? throw new InvalidDataException($"Worksheet part not found: {sheets[current].Path}");
        }

        // In-memory ZIP path twin of the ZipArchive overload above; same exception, same message shape.
        // Encodes the part path on the stack instead of Encoding.UTF8.GetBytes(string) — this runs on
        // every GetEnumerator() call (once per sheet switch), not just once at construction, so an
        // avoidable heap allocation here would repeat for the reader's whole lifetime.
        [SkipLocalsInit]
        internal static ZipEntryRef GetWorksheetEntry(ZipMemoryIndex memZip, (string Name, string Path)[] sheets, int current)
        {
            string path = sheets[current].Path;
            Span<byte> stackBuffer = stackalloc byte[256];
            ReadOnlySpan<byte> utf8Path = Encoding.UTF8.GetByteCount(path) <= stackBuffer.Length
                ? stackBuffer[..Encoding.UTF8.GetBytes(path, stackBuffer)]
                : Encoding.UTF8.GetBytes(path);
            return memZip.TryGetEntry(utf8Path, out ZipEntryRef entry)
                ? entry
                : throw new InvalidDataException($"Worksheet part not found: {path}");
        }

        // entryLimitName/entryLimit carry a per-part cap (e.g. MaxSharedStringBytes) that
        // LimitedReadStream enforces on top of the workbook-wide counter; parts relying solely on
        // MaxTotalDecompressedBytes omit them.
        internal static LimitedReadStream OpenEntryStream(
            ZipArchiveEntry entry, DecompressedByteCounter counter, ExcelReaderOptions options,
            string entryLimitName = "", long entryLimit = 0)
        {
            return Wrap(entry.Open(), counter, options, entryLimitName, entryLimit, entry.Length);
        }

        internal static async ValueTask<LimitedReadStream> OpenEntryStreamAsync(
            ZipArchiveEntry entry, DecompressedByteCounter counter, ExcelReaderOptions options,
            CancellationToken ct, string entryLimitName = "", long entryLimit = 0)
        {
            Stream opened = await entry.OpenAsync(ct).ConfigureAwait(false);
            return Wrap(opened, counter, options, entryLimitName, entryLimit, entry.Length);
        }

        // Below this, the overlap isn't worth a dedicated thread + producer/consumer handoff: a small
        // sheet decompresses faster than the Task.Run dispatch and teardown join cost it, so prefetch
        // would only add overhead. Matches InitialBufferCapacity's own 256 KB ceiling below.
        private const long PrefetchMinUncompressedSize = 256 * 1024;

        // Sole branch point for PrefetchDecompression, shared by the sync and async openers (and by
        // ZipMemoryIndex.OpenEntryStream, so the in-memory ZIP path gets the same prefetch overlap).
        // Wrapping order is load-bearing: prefetch innermost, limits outermost, so DecompressedByteCounter
        // accounting stays on the consumer thread and byte-for-byte identical to the serial path.
        internal static LimitedReadStream Wrap(
            Stream opened, DecompressedByteCounter counter, ExcelReaderOptions options,
            string entryLimitName, long entryLimit, long uncompressedSize)
        {
            if (!options.PrefetchDecompression || uncompressedSize < PrefetchMinUncompressedSize)
            {
                return new LimitedReadStream(opened, counter, entryLimitName, entryLimit);
            }
            return new LimitedReadStream(new PrefetchStream(opened), counter, entryLimitName, entryLimit);
        }

        // Sizes a worksheet's initial read buffer to its actual uncompressed size (entry.Length is exact,
        // from the ZIP central directory) instead of a fixed 64 KB — fewer DeflateStream.Read interop
        // transitions and PrepareBuffer compaction memmoves for sheets larger than that. Capped at 256 KB
        // so a huge sheet doesn't over-allocate; floored at 4 KB so a tiny/unknown-length entry doesn't
        // create a pathologically small buffer that immediately needs to grow.
        internal static int InitialBufferCapacity(long entryLength)
        {
            const int Min = 4 * 1024;
            const int Max = 256 * 1024;
            return entryLength <= 0 ? 64 * 1024 : (int)Math.Clamp(entryLength, Min, Max);
        }
    }
}
