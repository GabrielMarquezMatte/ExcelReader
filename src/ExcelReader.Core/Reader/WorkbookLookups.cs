using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using System.Text;

namespace ExcelReader.Core.Reader
{
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

        private const int MaxCachedSharedStrings = 4_000_000;

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

        internal static ZipArchiveEntry GetWorksheetEntry(ZipArchive zip, string path)
        {
            return zip.GetEntry(path)
                ?? throw new InvalidDataException($"Worksheet part not found: {path}");
        }

        [SkipLocalsInit]
        internal static ZipEntryRef GetWorksheetEntry(ZipMemoryIndex memZip, string path)
        {
            Span<byte> stackBuffer = stackalloc byte[256];
            ReadOnlySpan<byte> utf8Path = Encoding.UTF8.GetByteCount(path) <= stackBuffer.Length
                ? stackBuffer[..Encoding.UTF8.GetBytes(path, stackBuffer)]
                : Encoding.UTF8.GetBytes(path);
            return memZip.TryGetEntry(utf8Path, out ZipEntryRef entry)
                ? entry
                : throw new InvalidDataException($"Worksheet part not found: {path}");
        }

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

        private const long PrefetchMinUncompressedSize = 256 * 1024;

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

        internal static int InitialBufferCapacity(long entryLength)
        {
            const int Min = 4 * 1024;
            const int Max = 256 * 1024;
            return entryLength <= 0 ? 64 * 1024 : (int)Math.Clamp(entryLength, Min, Max);
        }
    }
}
