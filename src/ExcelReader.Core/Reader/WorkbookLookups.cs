using System.IO.Compression;

namespace ExcelReader.Core.Reader
{
    // Small lookups duplicated identically across the XLSX/XLS/XLSB readers: sheet-name resolution,
    // sheet-index validation, date-style flags, and shared-string offset lookup. Each reader keeps its
    // own sheets/styles/shared arrays (the sheet tuple's non-name element differs per format), so these
    // take the array in rather than requiring a shared interface.
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

        // A numFmtId is a date style if it's a custom format flagged as such, else the builtin table decides.
        internal static bool ResolveDateFlag(Dictionary<int, bool> customFormats, int numFmtId)
        {
            return customFormats.TryGetValue(numFmtId, out bool isDate) ? isDate : NumberFormat.IsBuiltinDate(numFmtId);
        }

        internal static (int Start, int Length) SharedAt(int[] sharedOffsets, int index)
        {
            if ((uint)index >= (uint)(sharedOffsets.Length - 1))
            {
                return (0, 0);
            }
            return (sharedOffsets[index], sharedOffsets[index + 1] - sharedOffsets[index]);
        }

        internal static ZipArchiveEntry GetWorksheetEntry(ZipArchive zip, (string Name, string Path)[] sheets, int current)
        {
            return zip.GetEntry(sheets[current].Path)
                ?? throw new InvalidDataException($"Worksheet part not found: {sheets[current].Path}");
        }

        internal static LimitedReadStream OpenEntryStream(ZipArchiveEntry entry, DecompressedByteCounter counter)
        {
            return new LimitedReadStream(entry.Open(), counter);
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
