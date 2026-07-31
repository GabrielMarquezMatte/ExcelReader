namespace ExcelReader.Core.Reader
{
    internal static class LimitChecks
    {
        internal static void ThrowIfOverSharedStringLimit(ExcelReaderOptions options, long needed)
        {
            if (options.MaxSharedStringBytes > 0 && needed > options.MaxSharedStringBytes)
            {
                throw new ExcelLimitExceededException(nameof(ExcelReaderOptions.MaxSharedStringBytes), options.MaxSharedStringBytes, needed);
            }
        }

        // Rejects an entry whose declared uncompressed size alone exceeds the cap, before anything
        // is allocated. Central-directory sizes are attacker-controlled, so this must run before the
        // destination buffer is sized — checking only actual decompressed bytes (as the streaming
        // counter does) is too late to stop a crafted small file from requesting a multi-GB buffer.
        internal static void ThrowIfEntryLengthExceeds(long declaredLength, long limit, string limitName)
        {
            if (limit > 0 && declaredLength > limit)
            {
                throw new ExcelLimitExceededException(limitName, limit, declaredLength);
            }
        }

        // A zip with an excessive entry count builds a large central-directory dictionary at open,
        // even though readers only ever GetEntry() a handful of well-known part names by name.
        internal static void ThrowIfTooManyEntries(int entryCount, ExcelReaderOptions options)
        {
            if (options.MaxZipEntries > 0 && entryCount > options.MaxZipEntries)
            {
                throw new ExcelLimitExceededException(nameof(ExcelReaderOptions.MaxZipEntries), options.MaxZipEntries, entryCount);
            }
        }

        // <sst uniqueCount="…"> is an attacker-controlled attribute independent of the part's actual
        // byte length — a file can be a few hundred bytes on disk yet declare uniqueCount in the
        // hundreds of millions, sizing the offsets array to that count before a single <si> is parsed.
        // The declared count can never legitimately exceed what the part could physically contain: the
        // smallest possible entry is "<si/>" (5 bytes), so partLength / 5 is a safe, generous ceiling
        // that also makes the "uniqueCount + 1" addition downstream incapable of overflowing (this
        // bound is always far below int.MaxValue for any MaxSharedStringBytes a caller would configure).
        private const int MinBytesPerSharedStringEntry = 5;

        internal static void ThrowIfSharedStringCountImplausible(int uniqueCount, int partLength)
        {
            long maxPlausibleCount = (long)partLength / MinBytesPerSharedStringEntry;
            if (uniqueCount > maxPlausibleCount)
            {
                throw new ExcelLimitExceededException(nameof(ExcelReaderOptions.MaxSharedStringBytes), maxPlausibleCount, uniqueCount);
            }
        }

        // Format-agnostic core shared by ExcelReaderOptions (XLSX/XLSB/XLS) and CsvReaderOptions —
        // both cap a single buffered cell/record the same way, just under different option types.
        internal static int NextBufferSize(int maxCellBytes, string limitName, int current, int needed, int elementSize = 1)
        {
            long doubled = (long)current * 2;
            long next = Math.Max(doubled, needed);
            long bytes = next * elementSize;
            if (maxCellBytes > 0 && bytes > maxCellBytes)
            {
                throw new ExcelLimitExceededException(limitName, maxCellBytes, bytes);
            }
            if (next > Array.MaxLength)
            {
                throw new ExcelLimitExceededException("ArrayMaxLength", Array.MaxLength, next);
            }
            return (int)next;
        }
    }
}
