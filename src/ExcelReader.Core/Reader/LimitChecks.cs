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
