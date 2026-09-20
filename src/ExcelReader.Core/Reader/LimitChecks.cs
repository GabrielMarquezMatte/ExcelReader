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

        internal static void ThrowIfEntryLengthExceeds(long declaredLength, long limit, string limitName)
        {
            if (limit > 0 && declaredLength > limit)
            {
                throw new ExcelLimitExceededException(limitName, limit, declaredLength);
            }
        }

        internal static void ThrowIfTooManyEntries(int entryCount, ExcelReaderOptions options)
        {
            if (options.MaxZipEntries > 0 && entryCount > options.MaxZipEntries)
            {
                throw new ExcelLimitExceededException(nameof(ExcelReaderOptions.MaxZipEntries), options.MaxZipEntries, entryCount);
            }
        }

        private const int MinBytesPerSharedStringEntry = 5;

        internal static void ThrowIfSharedStringCountImplausible(int uniqueCount, int partLength)
        {
            long maxPlausibleCount = (long)partLength / MinBytesPerSharedStringEntry;
            if (uniqueCount > maxPlausibleCount)
            {
                throw new ExcelLimitExceededException(nameof(ExcelReaderOptions.MaxSharedStringBytes), maxPlausibleCount, uniqueCount);
            }
        }

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
