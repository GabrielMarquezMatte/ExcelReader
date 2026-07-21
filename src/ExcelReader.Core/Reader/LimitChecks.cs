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
