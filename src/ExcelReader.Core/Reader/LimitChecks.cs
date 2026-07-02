namespace ExcelReader.Core.Reader
{
    internal static class LimitChecks
    {
        internal static void ThrowIfOverCellLimit(ExcelReaderOptions options, long needed)
        {
            if (options.MaxCellBytes > 0 && needed > options.MaxCellBytes)
            {
                throw new ExcelLimitExceededException(nameof(ExcelReaderOptions.MaxCellBytes), options.MaxCellBytes, needed);
            }
        }

        internal static void ThrowIfOverSharedStringLimit(ExcelReaderOptions options, long needed)
        {
            if (options.MaxSharedStringBytes > 0 && needed > options.MaxSharedStringBytes)
            {
                throw new ExcelLimitExceededException(nameof(ExcelReaderOptions.MaxSharedStringBytes), options.MaxSharedStringBytes, needed);
            }
        }

        internal static int NextBufferSize(ExcelReaderOptions options, int current, int needed)
        {
            return NextBufferSize(options.MaxCellBytes, nameof(ExcelReaderOptions.MaxCellBytes), current, needed);
        }

        // Format-agnostic core shared by ExcelReaderOptions (XLSX/XLSB/XLS) and CsvReaderOptions —
        // both cap a single buffered cell/record the same way, just under different option types.
        internal static int NextBufferSize(int maxCellBytes, string limitName, int current, int needed)
        {
            long doubled = (long)current * 2;
            long next = Math.Max(doubled, needed);
            if (maxCellBytes > 0 && next > maxCellBytes)
            {
                throw new ExcelLimitExceededException(limitName, maxCellBytes, next);
            }
            if (next > Array.MaxLength)
            {
                throw new ExcelLimitExceededException("ArrayMaxLength", Array.MaxLength, next);
            }
            return (int)next;
        }
    }
}
