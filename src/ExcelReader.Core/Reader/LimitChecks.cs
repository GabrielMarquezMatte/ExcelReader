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
            long doubled = (long)current * 2;
            long next = Math.Max(doubled, needed);
            ThrowIfOverCellLimit(options, next);
            if (next > Array.MaxLength)
            {
                throw new ExcelLimitExceededException("ArrayMaxLength", Array.MaxLength, next);
            }
            return (int)next;
        }
    }
}
