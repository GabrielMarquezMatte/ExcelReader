namespace ExcelReader.Core.Writer.Internal
{
    internal static class SheetColumnValidation
    {
        // Excel's widest column; BIFF8 stores the width in 16 bits as 256ths of a character.
        private const double MaxColumnWidth = 255;

        internal static void SetColumnStyle(
            ref Dictionary<int, int>? columnStyles, int columnIndex, int styleId, int styleCount, int maxColumns,
            WriterState state, object writer)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(columnIndex, maxColumns);
            ArgumentOutOfRangeException.ThrowIfNegative(styleId);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(styleId, styleCount);
            RequireNotStarted(state, writer);
            columnStyles ??= [];
            columnStyles[columnIndex] = styleId;
        }

        internal static void SetColumnWidth(
            ref Dictionary<int, double>? columnWidths, int columnIndex, double width, int maxColumns,
            WriterState state, object writer)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(columnIndex, maxColumns);
            ArgumentOutOfRangeException.ThrowIfNegative(width);
            ArgumentOutOfRangeException.ThrowIfGreaterThan(width, MaxColumnWidth);
            RequireNotStarted(state, writer);
            columnWidths ??= [];
            columnWidths[columnIndex] = width;
        }

        private static void RequireNotStarted(WriterState state, object writer)
        {
            WriterStateGuard.ThrowIfEnded(state, writer);
            if (state != WriterState.Created)
            {
                throw new InvalidOperationException(
                    $"{nameof(SetColumnStyle)}/{nameof(SetColumnWidth)} must be called before the first row is started.");
            }
        }
    }
}
