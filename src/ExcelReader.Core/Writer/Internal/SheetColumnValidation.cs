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

        // Styles are frozen once the first row starts, so cells read this array instead of the dictionary.
        internal static int[]? StyleLookup(Dictionary<int, int>? columnStyles)
        {
            if (columnStyles is null)
            {
                return null;
            }
            int[] lookup = new int[columnStyles.Keys.Max() + 1];
            foreach ((int columnIndex, int styleId) in columnStyles)
            {
                lookup[columnIndex] = styleId;
            }
            return lookup;
        }

        internal static int StyleAt(int[]? lookup, int columnIndex)
        {
            return lookup is not null && (uint)columnIndex < (uint)lookup.Length ? lookup[columnIndex] : 0;
        }

        internal static SortedSet<int> Columns(Dictionary<int, int>? columnStyles, Dictionary<int, double>? columnWidths)
        {
            var columns = new SortedSet<int>();
            if (columnStyles is not null)
            {
                columns.UnionWith(columnStyles.Keys);
            }
            if (columnWidths is not null)
            {
                columns.UnionWith(columnWidths.Keys);
            }
            return columns;
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
