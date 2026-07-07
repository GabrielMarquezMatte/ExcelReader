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

        internal static (int Start, int Length) SharedAt(int[] sharedOffsets, int index)
        {
            if ((uint)index >= (uint)(sharedOffsets.Length - 1))
            {
                return (0, 0);
            }
            return (sharedOffsets[index], sharedOffsets[index + 1] - sharedOffsets[index]);
        }
    }
}
