namespace ExcelReader.Core.Reader
{
    /// <summary>One sheet's position and name, as yielded by <see cref="ExcelRowReaderExtensions.Sheets"/>.</summary>
    /// <param name="Index">The sheet's zero-based index.</param>
    /// <param name="Name">The sheet's name.</param>
    public readonly record struct ExcelSheet(int Index, string Name);

    /// <summary>Convenience methods layered on <see cref="IExcelRowReader"/>.</summary>
    public static class ExcelRowReaderExtensions
    {
        /// <summary>
        /// Walks every sheet in <paramref name="reader"/>'s workbook, selecting each as the current sheet
        /// before yielding it.
        /// </summary>
        /// <remarks>
        /// Selecting a sheet moves the reader's one shared cursor (see <see cref="IExcelRowReader"/>'s
        /// thread-safety remarks), so read that sheet's rows inside the loop body before the next
        /// iteration moves on to the following sheet.
        /// </remarks>
        /// <param name="reader">The workbook reader to walk.</param>
        public static IEnumerable<ExcelSheet> Sheets(this IExcelRowReader reader)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return EnumerateSheets(reader);
        }

        private static IEnumerable<ExcelSheet> EnumerateSheets(IExcelRowReader reader)
        {
            for (int i = 0; i < reader.SheetCount; i++)
            {
                reader.MoveToSheet(i);
                yield return new ExcelSheet(i, reader.SheetName);
            }
        }
    }
}
