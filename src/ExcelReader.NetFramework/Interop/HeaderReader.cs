namespace ExcelReader.NetFramework.Interop
{
    internal static class HeaderReader
    {
        internal static Dictionary<string, int> Read(ExcelSheet sheet, int headerRow)
        {
            int index = 0;
            foreach (ExcelRow row in sheet.ReadRows())
            {
                index++;
                if (index == headerRow)
                {
                    return Map(row);
                }
            }

            throw new ExcelReaderException(ExcelStatus.InvalidArgument, $"The sheet has fewer than {headerRow} rows.");
        }

        private static Dictionary<string, int> Map(ExcelRow row)
        {
            Dictionary<string, int> map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (ExcelCell cell in row.Cells)
            {
                string? text = cell.Text?.Trim();
                if (text is null || map.ContainsKey(text))
                {
                    continue;
                }

                map[text] = cell.Column;
            }

            return map;
        }
    }
}
