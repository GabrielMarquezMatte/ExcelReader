namespace ExcelReader.Core.Writer.Internal
{
    internal static class ColumnName
    {
        // Returns the number of chars written. Max 3 chars (Excel limit: 16384 columns = "XFD").
        internal static int Write(Span<char> destination, int columnIndex)
        {
            if (columnIndex < 26)
            {
                destination[0] = (char)('A' + columnIndex);
                return 1;
            }
            if (columnIndex < 702) // 26 + 26*26
            {
                columnIndex -= 26;
                destination[0] = (char)('A' + (columnIndex / 26));
                destination[1] = (char)('A' + (columnIndex % 26));
                return 2;
            }
            columnIndex -= 702;
            destination[0] = (char)('A' + (columnIndex / 676));
            destination[1] = (char)('A' + (columnIndex / 26 % 26));
            destination[2] = (char)('A' + (columnIndex % 26));
            return 3;
        }
    }
}
