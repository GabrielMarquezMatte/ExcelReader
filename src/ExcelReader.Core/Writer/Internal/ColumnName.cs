using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Writer.Internal
{
    internal static class ColumnName
    {
        // Returns the number of bytes written. Max 3 bytes (Excel limit: 16384 columns = "XFD").
        internal static int Write(Span<byte> destination, int columnIndex)
        {
            if ((uint)columnIndex >= 16_384)
            {
                throw new ExcelLimitExceededException("Columns", 16_384, columnIndex + 1L);
            }
            if (columnIndex < 26)
            {
                destination[0] = (byte)('A' + columnIndex);
                return 1;
            }
            if (columnIndex < 702) // 26 + 26*26
            {
                columnIndex -= 26;
                destination[0] = (byte)('A' + (columnIndex / 26));
                destination[1] = (byte)('A' + (columnIndex % 26));
                return 2;
            }
            columnIndex -= 702;
            destination[0] = (byte)('A' + (columnIndex / 676));
            destination[1] = (byte)('A' + (columnIndex / 26 % 26));
            destination[2] = (byte)('A' + (columnIndex % 26));
            return 3;
        }
    }
}
