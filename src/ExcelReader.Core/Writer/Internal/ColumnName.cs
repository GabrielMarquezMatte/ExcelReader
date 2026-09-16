using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Writer.Internal
{
    internal static class ColumnName
    {
        internal static int Write(Span<byte> destination, int columnIndex)
        {
            ExcelLimits.ThrowIfColumnOutOfRange(columnIndex);
            if (columnIndex < 26)
            {
                destination[0] = (byte)('A' + columnIndex);
                return 1;
            }
            if (columnIndex < 702)
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
