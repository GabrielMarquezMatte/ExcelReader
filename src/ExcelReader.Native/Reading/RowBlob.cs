using System.Buffers.Binary;
using ExcelReader.Core.Reader;

namespace ExcelReader.Native.Reading
{
    internal static class RowBlob
    {
        internal const int CellHeaderSize = 3 * sizeof(int);
        private const int MaxFormattedValueLength = 32;

        internal static int Serialize(in Row row, ref byte[] scratch)
        {
            int required = sizeof(int);
            foreach (RowCell cell in row.Cells)
            {
                int valueEstimate = cell.Value.Value.IsEmpty ? MaxFormattedValueLength : cell.Value.Value.Length;
                required += CellHeaderSize + valueEstimate;
            }

            if (scratch.Length < required)
            {
                scratch = new byte[required];
            }

            Span<byte> destination = scratch;
            int offset = sizeof(int);
            int count = 0;
            foreach (RowCell cell in row.Cells)
            {
                if (!cell.Value.TryFormat(destination[(offset + CellHeaderSize)..], out int bytesWritten))
                {
                    throw new InvalidOperationException("Cell format buffer too small");
                }
                BinaryPrimitives.WriteInt32LittleEndian(destination[offset..], cell.ColumnIndex);
                BinaryPrimitives.WriteInt32LittleEndian(destination[(offset + 4)..], (int)cell.Value.Type);
                BinaryPrimitives.WriteInt32LittleEndian(destination[(offset + 8)..], bytesWritten);
                offset += CellHeaderSize + bytesWritten;
                count++;
            }

            BinaryPrimitives.WriteInt32LittleEndian(destination, count);
            return offset;
        }
    }
}
