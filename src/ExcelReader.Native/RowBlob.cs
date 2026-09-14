using System.Buffers.Binary;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Native
{
    // Serializes a Row into the flat little-endian blob described in
    // docs/plans/2026-08-13-native-ffi-python-reading.md (ABI Contract, "Row blob layout").
    //
    // Row and Cell are ref structs over the reader's internal buffers,
    // so they cannot be handed across the FFI boundary or stored between calls. Copying the whole
    // row once per call keeps the boundary to a single crossing per row and leaves the caller with
    // no lifetime rules to obey.
    internal static class RowBlob
    {
        // Byte size of one cell header (column, type, value_len) in the row blob. Shared with
        // NativeApi's decoder so the two can never disagree about the layout.
        internal const int CellHeaderSize = 3 * sizeof(int);
        private const int MaxFormattedValueLength = 32;

        // Writes row into scratch, growing it if needed. Returns the byte count.
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
