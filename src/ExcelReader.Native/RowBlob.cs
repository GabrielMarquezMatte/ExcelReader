using System.Buffers.Binary;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Native
{
    /// <summary>
    /// Serializes a <see cref="Row"/> into the flat little-endian blob described in
    /// docs/plans/2026-08-13-native-ffi-python-reading.md (ABI Contract, "Row blob layout").
    /// </summary>
    /// <remarks>
    /// <see cref="Row"/> and <see cref="Cell"/> are ref structs over the reader's internal buffers,
    /// so they cannot be handed across the FFI boundary or stored between calls. Copying the whole
    /// row once per call keeps the boundary to a single crossing per row and leaves the caller with
    /// no lifetime rules to obey.
    /// </remarks>
    internal static class RowBlob
    {
        private const int CellHeaderSize = 3 * sizeof(int);

        /// <summary>Writes <paramref name="row"/> into <paramref name="scratch"/>, growing it if needed. Returns the byte count.</summary>
        internal static int Serialize(in Row row, ref byte[] scratch)
        {
            int required = sizeof(int);
            foreach (RowCell cell in row.Cells)
            {
                required += CellHeaderSize + cell.Value.Value.Length;
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
                ReadOnlySpan<byte> value = cell.Value.Value;
                BinaryPrimitives.WriteInt32LittleEndian(destination[offset..], cell.ColumnIndex);
                BinaryPrimitives.WriteInt32LittleEndian(destination[(offset + 4)..], (int)cell.Value.Type);
                BinaryPrimitives.WriteInt32LittleEndian(destination[(offset + 8)..], value.Length);
                offset += CellHeaderSize;
                value.CopyTo(destination[offset..]);
                offset += value.Length;
                count++;
            }

            BinaryPrimitives.WriteInt32LittleEndian(destination, count);
            return offset;
        }
    }
}
