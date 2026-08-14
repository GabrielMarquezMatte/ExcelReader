using System.Buffers.Binary;
using System.Runtime.InteropServices;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Native
{
    internal static unsafe partial class NativeApi
    {
        internal static int NextRowDecoded(NativeHandle? handle, out NativeRow row)
        {
            row = default;
            int status = NextRow(handle, Span<byte>.Empty, out _);
            if (status != NativeStatus.BufferTooSmall)
            {
                return status;
            }

            try
            {
                return DecodePendingRow(handle!, out row);
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                FreeRow(ref row);
                return NativeStatus.Error;
            }
        }

        /// <summary>
        /// Releases the single block allocated by <see cref="DecodePendingRow"/>. Every
        /// <see cref="NativeRowCell.Value"/> points INTO that block (see <see cref="DecodePendingRow"/>), so
        /// freeing them individually would be a double free — this one <see cref="Marshal.FreeHGlobal(IntPtr)"/>
        /// covers the cell array and every value.
        /// </summary>
        internal static void FreeRow(ref NativeRow row)
        {
            if (row.Cells != IntPtr.Zero)
            {
                Marshal.FreeHGlobal(row.Cells);
            }
            row = default;
        }

        internal static int NextRow(NativeHandle? handle, Span<byte> buffer, out int written)
        {
            written = 0;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }

            ClearLastError();
            try
            {
                if (!handle.HasPending)
                {
                    handle.Rows ??= handle.Reader.GetEnumerator();
                    if (!handle.Rows.MoveNext())
                    {
                        return NativeStatus.Eof;
                    }

                    Row row = handle.Rows.Current;
                    byte[] scratch = handle.Scratch;
                    handle.PendingLength = RowBlob.Serialize(row, ref scratch);
                    handle.Scratch = scratch;
                    handle.HasPending = true;
                }

                written = handle.PendingLength;
                if (buffer.Length < handle.PendingLength)
                {
                    return NativeStatus.BufferTooSmall;
                }

                handle.Scratch.AsSpan(0, handle.PendingLength).CopyTo(buffer);
                handle.HasPending = false;
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        /// <summary>
        /// Decodes the row blob held pending in <paramref name="handle"/>'s scratch buffer into a
        /// <see cref="NativeRow"/> backed by a SINGLE allocation: the <see cref="NativeRowCell"/> array
        /// followed immediately by every cell's value bytes (each NUL-terminated), laid out back to back.
        /// Every <see cref="NativeRowCell.Value"/> pointer is an interior pointer into that one block — this
        /// is what lets <see cref="FreeRow"/> release a whole row with one <see cref="Marshal.FreeHGlobal"/>
        /// instead of one call per cell, cutting a 65K-row/10-column sheet from ~650K native allocations to
        /// one per row (see docs/NATIVE_BINDINGS_PLAN.md Task 1).
        /// </summary>
        private static int DecodePendingRow(NativeHandle handle, out NativeRow row)
        {
            row = default;
            ReadOnlySpan<byte> blob = handle.Scratch.AsSpan(0, handle.PendingLength);
            int cellCount = BinaryPrimitives.ReadInt32LittleEndian(blob);
            if (cellCount == 0)
            {
                handle.HasPending = false;
                return NativeStatus.Ok;
            }

            int cellSize = sizeof(NativeRowCell);
            int valueBytes = handle.PendingLength - sizeof(int) - checked(cellCount * RowBlob.CellHeaderSize);
            IntPtr block = Marshal.AllocHGlobal(checked((cellCount * cellSize) + valueBytes + cellCount));

            NativeRowCell* cells = (NativeRowCell*)block;
            byte* values = (byte*)block + (cellCount * cellSize);
            int offset = sizeof(int);
            int valueOffset = 0;
            for (int index = 0; index < cellCount; index++)
            {
                int column = BinaryPrimitives.ReadInt32LittleEndian(blob[offset..]);
                int type = BinaryPrimitives.ReadInt32LittleEndian(blob[(offset + 4)..]);
                int length = BinaryPrimitives.ReadInt32LittleEndian(blob[(offset + 8)..]);
                offset += RowBlob.CellHeaderSize;

                blob.Slice(offset, length).CopyTo(new Span<byte>(values + valueOffset, length));
                values[valueOffset + length] = 0;
                cells[index] = new NativeRowCell
                {
                    Column = column,
                    Type = type,
                    ValueLength = length,
                    Value = (IntPtr)(values + valueOffset),
                };
                offset += length;
                valueOffset += length + 1;
            }

            handle.HasPending = false;
            row = new NativeRow { CellCount = cellCount, Cells = block };
            return NativeStatus.Ok;
        }
    }
}
