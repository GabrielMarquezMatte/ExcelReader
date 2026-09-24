using System.Buffers.Binary;
using System.Runtime.InteropServices;
using ExcelReader.Core.Reader;

namespace ExcelReader.Native.Reading
{
    internal static unsafe partial class ReadApi
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
                NativeApi.SetLastError(exception.Message);
                FreeRow(ref row);
                return NativeStatus.Error;
            }
        }

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

            NativeApi.ClearLastError();
            try
            {
                if (!handle.HasPending)
                {
                    handle.FaultLiveSession("xl_next_row");
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
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

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
