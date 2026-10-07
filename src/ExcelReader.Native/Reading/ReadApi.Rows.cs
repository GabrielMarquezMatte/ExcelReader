using System.Buffers.Binary;
using System.Runtime.InteropServices;
using ExcelReader.Core.Reader;

namespace ExcelReader.Native.Reading
{
    internal static unsafe partial class ReadApi
    {
        internal static int OpenRows(NativeHandle? handle, int sheet, out NativeRowCursor? cursor)
        {
            cursor = null;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }

            NativeApi.ClearLastError();
            int status = ResolveSheet(handle, sheet, out IExcelSheet? resolved);
            if (status != NativeStatus.Ok)
            {
                return status;
            }

            try
            {
                cursor = new NativeRowCursor(resolved!.GetEnumerator());
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int CloseRows(NativeRowCursor? cursor)
        {
            if (cursor is null)
            {
                return NativeStatus.InvalidHandle;
            }

            NativeApi.ClearLastError();
            try
            {
                cursor.Dispose();
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int NextRowDecoded(NativeRowCursor? cursor, out NativeRow row)
        {
            row = default;
            int status = NextRow(cursor, Span<byte>.Empty, out _);
            if (status != NativeStatus.BufferTooSmall)
            {
                return status;
            }

            try
            {
                return DecodePendingRow(cursor!, out row);
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

        internal static int NextRow(NativeRowCursor? cursor, Span<byte> buffer, out int written)
        {
            written = 0;
            if (cursor is null)
            {
                return NativeStatus.InvalidHandle;
            }

            NativeApi.ClearLastError();
            try
            {
                if (!cursor.HasPending)
                {
                    if (!cursor.Rows.MoveNext())
                    {
                        return NativeStatus.Eof;
                    }

                    Row row = cursor.Rows.Current;
                    byte[] scratch = cursor.Scratch;
                    cursor.PendingLength = RowBlob.Serialize(row, ref scratch);
                    cursor.Scratch = scratch;
                    cursor.HasPending = true;
                }

                written = cursor.PendingLength;
                if (buffer.Length < cursor.PendingLength)
                {
                    return NativeStatus.BufferTooSmall;
                }

                cursor.Scratch.AsSpan(0, cursor.PendingLength).CopyTo(buffer);
                cursor.HasPending = false;
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int NextRowView(NativeRowCursor? cursor, out NativeRow row)
        {
            row = default;
            if (cursor is null)
            {
                return NativeStatus.InvalidHandle;
            }

            NativeApi.ClearLastError();
            try
            {
                cursor.View ??= new RowViewBuffer();
                if (cursor.HasPending)
                {
                    cursor.HasPending = false;
                    row = cursor.View.Fill(cursor.Scratch.AsSpan(0, cursor.PendingLength));
                    return NativeStatus.Ok;
                }

                if (!cursor.Rows.MoveNext())
                {
                    return NativeStatus.Eof;
                }
                row = cursor.View.Fill(cursor.Rows.Current);
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                row = default;
                return NativeStatus.Error;
            }
        }

        private static int DecodePendingRow(NativeRowCursor cursor, out NativeRow row)
        {
            row = default;
            ReadOnlySpan<byte> blob = cursor.Scratch.AsSpan(0, cursor.PendingLength);
            int cellCount = BinaryPrimitives.ReadInt32LittleEndian(blob);
            if (cellCount == 0)
            {
                cursor.HasPending = false;
                return NativeStatus.Ok;
            }

            int cellSize = sizeof(NativeRowCell);
            int valueBytes = cursor.PendingLength - sizeof(int) - checked(cellCount * RowBlob.CellHeaderSize);
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

            cursor.HasPending = false;
            row = new NativeRow { CellCount = cellCount, Cells = block };
            return NativeStatus.Ok;
        }
    }
}
