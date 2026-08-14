using System.Runtime.InteropServices;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Native
{
    internal static partial class NativeApi
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

        /// <summary>Releases all memory allocated for a row returned by <see cref="NextRowDecoded"/>.</summary>
        internal static void FreeRow(ref NativeRow row)
        {
            if (row.Cells == IntPtr.Zero)
            {
                row = default;
                return;
            }
            for (int index = 0; index < row.CellCount; index++)
            {
                NativeRowCell cell = Marshal.PtrToStructure<NativeRowCell>(IntPtr.Add(row.Cells, index * Marshal.SizeOf<NativeRowCell>()));
                IntPtr value = cell.Value;
                if (value != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(value);
                }
            }
            Marshal.FreeHGlobal(row.Cells);
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

        private static int DecodePendingRow(NativeHandle handle, out NativeRow row)
        {
            ReadOnlySpan<byte> blob = handle.Scratch.AsSpan(0, handle.PendingLength);
            int cellCount = BitConverter.ToInt32(blob);
            int cellSize = Marshal.SizeOf<NativeRowCell>();
            IntPtr cells = IntPtr.Zero;
            if (cellCount > 0)
            {
                cells = Marshal.AllocHGlobal(checked(cellCount * cellSize));
            }

            NativeRow result = new()
            {
                Cells = cells,
            };

            try
            {
                int offset = sizeof(int);
                for (int index = 0; index < cellCount; index++)
                {
                    int column = BitConverter.ToInt32(blob[offset..]);
                    int type = BitConverter.ToInt32(blob[(offset + sizeof(int))..]);
                    int valueLength = BitConverter.ToInt32(blob[(offset + (2 * sizeof(int)))..]);
                    offset += 3 * sizeof(int);

                    IntPtr value = Marshal.AllocHGlobal(checked(valueLength + 1));
                    Marshal.Copy(blob.Slice(offset, valueLength).ToArray(), 0, value, valueLength);
                    Marshal.WriteByte(value, valueLength, 0);
                    Marshal.StructureToPtr(new NativeRowCell
                    {
                        Column = column,
                        Type = type,
                        ValueLength = valueLength,
                        Value = value,
                    }, IntPtr.Add(result.Cells, index * cellSize), false);
                    result.CellCount++;
                    offset += valueLength;
                }

                handle.HasPending = false;
                row = result;
                return NativeStatus.Ok;
            }
            catch
            {
                FreeRow(ref result);
                throw;
            }
        }
    }
}
