using System.Buffers.Binary;
using System.Runtime.InteropServices;
using ExcelReader.Core.Reader;

namespace ExcelReader.Native.Reading
{
    internal static unsafe partial class ReadApi
    {
        internal static int ReadAllBlob(NativeRowCursor? cursor, Span<byte> buffer, out int written)
        {
            written = 0;
            if (cursor is null)
            {
                return NativeStatus.InvalidHandle;
            }

            NativeApi.ClearLastError();
            try
            {
                if (!cursor.AllRowsPending)
                {
                    AccumulateAllRows(cursor);
                }

                written = cursor.AllRowsLength;
                if (buffer.Length < cursor.AllRowsLength)
                {
                    return NativeStatus.BufferTooSmall;
                }

                BinaryPrimitives.WriteInt32LittleEndian(buffer, cursor.AllRowsCount);
                cursor.AllRowsScratch?.CopyTo(buffer[sizeof(int)..cursor.AllRowsLength]);
                cursor.AllRowsPending = false;
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                cursor.AllRowsPending = false;
                cursor.AllRowsLength = 0;
                cursor.AllRowsCount = 0;
                cursor.AllRowsScratch = null;
                return NativeStatus.Error;
            }
        }

        private static void AccumulateAllRows(NativeRowCursor cursor)
        {
            ChunkedBuffer<byte> output = new();
            int rowCount = 0;

            if (cursor.HasPending)
            {
                AppendRow(output, cursor.Scratch.AsSpan(0, cursor.PendingLength));
                cursor.HasPending = false;
                rowCount++;
            }

            byte[] rowScratch = [];
            while (cursor.Rows.MoveNext())
            {
                Row row = cursor.Rows.Current;
                int rowLength = RowBlob.Serialize(row, ref rowScratch);
                AppendRow(output, rowScratch.AsSpan(0, rowLength));
                rowCount++;
            }

            cursor.AllRowsScratch = output;
            cursor.AllRowsCount = rowCount;
            cursor.AllRowsLength = sizeof(int) + output.Count;
            cursor.AllRowsPending = true;
        }

        private static void AppendRow(ChunkedBuffer<byte> output, ReadOnlySpan<byte> rowBlob)
        {
            long required = (long)sizeof(int) + output.Count + sizeof(int) + rowBlob.Length;
            if (required > int.MaxValue)
            {
                throw new InvalidOperationException(
                    "xl_read_all_blob's accumulated result exceeds the 2 GiB int32 limit of this API; " +
                    "use xl_parse_typed instead, which is columnar, uses int64_t lengths, and is markedly faster.");
            }

            Span<byte> header = stackalloc byte[sizeof(int)];
            BinaryPrimitives.WriteInt32LittleEndian(header, rowBlob.Length);
            output.AddRange(header);
            output.AddRange(rowBlob);
        }

        internal static int ReadAllDecoded(NativeRowCursor? cursor, out NativeRows rows)
        {
            rows = default;
            if (cursor is null)
            {
                return NativeStatus.InvalidHandle;
            }

            List<NativeRow> decoded = [];
            try
            {
                while (true)
                {
                    int status = NextRowDecoded(cursor, out NativeRow row);
                    if (status == NativeStatus.Eof)
                    {
                        break;
                    }
                    if (status != NativeStatus.Ok)
                    {
                        FreeAll(decoded);
                        return status;
                    }
                    decoded.Add(row);
                }

                if (decoded.Count == 0)
                {
                    return NativeStatus.Ok;
                }

                IntPtr block = Marshal.AllocHGlobal(checked(decoded.Count * sizeof(NativeRow)));
                NativeRow* target = (NativeRow*)block;
                for (int index = 0; index < decoded.Count; index++)
                {
                    target[index] = decoded[index];
                }

                rows = new NativeRows { RowCount = decoded.Count, Rows = block };
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                FreeAll(decoded);
                NativeApi.SetLastError(exception.Message);
                rows = default;
                return NativeStatus.Error;
            }

            static void FreeAll(List<NativeRow> rowsToFree)
            {
                foreach (ref readonly NativeRow row in CollectionsMarshal.AsSpan(rowsToFree))
                {
                    NativeRow toFree = row;
                    FreeRow(ref toFree);
                }
            }
        }

        internal static void FreeRows(ref NativeRows rows)
        {
            if (rows.Rows == IntPtr.Zero)
            {
                rows = default;
                return;
            }

            NativeRow* stored = (NativeRow*)rows.Rows;
            for (int index = 0; index < rows.RowCount; index++)
            {
                FreeRow(ref stored[index]);
            }
            Marshal.FreeHGlobal(rows.Rows);
            rows = default;
        }
    }
}
