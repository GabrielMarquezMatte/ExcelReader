using System.Buffers.Binary;
using System.Runtime.InteropServices;
using ExcelReader.Core.Reader;

namespace ExcelReader.Native.Reading
{
    internal static unsafe partial class ReadApi
    {
        internal static int ReadAllBlob(NativeHandle? handle, Span<byte> buffer, out int written)
        {
            written = 0;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }

            NativeApi.ClearLastError();
            try
            {
                if (!handle.AllRowsPending)
                {
                    AccumulateAllRows(handle);
                }

                written = handle.AllRowsLength;
                if (buffer.Length < handle.AllRowsLength)
                {
                    return NativeStatus.BufferTooSmall;
                }

                BinaryPrimitives.WriteInt32LittleEndian(buffer, handle.AllRowsCount);
                handle.AllRowsScratch?.CopyTo(buffer[sizeof(int)..handle.AllRowsLength]);
                handle.AllRowsPending = false;
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                handle.AllRowsPending = false;
                handle.AllRowsLength = 0;
                handle.AllRowsCount = 0;
                handle.AllRowsScratch = null;
                return NativeStatus.Error;
            }
        }

        private static void AccumulateAllRows(NativeHandle handle)
        {
            ChunkedBuffer<byte> output = new();
            int rowCount = 0;

            if (handle.HasPending)
            {
                AppendRow(output, handle.Scratch.AsSpan(0, handle.PendingLength));
                handle.HasPending = false;
                rowCount++;
            }

            byte[] rowScratch = [];
            handle.FaultLiveSession("xl_read_all_blob");
            handle.Rows ??= handle.Reader.GetEnumerator();
            while (handle.Rows.MoveNext())
            {
                Row row = handle.Rows.Current;
                int rowLength = RowBlob.Serialize(row, ref rowScratch);
                AppendRow(output, rowScratch.AsSpan(0, rowLength));
                rowCount++;
            }

            handle.AllRowsScratch = output;
            handle.AllRowsCount = rowCount;
            handle.AllRowsLength = sizeof(int) + output.Count;
            handle.AllRowsPending = true;
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

        internal static int ReadAllDecoded(NativeHandle? handle, out NativeRows rows)
        {
            rows = default;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }

            List<NativeRow> decoded = [];
            try
            {
                while (true)
                {
                    int status = NextRowDecoded(handle, out NativeRow row);
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
