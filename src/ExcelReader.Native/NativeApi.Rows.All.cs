using System.Buffers.Binary;
using ExcelReader.Core.ValueObjects;
using System.Runtime.InteropServices;

namespace ExcelReader.Native
{
    internal static partial class NativeApi
    {
        /// <summary>
        /// Writes every remaining row of the current sheet into <paramref name="buffer"/> as one
        /// caller-owned blob (see excelreader.h for the layout) — the batch counterpart of
        /// <see cref="NextRow"/>, with zero native heap allocations for the row/cell data itself.
        /// On <see cref="NativeStatus.BufferTooSmall"/>, the accumulated bytes are held on
        /// <paramref name="handle"/> so a retry with a bigger buffer costs one copy, not a re-read —
        /// mirroring the single-row pending protocol <see cref="NextRow"/> already uses.
        /// </summary>
        internal static int ReadAllBlob(NativeHandle? handle, Span<byte> buffer, out int written)
        {
            written = 0;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }

            ClearLastError();
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

                handle.AllRowsScratch.AsSpan(0, handle.AllRowsLength).CopyTo(buffer);
                handle.AllRowsPending = false;
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                handle.AllRowsPending = false;
                handle.AllRowsLength = 0;
                return NativeStatus.Error;
            }
        }

        // Drains every remaining row of the current sheet (including one already held pending from a
        // prior xl_next_row that returned XL_BUFFER_TOO_SMALL) into handle.AllRowsScratch as
        //     int32 row_count
        //     repeated: int32 row_length, <row blob>
        // The underlying enumerator is fully consumed by the time this returns — a too-small buffer on
        // the ReadAllBlob call above does not mean rows remain unread, only that the accumulated result
        // hasn't been copied out to the caller yet.
        private static void AccumulateAllRows(NativeHandle handle)
        {
            byte[] output = handle.AllRowsScratch.Length > 0 ? handle.AllRowsScratch : new byte[4096];
            int offset = sizeof(int); // reserved for row_count, written once the count is known
            int rowCount = 0;

            if (handle.HasPending)
            {
                AppendRow(ref output, ref offset, handle.Scratch.AsSpan(0, handle.PendingLength));
                handle.HasPending = false;
                rowCount++;
            }

            byte[] rowScratch = [];
            handle.Rows ??= handle.Reader.GetEnumerator();
            while (handle.Rows.MoveNext())
            {
                Row row = handle.Rows.Current;
                int rowLength = RowBlob.Serialize(row, ref rowScratch);
                AppendRow(ref output, ref offset, rowScratch.AsSpan(0, rowLength));
                rowCount++;
            }

            BinaryPrimitives.WriteInt32LittleEndian(output, rowCount);
            handle.AllRowsScratch = output;
            handle.AllRowsLength = offset;
            handle.AllRowsPending = true;
        }

        // Appends one `int32 row_length` + `row blob` entry to `output`, growing it if needed.
        private static void AppendRow(ref byte[] output, ref int offset, ReadOnlySpan<byte> rowBlob)
        {
            // offset/capacity are int32 throughout this API (see excelreader.h on xl_read_all_blob),
            // so compute the next required size in a wider type first and give a precise message
            // when a sheet's accumulated blob would exceed that limit, rather than letting Array.Resize
            // fail on an oversized request with a generic "array too large" message.
            long required = (long)offset + sizeof(int) + rowBlob.Length;
            if (required > int.MaxValue)
            {
                throw new InvalidOperationException(
                    "xl_read_all_blob's accumulated result exceeds the 2 GiB int32 limit of this API; " +
                    "use xl_parse_typed instead, which is columnar, uses int64_t lengths, and is markedly faster.");
            }

            EnsureCapacity(ref output, (int)required);
            BinaryPrimitives.WriteInt32LittleEndian(output.AsSpan(offset), rowBlob.Length);
            rowBlob.CopyTo(output.AsSpan(offset + sizeof(int)));
            offset += sizeof(int) + rowBlob.Length;
        }

        private static void EnsureCapacity(ref byte[] buffer, int required)
        {
            if (buffer.Length >= required)
            {
                return;
            }
            Array.Resize(ref buffer, Math.Max(required, buffer.Length * 2));
        }

        /// <summary>
        /// Decodes every remaining row of the current sheet in one call. Unlike
        /// <see cref="NextRowDecoded"/>, end-of-sheet is not an error: it comes back as
        /// <see cref="NativeStatus.Ok"/> with <see cref="NativeRows.RowCount"/> equal to zero, since
        /// there's no per-call "keep going" signal here to distinguish EOF from an empty result.
        /// </summary>
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
                        // A real decode error mid-loop is a normal return, not a thrown exception —
                        // the catch block below won't run, so every row already decoded must be
                        // freed right here or it leaks.
                        FreeAll(decoded);
                        return status;
                    }
                    decoded.Add(row);
                }

                if (decoded.Count == 0)
                {
                    return NativeStatus.Ok;
                }

                int rowSize = Marshal.SizeOf<NativeRow>();
                IntPtr block = Marshal.AllocHGlobal(checked(decoded.Count * rowSize));
                for (int index = 0; index < decoded.Count; index++)
                {
                    Marshal.StructureToPtr(decoded[index], IntPtr.Add(block, index * rowSize), false);
                }

                rows = new NativeRows { RowCount = decoded.Count, Rows = block };
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                // Free whatever rows were already decoded before the failure — nothing leaks.
                FreeAll(decoded);
                SetLastError(exception.Message);
                rows = default;
                return NativeStatus.Error;
            }

            // Shared by both the mid-loop non-Ok/non-Eof return and the catch above, so the "free
            // everything decoded so far" behavior can't drift between the two paths.
            static void FreeAll(List<NativeRow> rowsToFree)
            {
                foreach (ref readonly NativeRow row in CollectionsMarshal.AsSpan(rowsToFree))
                {
                    NativeRow toFree = row;
                    FreeRow(ref toFree);
                }
            }
        }

        /// <summary>Releases a result returned by <see cref="ReadAllDecoded"/>. Safe on a zeroed value.</summary>
        internal static void FreeRows(ref NativeRows rows)
        {
            if (rows.Rows == IntPtr.Zero)
            {
                rows = default;
                return;
            }

            int rowSize = Marshal.SizeOf<NativeRow>();
            for (int index = 0; index < rows.RowCount; index++)
            {
                IntPtr rowPtr = IntPtr.Add(rows.Rows, index * rowSize);
                NativeRow row = Marshal.PtrToStructure<NativeRow>(rowPtr);
                FreeRow(ref row);
            }
            Marshal.FreeHGlobal(rows.Rows);
            rows = default;
        }
    }
}
