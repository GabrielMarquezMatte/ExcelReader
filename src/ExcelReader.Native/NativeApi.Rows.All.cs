using System.Buffers.Binary;
using ExcelReader.Core.ValueObjects;
using System.Runtime.InteropServices;

namespace ExcelReader.Native
{
    internal static unsafe partial class NativeApi
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

                // The leading row count is written here rather than back-patched into the accumulated
                // bytes: the accumulator is chunked and its first chunk is not addressable as "the
                // start of the blob" the way a single array's offset 0 was.
                BinaryPrimitives.WriteInt32LittleEndian(buffer, handle.AllRowsCount);
                handle.AllRowsScratch?.CopyTo(buffer[sizeof(int)..handle.AllRowsLength]);
                handle.AllRowsPending = false;
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                handle.AllRowsPending = false;
                handle.AllRowsLength = 0;
                handle.AllRowsCount = 0;
                // Dropped, not kept: a partially accumulated buffer can never be handed out, and on a
                // sheet big enough to have thrown it is the largest thing this handle is holding.
                handle.AllRowsScratch = null;
                return NativeStatus.Error;
            }
        }

        // Drains every remaining row of the current sheet (including one already held pending from a
        // prior xl_next_row that returned XL_BUFFER_TOO_SMALL) into handle.AllRowsScratch as
        //     repeated: int32 row_length, <row blob>
        // with the blob's leading int32 row_count carried separately on the handle and written out in
        // ReadAllBlob. The underlying enumerator is fully consumed by the time this returns — a
        // too-small buffer on the ReadAllBlob call above does not mean rows remain unread, only that
        // the accumulated result hasn't been copied out to the caller yet.
        //
        // Accumulates into a ChunkedBuffer rather than a byte[] grown with Array.Resize: a doubling
        // resize copies everything accumulated so far and discards the old array, so a 22 MB sheet
        // blob cost roughly another 22 MB of copying and LOH garbage on top of itself. Chunks are
        // never copied until the single CopyTo into the caller's buffer.
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

        // Appends one `int32 row_length` + `row blob` entry to `output`.
        private static void AppendRow(ChunkedBuffer<byte> output, ReadOnlySpan<byte> rowBlob)
        {
            // Lengths are int32 throughout this API (see excelreader.h on xl_read_all_blob), so the
            // next required size is computed in a wider type first — including the leading row count
            // the caller's buffer must also hold — and reported precisely, rather than surfacing as
            // an overflow deep inside the accumulator.
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

                // NativeRow is blittable (an int and a pointer), so the array is filled by plain
                // stores rather than Marshal.StructureToPtr — that goes through a marshalling stub
                // per element, and this loop runs once per row of the sheet.
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
