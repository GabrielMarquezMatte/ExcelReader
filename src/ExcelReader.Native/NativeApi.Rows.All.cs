using System.Runtime.InteropServices;

namespace ExcelReader.Native
{
    internal static partial class NativeApi
    {
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
