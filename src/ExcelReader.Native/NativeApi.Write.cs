using System.Diagnostics.CodeAnalysis;

namespace ExcelReader.Native
{
    internal static unsafe partial class NativeApi
    {
        /// <summary>
        /// Validates a caller-supplied write table before a single byte is written.
        /// </summary>
        /// <remarks>
        /// This is the trust boundary. Every pointer reachable from <paramref name="table"/> belongs to
        /// the caller and is about to be dereferenced against lengths the caller also supplied, so a
        /// value that is merely wrong here becomes an out-of-bounds read of the caller's process
        /// memory a few frames later. The string-offset walk below is the reason this runs to
        /// completion up front rather than checking each row as it is written: a partially written
        /// file plus a segfault is strictly worse than a rejected call.
        /// </remarks>
        /// <param name="hasHeader">True when every spec carries a name, so a header row must be written.</param>
        internal static bool TryValidateWriteTable(NativeColumnSpec[] specs, NativeTable table, out bool hasHeader, [NotNullWhen(false)] out string? error)
        {
            hasHeader = false;
            error = null;
            if (specs.Length == 0 || !IsValidSpecCount(table.ColumnCount))
            {
                error = $"xl_write_typed needs 1..{NativeLimits.MaxColumnSpecs} columns; got {table.ColumnCount}.";
                return false;
            }
            if (specs.Length != table.ColumnCount)
            {
                error = $"xl_write_typed got {specs.Length} spec(s) for {table.ColumnCount} column(s).";
                return false;
            }
            if (table.RowCount < 0 || table.Columns == IntPtr.Zero)
            {
                error = $"xl_write_typed needs a non-negative row_count and a non-NULL columns pointer; got {table.RowCount}.";
                return false;
            }

            hasHeader = specs[0].Name is not null;
            for (int index = 0; index < table.ColumnCount; index++)
            {
                if ((specs[index].Name is not null) != hasHeader)
                {
                    error = "every column spec must have a name, or none may — xl_write_typed cannot write a partial header row.";
                    return false;
                }
                if (!TryValidateWriteColumn(specs[index], ColumnAt(table, index), index, table.RowCount, out error))
                {
                    return false;
                }
            }
            return true;
        }

        // Split out of TryValidateWriteTable to keep both inside the style guide's nesting and length
        // limits; the string-offset walk alone is most of this method.
        private static bool TryValidateWriteColumn(NativeColumnSpec spec, NativeColumn column, int index, long rowCount, [NotNullWhen(false)] out string? error)
        {
            error = null;
            if (column.Type is < NativeColumnType.String or > NativeColumnType.Timestamp)
            {
                error = $"column {index} has unknown type {column.Type}.";
                return false;
            }
            if (spec.Type != column.Type)
            {
                error = $"column {index} has type {column.Type} but its spec says {spec.Type}.";
                return false;
            }
            if (column.Length != rowCount)
            {
                error = $"column {index} has length {column.Length}, but the table's row_count is {rowCount}.";
                return false;
            }
            if (rowCount > 0 && column.Values == IntPtr.Zero)
            {
                error = $"column {index} has {rowCount} row(s) but a NULL values pointer.";
                return false;
            }
            if (column.Type != NativeColumnType.String)
            {
                return true;
            }
            return TryValidateStringOffsets(column, index, rowCount, out error);
        }

        // The whole offsets array is walked here, once, before any of it is used as a slice bound.
        // Checking lazily per row would let a hostile offset reach `data` on the row before the one
        // that fails validation.
        private static bool TryValidateStringOffsets(NativeColumn column, int index, long rowCount, [NotNullWhen(false)] out string? error)
        {
            error = null;
            if (column.DataLen < 0 || (column.DataLen > 0 && column.Data == IntPtr.Zero))
            {
                error = $"column {index} has data_len {column.DataLen} but a NULL data pointer.";
                return false;
            }
            if (rowCount == 0)
            {
                return true;
            }

            int* offsets = (int*)column.Values;
            if (offsets[0] != 0)
            {
                error = $"column {index}: the first string offset must be 0, got {offsets[0]}.";
                return false;
            }
            for (long row = 0; row < rowCount; row++)
            {
                if (offsets[row + 1] < offsets[row])
                {
                    error = $"column {index}: string offset {row + 1} ({offsets[row + 1]}) is before offset {row} ({offsets[row]}).";
                    return false;
                }
            }
            if (offsets[rowCount] != column.DataLen)
            {
                error = $"column {index}: the last string offset is {offsets[rowCount]}, but data_len is {column.DataLen}.";
                return false;
            }
            return true;
        }
    }
}
