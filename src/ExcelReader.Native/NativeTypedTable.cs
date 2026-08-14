using System.Runtime.InteropServices;

namespace ExcelReader.Native
{
    /// <summary>Column value types accepted by <c>xl_parse_typed</c>. Mirrors XL_T_* in include/excelreader.h.</summary>
    internal static class NativeColumnType
    {
        internal const int String = 0;
        internal const int Int64 = 1;
        internal const int Float64 = 2;
        internal const int Bool = 3;
        /// <summary>Days since 1970-01-01, stored as a 4-byte value.</summary>
        internal const int Date = 4;
        /// <summary>Microseconds since midnight, stored as an 8-byte value.</summary>
        internal const int Time = 5;
        /// <summary>Microseconds since 1970-01-01T00:00:00Z, stored as an 8-byte value.</summary>
        internal const int Timestamp = 6;
    }

    /// <summary>
    /// Flat C ABI representation for one raw <c>xl_column_spec</c> as received across the boundary — the
    /// <see cref="Name"/> pointer is only valid for the duration of the call, so <see cref="Exports"/>
    /// decodes it into the UTF-8-decoded <see cref="NativeColumnSpec"/> before calling into
    /// <see cref="NativeApi"/>.
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    internal unsafe struct NativeColumnSpecRaw
    {
        public byte* Name;
        public int NameLen;
        public int Index;
        public int Type;
        public int Nullable;
    }

    /// <summary>Decoded, pointer-free form of <see cref="NativeColumnSpecRaw"/> — the layer
    /// <see cref="NativeApi"/> and its tests actually work with.</summary>
    internal readonly struct NativeColumnSpec
    {
        /// <summary>The header text to match (case-insensitively, trimmed), or <see langword="null"/> to
        /// resolve by <see cref="Index"/> instead.</summary>
        internal string? Name { get; init; }
        internal int Index { get; init; }
        internal int Type { get; init; }
        internal bool Nullable { get; init; }
    }

    /// <summary>
    /// Flat C ABI representation of one output column. <see cref="Values"/> is the only allocation this
    /// column owns directly: for <see cref="NativeColumnType.String"/> it holds the int32 offsets array
    /// followed immediately by the UTF-8 data blob in ONE block (<see cref="Data"/> is an interior
    /// pointer into it, same arena pattern as the decoded row in NativeApi.Rows.cs) — freeing
    /// <see cref="Data"/> separately would be a double free. <see cref="Validity"/>, when non-null, is a
    /// second, independent allocation.
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeColumn
    {
        public int Type;
        public long Length;
        public IntPtr Values;
        public IntPtr Validity;
        public IntPtr Data;
        public long DataLen;
    }

    /// <summary>Flat C ABI representation of the whole result. <see cref="Columns"/> is one allocation
    /// holding <see cref="ColumnCount"/> <see cref="NativeColumn"/> values; each column's own
    /// <see cref="NativeColumn.Values"/>/<see cref="NativeColumn.Validity"/> are separate allocations
    /// freed individually by <see cref="NativeApi.FreeTable"/>.</summary>
    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeTable
    {
        public int ColumnCount;
        public long RowCount;
        public IntPtr Columns;
    }
}
