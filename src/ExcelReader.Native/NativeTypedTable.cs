using System.Runtime.InteropServices;

namespace ExcelReader.Native
{
    // Column value types accepted by xl_parse_typed. Mirrors XL_T_* in include/excelreader.h.
    internal static class NativeColumnType
    {
        internal const int String = 0;
        internal const int Int64 = 1;
        internal const int Float64 = 2;
        internal const int Bool = 3;
        // Days since 1970-01-01, stored as a 4-byte value.
        internal const int Date = 4;
        // Microseconds since midnight, stored as an 8-byte value.
        internal const int Time = 5;
        // Microseconds since 1970-01-01T00:00:00Z, stored as an 8-byte value.
        internal const int Timestamp = 6;
    }

    // Flat C ABI representation for one raw xl_column_spec as received across the boundary — the
    // Names pointers are only valid for the duration of the call, so Exports
    // decodes them into the UTF-8-decoded NativeColumnSpec before calling into
    // NativeApi.
    [StructLayout(LayoutKind.Sequential)]
    internal unsafe struct NativeColumnSpecRaw
    {
        public byte** Names;
        public int* NameLens;
        public int NameCount;
        public int Index;
        public int Type;
        public int Nullable;
    }

    // Decoded, pointer-free form of NativeColumnSpecRaw — the layer
    // NativeApi and its tests actually work with.
    internal readonly struct NativeColumnSpec
    {
        public NativeColumnSpec()
        {
        }

        // Candidate header texts to match (case-insensitively, trimmed), tried in order — the
        // first one present in the header row wins. Empty to resolve by Index instead.
        internal string[] Names { get; init; } = [];
        internal int Index { get; init; }
        internal int Type { get; init; }
        internal bool Nullable { get; init; }
    }

    // Flat C ABI representation of one output column. Values is the only allocation this
    // column owns directly: for NativeColumnType.String it holds the int32 offsets array
    // followed immediately by the UTF-8 data blob in ONE block (Data is an interior
    // pointer into it, same arena pattern as the decoded row in NativeApi.Rows.cs) — freeing
    // Data separately would be a double free. Validity, when non-null, is a
    // second, independent allocation.
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

    // Flat C ABI representation of the whole result. Columns is one allocation
    // holding ColumnCount NativeColumn values; each column's own
    // NativeColumn.Values/NativeColumn.Validity are separate allocations
    // freed individually by NativeApi.FreeTable.
    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeTable
    {
        public int ColumnCount;
        public long RowCount;
        public IntPtr Columns;
    }
}
