using System.Runtime.InteropServices;

namespace ExcelReader.Native
{
    internal static class NativeColumnType
    {
        internal const int String = 0;
        internal const int Int64 = 1;
        internal const int Float64 = 2;
        internal const int Bool = 3;
        internal const int Date = 4;
        internal const int Time = 5;
        internal const int Timestamp = 6;
    }

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

    internal readonly struct NativeColumnSpec
    {
        public NativeColumnSpec()
        {
        }

        internal string[] Names { get; init; } = [];
        internal int Index { get; init; }
        internal int Type { get; init; }
        internal bool Nullable { get; init; }
    }

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

    [StructLayout(LayoutKind.Sequential)]
    internal unsafe struct NativeTable
    {
        public int ColumnCount;
        public long RowCount;
        public IntPtr Columns;

        internal readonly NativeColumn ColumnAt(int index)
        {
            return ((NativeColumn*)Columns)[index];
        }
    }
}
