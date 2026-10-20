using System.Runtime.InteropServices;

namespace ExcelReader.Native.Reading
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeRowCell
    {
        public int Column;
        public int Type;
        public int ValueLength;
        public IntPtr Value;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal unsafe struct NativeCellView
    {
        public int Column;
        public int Type;
        public int ValueLength;
        public int HasNumber;
        public byte* Value;
        public double Number;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeRow
    {
        public int CellCount;
        public IntPtr Cells;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeRows
    {
        public int RowCount;
        public IntPtr Rows;
    }
}
