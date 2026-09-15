using System.Runtime.InteropServices;

namespace ExcelReader.Native
{
    // Flat C ABI representation for a decoded row. Strings are copied into native memory owned by
    // the row allocation; call NativeApi.FreeRow(ref NativeRow) when done.
    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeRowCell
    {
        public int Column;
        public int Type;
        public int ValueLength;
        public IntPtr Value;
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
