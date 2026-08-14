using System.Runtime.InteropServices;

namespace ExcelReader.Native
{
    /// <summary>
    /// Flat C ABI representation for a decoded row. Strings are copied into native memory owned by
    /// the row allocation; call <see cref="NativeApi.FreeRow(ref NativeRow)"/> when done.
    /// </summary>
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
}
