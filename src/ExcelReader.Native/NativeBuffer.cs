using System.Runtime.InteropServices;

namespace ExcelReader.Native
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeBuffer
    {
        public IntPtr Data;
        public long Length;
    }
}
