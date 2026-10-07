using System.Runtime.InteropServices;

namespace ExcelReader.Native.Reading
{
    [StructLayout(LayoutKind.Sequential)]
    internal unsafe struct NativeSourceRaw
    {
        public int StructSize;
        public void* UserData;
        public long Length;
        public delegate* unmanaged<void*, long, byte*, long, long> ReadAt;
        public delegate* unmanaged<void*, void> Release;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal unsafe struct NativeStreamRaw
    {
        public int StructSize;
        public void* UserData;
        public delegate* unmanaged<void*, byte*, long, long> Read;
        public delegate* unmanaged<void*, void> Release;
    }
}
