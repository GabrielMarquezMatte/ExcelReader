using System.Runtime.InteropServices;

namespace ExcelReader.Native
{
    internal static class ArrowFlags
    {
        internal const long Nullable = 2;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct ArrowSchema
    {
        public IntPtr Format;
        public IntPtr Name;
        public IntPtr Metadata;
        public long Flags;
        public long NChildren;
        public IntPtr Children;
        public IntPtr Dictionary;
        public IntPtr Release;
        public IntPtr PrivateData;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct ArrowArray
    {
        public long Length;
        public long NullCount;
        public long Offset;
        public long NBuffers;
        public long NChildren;
        public IntPtr Buffers;
        public IntPtr Children;
        public IntPtr Dictionary;
        public IntPtr Release;
        public IntPtr PrivateData;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct ArrowArrayStream
    {
        public IntPtr GetSchema;
        public IntPtr GetNext;
        public IntPtr GetLastError;
        public IntPtr Release;
        public IntPtr PrivateData;
    }
}
