using System.Runtime.InteropServices;

namespace ExcelReader.Native
{
    // Arrow schema flag bits. Mirrors ARROW_FLAG_* from the Arrow C Data Interface spec.
    internal static class ArrowFlags
    {
        internal const long Nullable = 2;
    }

    // C# mirror of the Arrow C Data Interface's struct ArrowSchema (see excelreader_arrow.h).
    // Every pointer-shaped field is declared as IntPtr rather than a raw pointer type —
    // layout-identical to the real C struct (an IntPtr and a native pointer occupy the
    // same bytes), but this keeps the type usable from ordinary managed code (including this project's
    // test suite, which has no AllowUnsafeBlocks) without forcing every caller into an unsafe
    // context. Release holds a native function pointer's bit pattern, computed with
    // delegate* syntax only where the struct is actually built (NativeApi.Arrow.cs,
    // declared unsafe) and invoked only from native code — never called back into managed code
    // directly, so this field never needs to be anything more than an opaque address here.
    [StructLayout(LayoutKind.Sequential)]
    internal struct ArrowSchema
    {
        public IntPtr Format;
        public IntPtr Name;
        public IntPtr Metadata;
        public long Flags;
        public long NChildren;
        public IntPtr Children;   // ArrowSchema**
        public IntPtr Dictionary; // ArrowSchema*
        public IntPtr Release;    // void (*)(ArrowSchema*)
        public IntPtr PrivateData;
    }

    // C# mirror of the Arrow C Data Interface's struct ArrowArray. See ArrowSchema's
    // remarks for why every pointer field is IntPtr.
    [StructLayout(LayoutKind.Sequential)]
    internal struct ArrowArray
    {
        public long Length;
        public long NullCount;
        public long Offset;
        public long NBuffers;
        public long NChildren;
        public IntPtr Buffers;    // const void**
        public IntPtr Children;   // ArrowArray**
        public IntPtr Dictionary; // ArrowArray*
        public IntPtr Release;    // void (*)(ArrowArray*)
        public IntPtr PrivateData;
    }

    // C# mirror of the Arrow C Data Interface's struct ArrowArrayStream. See
    // ArrowSchema's remarks for why every pointer field is IntPtr.
    [StructLayout(LayoutKind.Sequential)]
    internal struct ArrowArrayStream
    {
        public IntPtr GetSchema;    // int (*)(ArrowArrayStream*, ArrowSchema*)
        public IntPtr GetNext;      // int (*)(ArrowArrayStream*, ArrowArray*)
        public IntPtr GetLastError; // const char* (*)(ArrowArrayStream*)
        public IntPtr Release;      // void (*)(ArrowArrayStream*)
        public IntPtr PrivateData;
    }
}
