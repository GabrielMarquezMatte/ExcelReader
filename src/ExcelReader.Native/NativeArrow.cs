using System.Runtime.InteropServices;

namespace ExcelReader.Native
{
    /// <summary>Arrow schema flag bits. Mirrors ARROW_FLAG_* from the Arrow C Data Interface spec.</summary>
    internal static class ArrowFlags
    {
        internal const long Nullable = 2;
    }

    /// <summary>
    /// C# mirror of the Arrow C Data Interface's <c>struct ArrowSchema</c> (see excelreader_arrow.h).
    /// Every pointer-shaped field is declared as <see cref="IntPtr"/> rather than a raw pointer type —
    /// layout-identical to the real C struct (an <see cref="IntPtr"/> and a native pointer occupy the
    /// same bytes), but this keeps the type usable from ordinary managed code (including this project's
    /// test suite, which has no <c>AllowUnsafeBlocks</c>) without forcing every caller into an unsafe
    /// context. <see cref="Release"/> holds a native function pointer's bit pattern, computed with
    /// <c>delegate*</c> syntax only where the struct is actually built (<c>NativeApi.Arrow.cs</c>,
    /// declared unsafe) and invoked only from native code — never called back into managed code
    /// directly, so this field never needs to be anything more than an opaque address here.
    /// </summary>
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

    /// <summary>C# mirror of the Arrow C Data Interface's <c>struct ArrowArray</c>. See <see cref="ArrowSchema"/>'s
    /// remarks for why every pointer field is <see cref="IntPtr"/>.</summary>
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
}
