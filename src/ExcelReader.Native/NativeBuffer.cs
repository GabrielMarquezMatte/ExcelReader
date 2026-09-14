using System.Runtime.InteropServices;

namespace ExcelReader.Native
{
    // Flat C ABI representation of xl_buffer: an owned block of unmanaged memory returned by
    // xl_write_typed_to_memory or xl_write_handle_bytes. The caller must release it
    // with xl_free_buffer — same ownership convention as NativeTable and
    // NativeInferredSchema.
    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeBuffer
    {
        public IntPtr Data;
        public long Length;
    }
}
