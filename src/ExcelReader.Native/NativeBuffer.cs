using System.Runtime.InteropServices;

namespace ExcelReader.Native
{
    /// <summary>
    /// Flat C ABI representation of <c>xl_buffer</c>: an owned block of unmanaged memory returned by
    /// <c>xl_write_typed_to_memory</c> or <c>xl_write_handle_bytes</c>. The caller must release it
    /// with <c>xl_free_buffer</c> — same ownership convention as <see cref="NativeTable"/> and
    /// <see cref="NativeInferredSchema"/>.
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeBuffer
    {
        public IntPtr Data;
        public long Length;
    }
}
