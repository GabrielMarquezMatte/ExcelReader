using System.Text;
using ExcelReader.NetFramework.Interop;

namespace ExcelReader.NetFramework.Tests.Interop
{
    public sealed unsafe class WriterHandleTests
    {
        private static WriterHandle OpenMemory()
        {
            XlWriteOptions options = default;
            options.StructSize = sizeof(XlWriteOptions);
            IntPtr raw = IntPtr.Zero;
            NativeError.ThrowIfFailed(NativeMethods.xl_open_write_handle_to_memory(Xl.XL_FORMAT_XLSX, &options, &raw));
            return new WriterHandle(raw);
        }

        [Fact]
        public void OpenInMemoryAndAbort()
        {
            NativeLoader.EnsureLoaded();
            WriterHandle handle = OpenMemory();
            IntPtr raw = handle.DangerousGetHandle();

            handle.Dispose();

            Assert.True(handle.IsClosed);
            byte[] name = Encoding.UTF8.GetBytes("S");
            fixed (byte* namePointer = name)
            {
                Assert.Equal(-3, NativeMethods.xl_start_sheet(raw, namePointer, name.Length));
            }
        }

        [Fact]
        public void MemoryRoundTripBytes()
        {
            NativeLoader.EnsureLoaded();
            using (WriterHandle handle = OpenMemory())
            {
                IntPtr raw = handle.DangerousGetHandle();
                byte[] name = Encoding.UTF8.GetBytes("S");
                byte[] value = Encoding.UTF8.GetBytes("x");
                fixed (byte* namePointer = name)
                fixed (byte* valuePointer = value)
                {
                    NativeError.ThrowIfFailed(NativeMethods.xl_start_sheet(raw, namePointer, name.Length));
                    NativeError.ThrowIfFailed(NativeMethods.xl_start_row(raw));
                    NativeError.ThrowIfFailed(NativeMethods.xl_write_string(raw, valuePointer, value.Length));
                }

                NativeError.ThrowIfFailed(NativeMethods.xl_end_row(raw));
                NativeError.ThrowIfFailed(NativeMethods.xl_end_sheet(raw));

                XlBuffer buffer = default;
                NativeError.ThrowIfFailed(NativeMethods.xl_write_handle_bytes(raw, &buffer));
                byte[] bytes = NativeBuffers.TakeArray(ref buffer);

                Assert.True(bytes.Length > 2);
                Assert.Equal(0x50, bytes[0]);
                Assert.Equal(0x4B, bytes[1]);
                NativeError.ThrowIfFailed(NativeMethods.xl_close_write_handle(raw));
                handle.MarkClosed();
            }
        }
    }
}
