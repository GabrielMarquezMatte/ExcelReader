using System.Runtime.InteropServices;
using ExcelReader.NetFramework.Reader;

namespace ExcelReader.NetFramework.Interop
{
    internal static unsafe class NativeBuffers
    {
        internal static byte[] TakeArray(ref XlBuffer buffer)
        {
            try
            {
                if (buffer.Len > int.MaxValue)
                {
                    throw new ExcelReaderException(ExcelStatus.Error, "The native buffer is larger than 2 GiB.");
                }

                byte[] result = new byte[(int)buffer.Len];
                if (result.Length > 0)
                {
                    Marshal.Copy((IntPtr)buffer.Data, result, 0, result.Length);
                }

                return result;
            }
            finally
            {
                fixed (XlBuffer* pinned = &buffer)
                {
                    NativeMethods.xl_free_buffer(pinned);
                }
            }
        }
    }
}
