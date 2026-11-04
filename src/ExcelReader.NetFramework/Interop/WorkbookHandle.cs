using ExcelReader.NetFramework.Reader;
using Microsoft.Win32.SafeHandles;

namespace ExcelReader.NetFramework.Interop
{
    internal sealed class WorkbookHandle : SafeHandleZeroOrMinusOneIsInvalid
    {
        internal WorkbookHandle(IntPtr handle) : base(true)
        {
            SetHandle(handle);
        }

        protected override bool ReleaseHandle()
        {
            return NativeMethods.xl_close(handle) == (int)ExcelStatus.Ok;
        }
    }
}
