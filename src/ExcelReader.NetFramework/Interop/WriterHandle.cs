using Microsoft.Win32.SafeHandles;

namespace ExcelReader.NetFramework.Interop
{
    internal sealed class WriterHandle : SafeHandleZeroOrMinusOneIsInvalid
    {
        private bool _closed;

        internal WriterHandle(IntPtr raw) : base(true)
        {
            SetHandle(raw);
        }

        internal void MarkClosed()
        {
            _closed = true;
        }

        protected override bool ReleaseHandle()
        {
            if (_closed)
            {
                return true;
            }

            _ = NativeMethods.xl_abort_write_handle(handle);
            return true;
        }
    }
}
