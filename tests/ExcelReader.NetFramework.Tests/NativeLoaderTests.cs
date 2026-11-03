using System.Text;
using ExcelReader.NetFramework.Interop;

namespace ExcelReader.NetFramework.Tests
{
    public sealed class NativeLoaderTests
    {
        [Fact]
        public void ResolvePathPicksArchFolder()
        {
            Assert.Equal(@"C:\app\x64\ExcelReader.Native.dll", NativeLoader.ResolvePath(@"C:\app", true));
            Assert.Equal(@"C:\app\x86\ExcelReader.Native.dll", NativeLoader.ResolvePath(@"C:\app", false));
        }

        [Fact]
        public void CheckAbiRejectsMismatch()
        {
            InvalidOperationException ex = Assert.Throws<InvalidOperationException>(() => NativeLoader.CheckAbi(5));
            Assert.Contains("6", ex.Message, StringComparison.Ordinal);
            Assert.Contains("5", ex.Message, StringComparison.Ordinal);
            NativeLoader.CheckAbi(6);
        }

        [Fact]
        public void NativeLibraryLoadsAndReportsAbi()
        {
            Assert.Equal(6, NativeMethods.xl_abi_version());
        }

        [Fact]
        public void MissingDllMessageNamesTheArchitecture()
        {
            PlatformNotSupportedException ex = Assert.Throws<PlatformNotSupportedException>(() => NativeLoader.Load(@"C:\nowhere\x86\ExcelReader.Native.dll"));
            Assert.Contains("x86", ex.Message, StringComparison.Ordinal);
            Assert.Contains("PlatformTarget", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void CreateFallsBackToStatusNameWithoutNativeMessage()
        {
            Assert.Equal("Native call failed with status Error.", NativeError.Create((int)ExcelStatus.Error, string.Empty).Message);
            Assert.Equal("Native call failed with status PasswordRequired.", NativeError.Create((int)ExcelStatus.PasswordRequired, null).Message);
        }

        [Fact]
        public void CreateKeepsTheNativeMessage()
        {
            ExcelReaderException ex = NativeError.Create((int)ExcelStatus.InvalidArgument, "bad argument");
            Assert.Equal(ExcelStatus.InvalidArgument, ex.Status);
            Assert.Equal("bad argument", ex.Message);
        }

        [Fact]
        public void ThrowIfFailedThrowsWithTheStatus()
        {
            ExcelReaderException ex = Assert.Throws<ExcelReaderException>(() => NativeError.ThrowIfFailed((int)ExcelStatus.InvalidHandle));
            Assert.Equal(ExcelStatus.InvalidHandle, ex.Status);
        }

        [Fact]
        public unsafe void NativeFailureCarriesTheNativeMessage()
        {
            byte[] path = Encoding.UTF8.GetBytes(Fixtures.Path("missing.xlsx"));
            IntPtr handle = IntPtr.Zero;
            int status;
            fixed (byte* pathPointer = path)
            {
                status = NativeMethods.xl_open_file(pathPointer, path.Length, Xl.XL_FORMAT_AUTO, null, &handle);
            }

            ExcelReaderException ex = Assert.Throws<ExcelReaderException>(() => NativeError.ThrowIfFailed(status));
            Assert.NotEqual(ExcelStatus.Ok, ex.Status);
            Assert.Equal((ExcelStatus)status, ex.Status);
            Assert.False(string.IsNullOrWhiteSpace(ex.Message));
            Assert.False(ex.Message.StartsWith("Native call failed", StringComparison.Ordinal), ex.Message);
        }

        [Fact]
        public unsafe void ThrowIfFailedAcceptsARealSuccessfulCall()
        {
            byte[] path = Encoding.UTF8.GetBytes(Fixtures.Path("RealExcel.xlsx"));
            IntPtr handle = IntPtr.Zero;
            int status;
            fixed (byte* pathPointer = path)
            {
                status = NativeMethods.xl_open_file(pathPointer, path.Length, Xl.XL_FORMAT_AUTO, null, &handle);
            }

            NativeError.ThrowIfFailed(status);
            Assert.NotEqual(IntPtr.Zero, handle);
            NativeError.ThrowIfFailed(NativeMethods.xl_close(handle));
        }
    }
}
