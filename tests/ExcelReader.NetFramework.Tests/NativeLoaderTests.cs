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
            var ex = Assert.Throws<InvalidOperationException>(() => NativeLoader.CheckAbi(5));
            Assert.Contains("6", ex.Message, StringComparison.Ordinal);
            Assert.Contains("5", ex.Message, StringComparison.Ordinal);
            NativeLoader.CheckAbi(6);
        }

        [Fact]
        public void NativeLibraryLoadsAndReportsAbi() => Assert.Equal(6, NativeMethods.xl_abi_version());

        [Fact]
        public void MissingDllMessageNamesTheArchitecture()
        {
            var ex = Assert.Throws<PlatformNotSupportedException>(() => NativeLoader.Load(@"C:\nowhere\x86\ExcelReader.Native.dll"));
            Assert.Contains("x86", ex.Message, StringComparison.Ordinal);
            Assert.Contains("PlatformTarget", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void NativeErrorCarriesStatusAndMessage()
        {
            var ex = NativeError.Create((int)ExcelStatus.InvalidHandle);
            Assert.Equal(ExcelStatus.InvalidHandle, ex.Status);
            Assert.False(string.IsNullOrEmpty(ex.Message));
        }

        [Fact]
        public void ThrowIfFailedPassesOk() => NativeError.ThrowIfFailed(0);
    }
}
