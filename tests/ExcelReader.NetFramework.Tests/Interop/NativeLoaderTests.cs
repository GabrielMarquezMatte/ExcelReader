using System.Text;
using ExcelReader.NetFramework.Interop;
using ExcelReader.NetFramework.Reader;

namespace ExcelReader.NetFramework.Tests.Interop
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
            Assert.Contains("7", ex.Message, StringComparison.Ordinal);
            Assert.Contains("5", ex.Message, StringComparison.Ordinal);
            NativeLoader.CheckAbi(7);
        }

        [Fact]
        public void NativeLibraryLoadsAndReportsAbi()
        {
            NativeLoader.EnsureLoaded();
            Assert.Equal(7, NativeMethods.xl_abi_version());
        }

        [Fact]
        public void ProbePathsCoverShadowCopyHosts()
        {
            string[] expected =
            {
                @"C:\site\x64\ExcelReader.Native.dll",
                @"C:\site\bin\x64\ExcelReader.Native.dll",
                @"C:\site\lib\x64\ExcelReader.Native.dll",
                @"C:\Temp\shadow\ab12\x64\ExcelReader.Native.dll",
            };
            IReadOnlyList<string> paths = NativeLoader.ProbePaths(
                @"C:\site\",
                "bin; lib;",
                "file:///C:/Site/Bin/ExcelReader.NetFramework.DLL",
                @"C:\Temp\shadow\ab12\ExcelReader.NetFramework.dll",
                true);
            Assert.Equal(expected, paths);
        }

        [Fact]
        public void ProbePathsDecodeEscapedCodeBaseWithHashPercentAndSpace()
        {
            string[] expected =
            {
                @"C:\app\x64\ExcelReader.Native.dll",
                @"C:\My C#%20Dir\bin\x64\ExcelReader.Native.dll",
            };
            IReadOnlyList<string> paths = NativeLoader.ProbePaths(
                @"C:\app\",
                null,
                "file:///C:/My%20C%23%2520Dir/bin/ExcelReader.NetFramework.dll",
                string.Empty,
                true);
            Assert.Equal(expected, paths);
        }

        [Fact]
        public void ProbePathsSkipMissingInputs()
        {
            string[] expected = { @"C:\app\x86\ExcelReader.Native.dll" };
            Assert.Equal(expected, NativeLoader.ProbePaths(@"C:\app\", null, null, string.Empty, false));
            Assert.Equal(expected, NativeLoader.ProbePaths(@"C:\app\", null, "file:///C:/app/ExcelReader.NetFramework.dll", @"C:\app\ExcelReader.NetFramework.dll", false));
        }

        [Fact]
        public void NotFoundMessageListsEveryProbedPath()
        {
            string[] probed = { @"C:\one\x64\ExcelReader.Native.dll", @"C:\two\x64\ExcelReader.Native.dll" };
            PlatformNotSupportedException ex = NativeLoader.NotFound(probed);
            Assert.Contains(probed[0], ex.Message, StringComparison.Ordinal);
            Assert.Contains(probed[1], ex.Message, StringComparison.Ordinal);
            Assert.Contains("PlatformTarget", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void LoadFailureIsCachedAndRethrownWithItsOwnType()
        {
            int calls = 0;
            LoadGuard guard = new LoadGuard(() =>
            {
                calls++;
                throw new PlatformNotSupportedException("no native library");
            });
            PlatformNotSupportedException first = Assert.Throws<PlatformNotSupportedException>(guard.Ensure);
            PlatformNotSupportedException second = Assert.Throws<PlatformNotSupportedException>(guard.Ensure);
            Assert.Equal("no native library", first.Message);
            Assert.Equal(first.Message, second.Message);
            Assert.Equal(1, calls);
        }

        [Fact]
        public void ConcurrentCallersLoadOnce()
        {
            int calls = 0;
            LoadGuard guard = new LoadGuard(() =>
            {
                Interlocked.Increment(ref calls);
                Thread.Sleep(50);
            });
            using Barrier barrier = new Barrier(8);
            Thread[] threads = new Thread[8];
            for (int i = 0; i < threads.Length; i++)
            {
                threads[i] = new Thread(() =>
                {
                    barrier.SignalAndWait();
                    guard.Ensure();
                });
                threads[i].Start();
            }

            foreach (Thread thread in threads)
            {
                thread.Join();
            }

            Assert.Equal(1, calls);
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
            NativeLoader.EnsureLoaded();
            ExcelReaderException ex = Assert.Throws<ExcelReaderException>(() => NativeError.ThrowIfFailed((int)ExcelStatus.InvalidHandle));
            Assert.Equal(ExcelStatus.InvalidHandle, ex.Status);
        }

        [Fact]
        public unsafe void NativeFailureCarriesTheNativeMessage()
        {
            NativeLoader.EnsureLoaded();
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
            NativeLoader.EnsureLoaded();
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
