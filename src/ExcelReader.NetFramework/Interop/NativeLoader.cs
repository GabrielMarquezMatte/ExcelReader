using System.Globalization;
using System.Runtime.InteropServices;

namespace ExcelReader.NetFramework.Interop
{
    internal static class NativeLoader
    {
        internal const int AbiVersion = 6;

        private const string DllName = "ExcelReader.Native.dll";

        [DllImport("kernel32", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "LoadLibraryW", ExactSpelling = true)]
        private static extern IntPtr LoadLibraryW(string fileName);

        // Called only from the NativeMethods static constructor, which the CLR runs once.
        internal static void EnsureLoaded()
        {
            Load(FindDll());
            CheckAbi(NativeMethods.xl_abi_version());
        }

        internal static string ResolvePath(string baseDirectory, bool is64BitProcess)
        {
            return Path.Combine(baseDirectory, is64BitProcess ? "x64" : "x86", DllName);
        }

        internal static void Load(string fullPath)
        {
            if (LoadLibraryW(fullPath) != IntPtr.Zero)
            {
                return;
            }

            int error = Marshal.GetLastWin32Error();
            string? folder = Path.GetFileName(Path.GetDirectoryName(fullPath));
            string arch = Environment.Is64BitProcess ? "x64" : "x86";
            throw new PlatformNotSupportedException(
                $"Could not load the native ExcelReader library '{fullPath}' (Win32 error {error.ToString(CultureInfo.InvariantCulture)}). "
                + $"This is a {arch} process and the library is expected in a '{folder}' folder beside the application. "
                + "Set PlatformTarget to x64 or turn off Prefer 32-bit, and deploy the matching ExcelReader.Native.dll.");
        }

        internal static void CheckAbi(int actual)
        {
            if (actual == AbiVersion)
            {
                return;
            }

            throw new InvalidOperationException(
                $"ExcelReader.Native reports ABI version {actual.ToString(CultureInfo.InvariantCulture)}, but this package requires {AbiVersion.ToString(CultureInfo.InvariantCulture)}.");
        }

        private static string FindDll()
        {
            bool is64Bit = Environment.Is64BitProcess;
            string primary = ResolvePath(AppDomain.CurrentDomain.BaseDirectory, is64Bit);
            if (File.Exists(primary))
            {
                return primary;
            }

            string? assemblyDirectory = Path.GetDirectoryName(typeof(NativeLoader).Assembly.Location);
            if (assemblyDirectory is null)
            {
                return primary;
            }

            string beside = ResolvePath(assemblyDirectory, is64Bit);
            if (File.Exists(beside))
            {
                return beside;
            }

            return primary;
        }
    }
}
