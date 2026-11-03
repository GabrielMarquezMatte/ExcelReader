using System.Globalization;
using System.Runtime.InteropServices;

namespace ExcelReader.NetFramework.Interop
{
    internal static class NativeLoader
    {
        internal const int AbiVersion = 6;

        private const string DllName = "ExcelReader.Native.dll";
        private static readonly object Gate = new();
        private static bool loaded;

        [DllImport("kernel32", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "LoadLibraryW", ExactSpelling = true)]
        private static extern IntPtr LoadLibraryW(string fileName);

        internal static void EnsureLoaded()
        {
            lock (Gate)
            {
                if (loaded) return;
                Load(FindDll());
                loaded = true;
                CheckAbi(NativeMethods.xl_abi_version());
            }
        }

        internal static string ResolvePath(string baseDirectory, bool is64BitProcess)
            => Path.Combine(baseDirectory, is64BitProcess ? "x64" : "x86", DllName);

        internal static void Load(string fullPath)
        {
            if (LoadLibraryW(fullPath) != IntPtr.Zero) return;
            var error = Marshal.GetLastWin32Error();
            var folder = Path.GetFileName(Path.GetDirectoryName(fullPath));
            var arch = Environment.Is64BitProcess ? "x64" : "x86";
            throw new PlatformNotSupportedException(
                $"Could not load the native ExcelReader library '{fullPath}' (Win32 error {error.ToString(CultureInfo.InvariantCulture)}). "
                + $"This is a {arch} process and the library is expected in a '{folder}' folder beside the application. "
                + "Set PlatformTarget to x64 or turn off Prefer 32-bit, and deploy the matching ExcelReader.Native.dll.");
        }

        internal static void CheckAbi(int actual)
        {
            if (actual == AbiVersion) return;
            throw new InvalidOperationException(
                $"ExcelReader.Native reports ABI version {actual.ToString(CultureInfo.InvariantCulture)}, but this package requires {AbiVersion.ToString(CultureInfo.InvariantCulture)}.");
        }

        private static string FindDll()
        {
            var is64Bit = Environment.Is64BitProcess;
            var primary = ResolvePath(AppDomain.CurrentDomain.BaseDirectory, is64Bit);
            if (File.Exists(primary)) return primary;
            var assemblyDirectory = Path.GetDirectoryName(typeof(NativeLoader).Assembly.Location);
            if (assemblyDirectory is null) return primary;
            var beside = ResolvePath(assemblyDirectory, is64Bit);
            return File.Exists(beside) ? beside : primary;
        }
    }
}
