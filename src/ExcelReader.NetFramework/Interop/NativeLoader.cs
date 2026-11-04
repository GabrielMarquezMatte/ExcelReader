using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelReader.NetFramework.Interop
{
    internal static class NativeLoader
    {
        internal const int AbiVersion = 7;

        private const string DllName = "ExcelReader.Native.dll";

        private static readonly char[] SearchPathSeparator = { ';' };

        private static readonly LoadGuard Guard = new LoadGuard(LoadNative);

        [DllImport("kernel32", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "LoadLibraryW", ExactSpelling = true)]
        private static extern IntPtr LoadLibraryW(string fileName);

        private static string Architecture => Environment.Is64BitProcess ? "x64" : "x86";

        internal static void EnsureLoaded()
        {
            Guard.Ensure();
        }

        internal static string ResolvePath(string baseDirectory, bool is64BitProcess)
        {
            return Path.Combine(baseDirectory, is64BitProcess ? "x64" : "x86", DllName);
        }

        internal static IReadOnlyList<string> ProbePaths(string baseDirectory, string? relativeSearchPath, string? codeBase, string? location, bool is64BitProcess)
        {
            List<string> directories = new List<string> { baseDirectory };
            if (relativeSearchPath is not null)
            {
                foreach (string entry in relativeSearchPath.Split(SearchPathSeparator, StringSplitOptions.RemoveEmptyEntries))
                {
                    string trimmed = entry.Trim();
                    if (trimmed.Length > 0)
                    {
                        directories.Add(Path.Combine(baseDirectory, trimmed));
                    }
                }
            }

            if (Uri.TryCreate(codeBase, UriKind.Absolute, out Uri? uri) && uri.IsFile)
            {
                directories.Add(Path.GetDirectoryName(uri.LocalPath));
            }

            if (!string.IsNullOrEmpty(location))
            {
                directories.Add(Path.GetDirectoryName(location));
            }

            List<string> paths = new List<string>(directories.Count);
            foreach (string directory in directories)
            {
                string path = Path.GetFullPath(ResolvePath(directory, is64BitProcess));
                if (!paths.Contains(path, StringComparer.OrdinalIgnoreCase))
                {
                    paths.Add(path);
                }
            }

            return paths;
        }

        internal static void Load(string fullPath)
        {
            if (LoadLibraryW(fullPath) != IntPtr.Zero)
            {
                return;
            }

            int error = Marshal.GetLastWin32Error();
            string? folder = Path.GetFileName(Path.GetDirectoryName(fullPath));
            throw new PlatformNotSupportedException(
                $"Could not load the native ExcelReader library '{fullPath}' (Win32 error {error.ToString(CultureInfo.InvariantCulture)}). "
                + $"This is a {Architecture} process and the library is expected in a '{folder}' folder beside the application. "
                + "Set PlatformTarget to x64 or turn off Prefer 32-bit, and deploy the matching ExcelReader.Native.dll.");
        }

        internal static PlatformNotSupportedException NotFound(IReadOnlyList<string> probed)
        {
            return new PlatformNotSupportedException(
                $"Could not find the native ExcelReader library. This is a {Architecture} process; looked for: {string.Join("; ", probed)}. "
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

        private static void LoadNative()
        {
            Assembly assembly = typeof(NativeLoader).Assembly;
            AppDomain domain = AppDomain.CurrentDomain;
            IReadOnlyList<string> probed = ProbePaths(domain.BaseDirectory, domain.RelativeSearchPath, assembly.EscapedCodeBase, assembly.Location, Environment.Is64BitProcess);
            string? found = probed.FirstOrDefault(File.Exists);
            if (found is null)
            {
                throw NotFound(probed);
            }

            Load(found);
            CheckAbi(NativeMethods.xl_abi_version());
        }
    }
}
