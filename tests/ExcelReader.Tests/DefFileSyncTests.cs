using System.Reflection;
using System.Runtime.InteropServices;
using ExcelReader.Native;

namespace ExcelReader.Tests
{
    public sealed class DefFileSyncTests
    {
        private static string RepoRoot()
        {
            DirectoryInfo? dir = new(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!);
            while (dir is not null && !File.Exists(Path.Combine(dir.FullName, "ExcelReader.slnx")))
            {
                dir = dir.Parent;
            }
            Assert.NotNull(dir);
            return dir.FullName;
        }

        private static readonly string[] DefPaths =
        [
            Path.Combine("src", "ExcelReader.Native", "include", "excelreader.def"),
            Path.Combine("rust", "excelreader", "excelreader.def"),
        ];

        private static string[] ReadExports(string root, string relative)
        {
            string[] lines = File.ReadAllLines(Path.Combine(root, relative));
            Assert.Equal("EXPORTS", lines[0].Trim());
            return [.. lines.Skip(1).Select(l => l.Trim()).Where(l => l.Length > 0)];
        }

        private static string[] ReflectedExports()
        {
            return
            [
                .. typeof(NativeApi).Assembly.GetTypes()
                    .SelectMany(t => t.GetMethods(BindingFlags.Public | BindingFlags.NonPublic
                        | BindingFlags.Static | BindingFlags.Instance))
                    .Select(m => m.GetCustomAttribute<UnmanagedCallersOnlyAttribute>())
                    .Where(a => a is not null)
                    .Select(a => a!.EntryPoint)
                    .OfType<string>()
                    .Order(StringComparer.Ordinal),
            ];
        }

        [Fact]
        public void Should_ListIdenticalExports_When_ComparingTheCanonicalAndRustDefCopies()
        {
            string root = RepoRoot();
            string[] canonical = ReadExports(root, DefPaths[0]);

            foreach (string copy in DefPaths.Skip(1))
            {
                Assert.Equal(canonical, ReadExports(root, copy));
            }
        }

        [Fact]
        public void Should_MatchReflectedUnmanagedCallersOnlyEntryPoints_When_ReadingTheCanonicalDef()
        {
            string[] canonical = [.. ReadExports(RepoRoot(), DefPaths[0]).OrderBy(name => name, StringComparer.Ordinal)];
            Assert.Equal(ReflectedExports(), canonical);
        }
    }
}
