using System.Reflection;

namespace ExcelReader.Tests
{
    public sealed class DefFileSyncTests
    {
        // The repo root, found by walking up from the test binary until the solution file appears.
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

        // There are only two physically distinct .def files in this repo. `cpp/include/xl` is a git
        // symlink to `src/ExcelReader.Native/include`, so the C++ package reaches the canonical
        // file through that symlink rather than through a maintained copy of its own - comparing it
        // here would just compare the canonical file to itself and could never catch real drift.
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
        public void Should_ExportParseArrow_When_ReadingTheCanonicalDef()
        {
            Assert.Contains("xl_parse_arrow", ReadExports(RepoRoot(), DefPaths[0]), StringComparer.Ordinal);
        }
    }
}
