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

        private static readonly string[] DefPaths =
        [
            Path.Combine("src", "ExcelReader.Native", "include", "excelreader.def"),
            Path.Combine("rust", "excelreader", "excelreader.def"),
            Path.Combine("cpp", "include", "xl", "excelreader.def"),
        ];

        private static string[] ReadExports(string root, string relative)
        {
            string[] lines = File.ReadAllLines(Path.Combine(root, relative));
            Assert.Equal("EXPORTS", lines[0].Trim());
            return [.. lines.Skip(1).Select(l => l.Trim()).Where(l => l.Length > 0)];
        }

        [Fact]
        public void Should_ListIdenticalExports_When_ComparingTheThreeDefCopies()
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
