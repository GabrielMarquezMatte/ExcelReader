using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Tests.Writer
{
    public sealed class StagedFileTests
    {
        [Fact]
        public void Should_Commit_When_TheDestinationNameIsAsLongAsTheFileSystemAllows()
        {
            string directory = Path.Combine(Path.GetTempPath(), $"excelreader-staged-{Guid.NewGuid():N}");
            Directory.CreateDirectory(directory);
            try
            {
                string destination = Path.Combine(directory, new string('a', 240) + ".xlsx");
                File.WriteAllText(destination, "ORIGINAL");

                using (StagedFile staged = new(destination))
                {
                    staged.Stream.Write("new"u8);
                    staged.Commit();
                }

                Assert.Equal("new", File.ReadAllText(destination));
                Assert.Equal([destination], Directory.GetFiles(directory));
            }
            finally
            {
                Directory.Delete(directory, recursive: true);
            }
        }

        [Fact]
        public void Should_Not_Throw_When_Disposing_An_Uncommitted_File_That_Cannot_Be_Deleted()
        {
            Assert.SkipUnless(OperatingSystem.IsWindows(), "a read-only file blocks deletion only on Windows");
            string directory = Path.Combine(Path.GetTempPath(), $"excelreader-staged-{Guid.NewGuid():N}");
            Directory.CreateDirectory(directory);
            try
            {
                using (new StagedFile(Path.Combine(directory, "out.xlsx")))
                {
                    File.SetAttributes(Directory.GetFiles(directory)[0], FileAttributes.ReadOnly);
                }

                Assert.Single(Directory.GetFiles(directory));
            }
            finally
            {
                foreach (string file in Directory.GetFiles(directory))
                {
                    File.SetAttributes(file, FileAttributes.Normal);
                }
                Directory.Delete(directory, recursive: true);
            }
        }
    }
}
