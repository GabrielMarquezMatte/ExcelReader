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
    }
}
