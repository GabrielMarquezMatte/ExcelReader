using ExcelReader.NetFramework.Reader;
using ExcelReader.NetFramework.Writer;

namespace ExcelReader.NetFramework.Tests.Writer
{
    public sealed class EncryptedWriteTests
    {
        private const string Secret = "hunter2";

        private static List<WriteRow> Rows()
        {
            return new List<WriteRow>
            {
                new WriteRow { Name = "ação 😀", Id = 7, Qty = 1, When = new DateTime(2026, 3, 15, 10, 20, 30) },
                new WriteRow { Name = "plain", Id = 8, Qty = 2, When = new DateTime(2026, 1, 1) },
            };
        }

        private static void AssertEncrypted(byte[] bytes, ExcelFormat format)
        {
            ExcelReaderException error = Assert.Throws<ExcelReaderException>(() => ExcelWorkbook.Open(bytes, new ExcelOpenOptions { Format = format }));
            Assert.Equal(ExcelStatus.PasswordRequired, error.Status);

            using ExcelWorkbook workbook = ExcelWorkbook.Open(bytes, new ExcelOpenOptions { Format = format, Password = Secret });
            WriteRow[] read = workbook.Sheet("Vendas").Read<WriteRow>();
            Assert.Equal(2, read.Length);
            Assert.Equal("ação 😀", read[0].Name);
            Assert.Equal(8, read[1].Id);
        }

        [Theory]
        [InlineData(ExcelFormat.Xlsx)]
        [InlineData(ExcelFormat.Xlsb)]
        public void EncryptedRoundTripInMemory(ExcelFormat format)
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(format, new ExcelWriteOptions { Password = Secret });
            writer.WriteSheet("Vendas", Rows());
            AssertEncrypted(writer.CloseToArray(), format);
        }

        [Theory]
        [InlineData(ExcelFormat.Xlsx, ".xlsx")]
        [InlineData(ExcelFormat.Xlsb, ".xlsb")]
        public void EncryptedRoundTripToPath(ExcelFormat format, string extension)
        {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);
            try
            {
                using (ExcelWorkbookWriter writer = ExcelWorkbookWriter.Create(path, new ExcelWriteOptions { Password = Secret }))
                {
                    writer.WriteSheet("Vendas", Rows());
                    Assert.False(File.Exists(path));
                    writer.Close();
                }

                AssertEncrypted(File.ReadAllBytes(path), format);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void EncryptedCloseToStream()
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx, new ExcelWriteOptions { Password = Secret });
            writer.WriteSheet("Vendas", Rows());
            using MemoryStream stream = new MemoryStream();
            writer.CloseTo(stream);
            AssertEncrypted(stream.ToArray(), ExcelFormat.Xlsx);
        }

        [Theory]
        [InlineData(ExcelFormat.Csv)]
        [InlineData(ExcelFormat.Xls)]
        public void PasswordRejectedForCsvAndXls(ExcelFormat format)
        {
            Assert.Throws<NotSupportedException>(() => ExcelWorkbookWriter.CreateInMemory(format, new ExcelWriteOptions { Password = "x" }));
        }

        [Fact]
        public void PasswordRejectedForCsvPath()
        {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".csv");
            Assert.Throws<NotSupportedException>(() => ExcelWorkbookWriter.Create(path, new ExcelWriteOptions { Password = "x" }));
            Assert.False(File.Exists(path));
        }

        [Fact]
        public void EncryptedPathAbortLeavesNoFile()
        {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelWorkbookWriter writer = ExcelWorkbookWriter.Create(path, new ExcelWriteOptions { Password = Secret });
            writer.WriteSheet("Vendas", Rows());
            writer.Dispose();
            Assert.False(File.Exists(path));
        }

        [Fact]
        public void EncryptedCloseReplacesExistingFileAndLeavesNoTemp()
        {
            string directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                string path = Path.Combine(directory, "out.xlsx");
                File.WriteAllBytes(path, new byte[] { 1, 2, 3 });
                using (ExcelWorkbookWriter writer = ExcelWorkbookWriter.Create(path, new ExcelWriteOptions { Password = Secret }))
                {
                    writer.WriteSheet("Vendas", Rows());
                    writer.Close();
                }

                AssertEncrypted(File.ReadAllBytes(path), ExcelFormat.Xlsx);
                Assert.Equal(new[] { path }, Directory.GetFiles(directory));
            }
            finally
            {
                Directory.Delete(directory, true);
            }
        }

        [Fact]
        public void EmptyPasswordThrowsAtCreate()
        {
            Assert.Throws<ArgumentException>(() => ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx, new ExcelWriteOptions { Password = "" }));
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            Assert.Throws<ArgumentException>(() => ExcelWorkbookWriter.Create(path, new ExcelWriteOptions { Password = "" }));
        }
    }
}
