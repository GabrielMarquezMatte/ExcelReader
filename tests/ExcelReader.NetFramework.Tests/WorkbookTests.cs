using ExcelReader.NetFramework.Interop;

namespace ExcelReader.NetFramework.Tests
{
    public sealed class WorkbookTests
    {
        private static readonly string Encrypted = Fixtures.Path("encrypted", "agile-aes256-sha512.xlsx");

        [Theory]
        [InlineData("RealExcel.xlsx")]
        [InlineData("RealExcel.xlsb")]
        public void OpensAndListsSheets(string file)
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path(file));
            Assert.NotEmpty(workbook.Sheets);
            Assert.False(string.IsNullOrEmpty(workbook.Sheets[0].Name));
            Assert.Equal(0, workbook.Sheet(workbook.Sheets[0].Name).Id);
            Assert.False(workbook.IsDate1904);
        }

        [Fact]
        public void OpenBytesMatchesOpenPath()
        {
            string path = Fixtures.Path("RealExcel.xlsx");
            using ExcelWorkbook fromPath = ExcelWorkbook.Open(path);
            using ExcelWorkbook fromBytes = ExcelWorkbook.Open(File.ReadAllBytes(path));
            Assert.Equal(fromPath.Sheets.Select(s => s.Name).ToArray(), fromBytes.Sheets.Select(s => s.Name).ToArray(), StringComparer.Ordinal);
        }

        [Theory]
        [InlineData("tables.xlsx")]
        [InlineData("tables.xlsb")]
        public void TablesReportTheirMetadata(string file)
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("tables", file));
            Assert.Equal(4, workbook.Tables.Count);
            AssertTable(workbook, "Sales", 0, 3, 2, 7, 4, 1, 1);
            AssertTable(workbook, "Stock", 0, 3, 6, 6, 7, 1, 0);
            AssertTable(workbook, "NoHeader", 1, 1, 0, 3, 1, 0, 0);
            AssertTable(workbook, "NoStyle", 1, 0, 3, 2, 4, 1, 0);
        }

        [Fact]
        public void TableLookupIgnoresCase()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("tables", "tables.xlsx"));
            Assert.Equal(Xl.TableSheet(0), workbook.Table("sales").Id);
        }

        [Fact]
        public void UnknownSheetOrTableThrows()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("tables", "tables.xlsx"));
            Assert.Contains("nope", Assert.Throws<KeyNotFoundException>(() => workbook.Sheet("nope")).Message, StringComparison.Ordinal);
            Assert.Contains("nope", Assert.Throws<KeyNotFoundException>(() => workbook.Table("nope")).Message, StringComparison.Ordinal);
            Assert.Throws<ArgumentOutOfRangeException>(() => workbook.Sheet(99));
        }

        [Fact]
        public void PasswordRequired()
        {
            ExcelReaderException error = Assert.Throws<ExcelReaderException>(() => ExcelWorkbook.Open(Encrypted));
            Assert.Equal(ExcelStatus.PasswordRequired, error.Status);
        }

        [Fact]
        public void PasswordIncorrect()
        {
            ExcelOpenOptions options = new ExcelOpenOptions { Password = "wrong" };
            ExcelReaderException error = Assert.Throws<ExcelReaderException>(() => ExcelWorkbook.Open(Encrypted, options));
            Assert.Equal(ExcelStatus.PasswordIncorrect, error.Status);
        }

        [Fact]
        public void PasswordOpensEncryptedWorkbook()
        {
            using ExcelWorkbook plain = ExcelWorkbook.Open(Fixtures.Path("encrypted", "agile-aes256-sha512.plain.xlsx"));
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Encrypted, new ExcelOpenOptions { Password = "hunter2" });
            Assert.Equal(plain.Sheets[0].Name, workbook.Sheets[0].Name);
        }

        [Fact]
        public void MissingFileThrowsWithNativeMessage()
        {
            ExcelReaderException error = Assert.Throws<ExcelReaderException>(() => ExcelWorkbook.Open(Fixtures.Path("missing.xlsx")));
            Assert.NotEmpty(error.Message);
            Assert.NotEqual(ExcelStatus.Ok, error.Status);
        }

        [Fact]
        public void CsvMustBeExplicit()
        {
            string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".csv");
            File.WriteAllText(path, "a;b\r\n1;2\r\n");
            try
            {
                ExcelOpenOptions options = new ExcelOpenOptions { Format = ExcelFormat.Csv, CsvDelimiter = ';' };
                using ExcelWorkbook workbook = ExcelWorkbook.Open(path, options);
                Assert.Single(workbook.Sheets);
                ExcelReaderException error = Assert.Throws<ExcelReaderException>(() => ExcelWorkbook.Open(path, new ExcelOpenOptions { CsvDelimiter = ';' }));
                Assert.Equal(ExcelStatus.Error, error.Status);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void NonAsciiCsvDelimiterIsRejected()
        {
            ExcelOpenOptions options = new ExcelOpenOptions { Format = ExcelFormat.Csv, CsvDelimiter = '§' };
            Assert.Throws<ArgumentOutOfRangeException>(() => ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"), options));
        }

        [Fact]
        public void DisposeTwiceIsSafe()
        {
            ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            workbook.Dispose();
            workbook.Dispose();
            Assert.Throws<ObjectDisposedException>(() => workbook.Sheets);
        }

        [Fact]
        public void ExceptionCrossesAppDomainsWithItsStatus()
        {
            AppDomain domain = AppDomain.CreateDomain("serialization", null, AppDomain.CurrentDomain.SetupInformation);
            try
            {
                ExcelReaderException error = Assert.Throws<ExcelReaderException>(() => domain.DoCallBack(ThrowPasswordRequired));
                Assert.Equal(ExcelStatus.PasswordRequired, error.Status);
                Assert.Equal("locked", error.Message);
            }
            finally
            {
                AppDomain.Unload(domain);
            }
        }

        [Fact]
        public void UseAfterDisposeThrowsObjectDisposedException()
        {
            ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            _ = workbook.Sheets;
            workbook.Dispose();
            Assert.Throws<ObjectDisposedException>(() => workbook.Sheets);
            Assert.Throws<ObjectDisposedException>(() => workbook.Sheet(0));
        }

        private static void ThrowPasswordRequired()
        {
            throw new ExcelReaderException(ExcelStatus.PasswordRequired, "locked");
        }

        private static void AssertTable(ExcelWorkbook workbook, string name, int sheet, int firstRow, int firstColumn, int lastRow, int lastColumn, int headers, int totals)
        {
            ExcelTableInfo table = workbook.Tables.Single(t => string.Equals(t.Name, name, StringComparison.Ordinal));
            Assert.Equal(new[] { sheet, firstRow, firstColumn, lastRow, lastColumn, headers, totals },
                new[] { table.SheetIndex, table.FirstRow, table.FirstColumn, table.LastRow, table.LastColumn, table.HeaderRowCount, table.TotalsRowCount });
        }
    }
}
