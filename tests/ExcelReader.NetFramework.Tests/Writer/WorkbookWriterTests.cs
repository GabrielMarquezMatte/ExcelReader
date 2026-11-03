using System.Text;
using ExcelReader.NetFramework.Reader;
using ExcelReader.NetFramework.Writer;

namespace ExcelReader.NetFramework.Tests.Writer
{
    public sealed class WorkbookWriterTests
    {
        private const string Unicode = "ação 😀";
        private static readonly string[] SheetNames = { "Resumo", "Vendas" };

        private static string TempPath(string extension)
        {
            return Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);
        }

        private static void WriteSingleRow(ExcelWorkbookWriter writer)
        {
            writer.StartSheet("Data");
            writer.StartRow();
            writer.Write(Unicode);
            writer.Write(42L);
            writer.Write(1.5);
            writer.Write(true);
            writer.WriteDate(new DateTime(2026, 1, 1));
            writer.WriteTime(new TimeSpan(13, 30, 0));
            writer.WriteDateTime(new DateTime(2026, 1, 1, 13, 30, 0));
            writer.WriteEmpty();
            writer.EndRow();
            writer.EndSheet();
        }

        private static void WriteTwoRows(ExcelWorkbookWriter writer)
        {
            writer.StartSheet("Data");
            for (int i = 0; i < 2; i++)
            {
                writer.StartRow();
                writer.Write("row" + i);
                writer.Write((long)i);
                writer.EndRow();
            }

            writer.EndSheet();
        }

        private static List<List<string>> ReadAll(byte[] bytes, ExcelFormat format)
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(bytes, new ExcelOpenOptions { Format = format });
            return Dump(workbook);
        }

        private static List<List<string>> Dump(ExcelWorkbook workbook)
        {
            return workbook.Sheet(0).ReadRows().Select(r => r.Cells.Select(c => c.Type + ":" + c.Text).ToList()).ToList();
        }

        [Theory]
        [InlineData(ExcelFormat.Xlsx, "String:ação 😀|Number:42|Number:1.5|Bool:1|Date:46023|Number:0.5625|Date:46023.5625")]
        [InlineData(ExcelFormat.Xlsb, "String:ação 😀|Number:42|Number:1.5|Bool:1|Date:46023|Number:0.5625|Date:46023.5625")]
        [InlineData(ExcelFormat.Xls, "String:ação 😀|Number:42|Number:1.5|Bool:1|Date:46023|Number:0.5625|Date:46023.5625")]
        [InlineData(ExcelFormat.Csv, "String:ação 😀|String:42|String:1.5|String:true|String:2026-01-01|String:0.5625|String:2026-01-01T13:30:00.0000000Z|Empty:")]
        public void CellsRoundTrip(ExcelFormat format, string expected)
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(format);
            WriteSingleRow(writer);
            byte[] bytes = writer.CloseToArray();

            List<List<string>> rows = ReadAll(bytes, format);

            Assert.Single(rows);
            Assert.Equal(expected.Split('|'), rows[0].ToArray());
        }

        [Theory]
        [InlineData(ExcelFormat.Xlsx, "0:String:|2:String:x")]
        [InlineData(ExcelFormat.Xlsb, "0:String:|2:String:x")]
        [InlineData(ExcelFormat.Xls, "0:String:|2:String:x")]
        [InlineData(ExcelFormat.Csv, "0:Empty:|1:Empty:|2:String:x")]
        public void EmptyStringAndNullAreDistinctCells(ExcelFormat format, string expected)
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(format);
            writer.StartSheet("Data");
            writer.StartRow();
            writer.Write(string.Empty);
            writer.Write((string?)null);
            writer.Write("x");
            writer.EndRow();
            writer.EndSheet();

            using ExcelWorkbook workbook = ExcelWorkbook.Open(writer.CloseToArray(), new ExcelOpenOptions { Format = format });
            ExcelRow row = workbook.Sheet(0).ReadRows().Single();

            Assert.Equal(expected, string.Join("|", row.Cells.Select(c => c.Column + ":" + c.Type + ":" + c.Text)));
        }

        [Fact]
        public void EmptySheetNameReachesNativeAsAnEmptyName()
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx);
            ExcelReaderException exception = Assert.Throws<ExcelReaderException>(() => writer.StartSheet(string.Empty));
            Assert.Equal(ExcelStatus.Error, exception.Status);
            Assert.DoesNotContain("name_len", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void MultipleSheets()
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx);
            writer.StartSheet("Resumo");
            writer.EndSheet();
            writer.StartSheet("Vendas");
            writer.EndSheet();
            byte[] bytes = writer.CloseToArray();

            using ExcelWorkbook workbook = ExcelWorkbook.Open(bytes);
            Assert.Equal(SheetNames, workbook.Sheets.Select(s => s.Name).ToArray());
        }

        [Fact]
        public void PathAndMemoryProduceSameContent()
        {
            string path = TempPath(".xlsx");
            try
            {
                using (ExcelWorkbookWriter writer = ExcelWorkbookWriter.Create(path))
                {
                    WriteTwoRows(writer);
                    writer.Close();
                }

                byte[] memory;
                using (ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx))
                {
                    WriteTwoRows(writer);
                    memory = writer.CloseToArray();
                }

                List<List<string>> fromMemory = ReadAll(memory, ExcelFormat.Xlsx);
                using ExcelWorkbook workbook = ExcelWorkbook.Open(path);
                Assert.Equal(2, fromMemory.Count);
                Assert.Equal(fromMemory, Dump(workbook));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void CloseToStreamMatchesCloseToArrayAndLeavesStreamOpen()
        {
            byte[] array;
            using (ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Csv))
            {
                WriteTwoRows(writer);
                array = writer.CloseToArray();
            }

            using MemoryStream stream = new MemoryStream();
            using (ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Csv))
            {
                WriteTwoRows(writer);
                writer.CloseTo(stream);
            }

            Assert.True(stream.CanWrite);
            Assert.Equal(array, stream.ToArray());
            Assert.Equal("row0,0\r\nrow1,1\r\n", Encoding.UTF8.GetString(array));
        }

        [Fact]
        public void FormatFromExtensionIsCaseInsensitive()
        {
            string path = TempPath(".XLSB");
            try
            {
                using (ExcelWorkbookWriter writer = ExcelWorkbookWriter.Create(path))
                {
                    WriteTwoRows(writer);
                    writer.Close();
                }

                using ExcelWorkbook workbook = ExcelWorkbook.Open(path, new ExcelOpenOptions { Format = ExcelFormat.Xlsb });
                Assert.Equal(2, workbook.Sheet("Data").ReadRows().Count());
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void UnknownExtensionWithoutFormatThrowsAndCreatesNoFile()
        {
            string path = TempPath(".txt");
            Assert.Throws<ArgumentException>(() => ExcelWorkbookWriter.Create(path));
            Assert.False(File.Exists(path));
        }

        [Fact]
        public void ExplicitFormatOverridesExtension()
        {
            string path = TempPath(".txt");
            try
            {
                using (ExcelWorkbookWriter writer = ExcelWorkbookWriter.Create(path, new ExcelWriteOptions { Format = ExcelFormat.Csv }))
                {
                    WriteTwoRows(writer);
                    writer.Close();
                }

                Assert.Equal("row0,0\r\nrow1,1\r\n", File.ReadAllText(path));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void AutoFormatIsRejected()
        {
            Assert.Throws<ArgumentException>(() => ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Auto));
            Assert.Throws<ArgumentException>(() => ExcelWorkbookWriter.Create(TempPath(".xlsx"), new ExcelWriteOptions { Format = ExcelFormat.Auto }));
        }

        [Fact]
        public void OutOfOrderCallsThrowNativeError()
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx);
            ExcelReaderException exception = Assert.Throws<ExcelReaderException>(() => writer.StartRow());
            Assert.False(string.IsNullOrEmpty(exception.Message));
        }

        [Fact]
        public void InvalidSheetNameThrowsAndWriterCanStillAbort()
        {
            string[] invalid = { new string('x', 32), "a/b" };
            foreach (string name in invalid)
            {
                ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx);
                ExcelReaderException exception = Assert.Throws<ExcelReaderException>(() => writer.StartSheet(name));
                Assert.False(string.IsNullOrEmpty(exception.Message));
                writer.Abort();
            }

            ExcelWorkbookWriter duplicate = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx);
            duplicate.StartSheet("Same");
            duplicate.EndSheet();
            Assert.Throws<ExcelReaderException>(() => duplicate.StartSheet("Same"));
            duplicate.Abort();
        }

        [Fact]
        public void DisposeWithoutCloseLeavesNoFileAtNewPath()
        {
            string path = TempPath(".xlsx");
            try
            {
                using (ExcelWorkbookWriter writer = ExcelWorkbookWriter.Create(path))
                {
                    WriteTwoRows(writer);
                }

                Assert.False(File.Exists(path));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void DiscardKeepsExistingFileBytes(bool explicitAbort)
        {
            string path = TempPath(".xlsx");
            byte[] original = { 1, 2, 3, 4, 5 };
            File.WriteAllBytes(path, original);
            try
            {
                ExcelWorkbookWriter writer = ExcelWorkbookWriter.Create(path);
                WriteTwoRows(writer);
                if (explicitAbort)
                {
                    writer.Abort();
                }

                writer.Dispose();

                Assert.Equal(original, File.ReadAllBytes(path));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ExceptionMidGenerationLeavesExistingFileIntact()
        {
            string path = TempPath(".xlsx");
            byte[] original = { 9, 8, 7 };
            File.WriteAllBytes(path, original);
            try
            {
                Assert.Throws<ExcelReaderException>(() =>
                {
                    using ExcelWorkbookWriter writer = ExcelWorkbookWriter.Create(path);
                    writer.StartSheet("Data");
                    writer.StartRow();
                    writer.Write(1L);
                    writer.StartSheet("Nested");
                });

                Assert.Equal(original, File.ReadAllBytes(path));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void UseAfterCloseThrowsObjectDisposed()
        {
            string path = TempPath(".xlsx");
            try
            {
                ExcelWorkbookWriter writer = ExcelWorkbookWriter.Create(path);
                WriteTwoRows(writer);
                writer.Close();

                Assert.Throws<ObjectDisposedException>(() => writer.Close());
                Assert.Throws<ObjectDisposedException>(() => writer.StartSheet("Again"));
                Assert.Throws<ObjectDisposedException>(() => writer.Write(1L));
                Assert.Throws<ObjectDisposedException>(() => writer.Abort());
                writer.Dispose();
                Assert.True(File.Exists(path));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void UseAfterAbortThrowsObjectDisposed()
        {
            ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Csv);
            writer.Abort();
            Assert.Throws<ObjectDisposedException>(() => writer.StartSheet("Data"));
            Assert.Throws<ObjectDisposedException>(() => writer.CloseToArray());
            writer.Dispose();
            writer.Dispose();
        }

        [Fact]
        public void WrongFinisherThrowsInvalidOperation()
        {
            string path = TempPath(".xlsx");
            using (ExcelWorkbookWriter fileWriter = ExcelWorkbookWriter.Create(path))
            {
                Assert.Throws<InvalidOperationException>(() => fileWriter.CloseToArray());
                using MemoryStream stream = new MemoryStream();
                Assert.Throws<InvalidOperationException>(() => fileWriter.CloseTo(stream));
            }

            using ExcelWorkbookWriter memoryWriter = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx);
            Assert.Throws<InvalidOperationException>(() => memoryWriter.Close());
            Assert.False(File.Exists(path));
        }

        [Fact]
        public void OptionsAreValidated()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Csv, new ExcelWriteOptions { CsvDelimiter = 'ç' }));
            Assert.Throws<ArgumentOutOfRangeException>(() => ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Csv, new ExcelWriteOptions { CsvQuote = 'ç' }));
            Assert.Throws<ArgumentOutOfRangeException>(() => ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx, new ExcelWriteOptions { DegreeOfParallelism = -1 }));
            Assert.Throws<ArgumentOutOfRangeException>(() => ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx, new ExcelWriteOptions { BatchRows = -1 }));
        }

        [Fact]
        public void CsvDelimiterOptionIsHonoured()
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Csv, new ExcelWriteOptions { CsvDelimiter = ';' });
            WriteTwoRows(writer);
            Assert.Equal("row0;0\r\nrow1;1\r\n", Encoding.UTF8.GetString(writer.CloseToArray()));
        }

        [Fact]
        public void TimeOutsideOneDayIsRejected()
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx);
            Assert.Throws<ArgumentOutOfRangeException>(() => writer.WriteTime(TimeSpan.FromTicks(-1)));
            Assert.Throws<ArgumentOutOfRangeException>(() => writer.WriteTime(TimeSpan.FromDays(1)));
        }
    }
}
