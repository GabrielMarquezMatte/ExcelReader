using System.Globalization;
using System.Text;
using ExcelReader.Cli;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Tests
{
    public class CsvDialectAndFormatTests
    {
        private sealed class Person
        {
            public string? Name { get; set; }
            public int Age { get; set; }
        }

        private static string WriteTemp(byte[] bytes)
        {
            string path = Path.Combine(Path.GetTempPath(), $"exr-{Guid.NewGuid():N}.csv");
            File.WriteAllBytes(path, bytes);
            return path;
        }

        private static string WriteTemp(string content, Encoding? encoding = null)
        {
            return WriteTemp((encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false)).GetBytes(content));
        }

        private static List<string[]> ReadAll(IExcelRowReader reader)
        {
            var rows = new List<string[]>();
            using IExcelRowEnumerator enumerator = reader.GetEnumerator();
            while (enumerator.MoveNext())
            {
                var cells = new List<string>();
                foreach (RowCell cell in enumerator.Current.Cells)
                {
                    cells.Add(cell.Value.GetString());
                }
                rows.Add([.. cells]);
            }
            return rows;
        }

        [Fact]
        public void Should_ReadDelimitedText_When_FormatIsCsv()
        {
            string path = WriteTemp("Name,Age\nAda,36\n");
            try
            {
                using IExcelRowReader reader = Excel.Open(path, ExcelFileFormat.Csv);
                List<string[]> rows = ReadAll(reader);

                Assert.Equal(2, rows.Count);
                Assert.Equal(["Ada", "36"], rows[1]);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void Should_HonorCsvOptions_When_FormatIsCsv()
        {
            string path = WriteTemp("Name;City;Age\nAda;London;36\n");
            try
            {
                var options = new ExcelReaderOptions { Csv = CsvReaderOptions.Default with { Delimiter = (byte)';' } };
                using IExcelRowReader reader = Excel.Open(path, ExcelFileFormat.Csv, options);

                Assert.Equal(["Ada", "London", "36"], ReadAll(reader)[1]);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void Should_ReadDelimitedText_When_OpeningAStreamOrBufferAsCsv()
        {
            byte[] csv = Encoding.UTF8.GetBytes("Name,Age\nAda,36\n");

            using (var stream = new MemoryStream(csv, writable: false))
            {
                using IExcelRowReader fromStream = Excel.Open(stream, ExcelFileFormat.Csv);
                Assert.Equal(["Ada", "36"], ReadAll(fromStream)[1]);
            }

            using IExcelRowReader fromMemory = Excel.Open(csv.AsMemory(), ExcelFileFormat.Csv);
            Assert.Equal(["Ada", "36"], ReadAll(fromMemory)[1]);
        }

        [Fact]
        public async Task Should_ReadDelimitedText_When_OpeningAsyncAsCsv()
        {
            string path = WriteTemp("Name,Age\nAda,36\n");
            try
            {
                IExcelRowReader reader = await Excel.OpenAsync(path, ExcelFileFormat.Csv, ct: TestContext.Current.CancellationToken);
                using (reader)
                {
                    Assert.Equal(["Ada", "36"], ReadAll(reader)[1]);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void Should_NotGuessCsv_When_FormatIsAutoDetected()
        {
            byte[] csv = Encoding.UTF8.GetBytes("Name,Age\nAda,36\n");

            Assert.Throws<InvalidDataException>(() => Excel.Open(csv.AsMemory()));
            Assert.Throws<InvalidDataException>(() => Excel.Open(csv.AsMemory(), ExcelFileFormat.Unknown));
        }

        [Fact]
        public void Should_Reject_When_FormatIsEncryptedOoxml()
        {
            byte[] csv = Encoding.UTF8.GetBytes("Name,Age\n");

            Assert.Throws<ArgumentOutOfRangeException>(() => Excel.Open(csv.AsMemory(), ExcelFileFormat.EncryptedOoxml));
            Assert.Throws<ArgumentOutOfRangeException>(() => Excel.Open(csv.AsMemory(), (ExcelFileFormat)99));
        }

        [Fact]
        public void Should_InferDelimiter_When_SniffDialectIsSet()
        {
            string path = WriteTemp("Name;City;Age\nAda;London;36\n");
            try
            {
                using CsvReader reader = Excel.FromCsvFile(path, CsvReaderOptions.Default with { SniffDialect = true });

                Assert.Equal(["Ada", "London", "36"], ReadAll(reader)[1]);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void Should_KeepExplicitEncoding_When_SniffingFindsNoByteOrderMark()
        {
            string path = WriteTemp("Name;City\nJosé;São Paulo\n", Encoding.Latin1);
            try
            {
                var options = CsvReaderOptions.Default with { SniffDialect = true, Encoding = Encoding.Latin1 };
                using CsvReader reader = Excel.FromCsvFile(path, options);

                Assert.Equal(["José", "São Paulo"], ReadAll(reader)[1]);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public async Task Should_InferDelimiter_When_SniffingOnTheParallelPath()
        {
            var sb = new StringBuilder("Name;Age\n");
            for (int i = 0; i < 20_000; i++)
            {
                sb.Append(CultureInfo.InvariantCulture, $"name{i:D5};{i}\n");
            }
            string path = WriteTemp(sb.ToString());
            try
            {
                var options = new CsvParallelOptions
                {
                    DegreeOfParallelism = 4,
                    Reader = CsvReaderOptions.Default with { SniffDialect = true },
                };

                var rows = new List<Person>();
                await foreach (Person person in Excel.ParseCsvParallelAsync(path, ExcelParser.FromAttributes<Person>(), options, ct: TestContext.Current.CancellationToken))
                {
                    rows.Add(person);
                }

                Assert.Equal(20_000, rows.Count);
                Assert.Equal("name00000", rows[0].Name);
                Assert.Equal(19_999, rows[^1].Age);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void Should_SniffDelimiter_When_TheCliReadsCsvWithoutAnInputDelimiter()
        {
            string path = WriteTemp("Name;City;Age\nAda;London;36\n");
            try
            {
                using IExcelRowReader reader = CliCommands.Open(path, sheet: null);

                Assert.Equal(["Ada", "London", "36"], ReadAll(reader)[1]);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void Should_UseTheGivenDelimiter_When_TheCliIsPassedAnInputDelimiter()
        {
            string path = WriteTemp("Name;City;Age\nAda;London;36\n");
            try
            {
                using IExcelRowReader reader = CliCommands.Open(path, sheet: null, password: null, inputDelimiter: ',');

                Assert.Equal(["Ada;London;36"], ReadAll(reader)[1]);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void Should_Reject_When_TheCliIsPassedANonAsciiDelimiter()
        {
            string path = WriteTemp("Name;Age\n");
            try
            {
                Assert.Throws<ArgumentException>(() => CliCommands.Open(path, sheet: null, password: null, inputDelimiter: 'ç'));
            }
            finally
            {
                File.Delete(path);
            }
        }
    }
}
