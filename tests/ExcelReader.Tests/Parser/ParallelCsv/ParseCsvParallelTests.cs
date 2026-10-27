using System.Globalization;
using System.Text;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;

namespace ExcelReader.Tests.Parser.ParallelCsv
{
    public class ParseCsvParallelTests
    {
        private sealed class Row
        {
            public string? Name { get; set; }
            public int Age { get; set; }
        }

        private static byte[] BuildCsv(int rows)
        {
            var sb = new StringBuilder("Name,Age\n");
            for (int i = 0; i < rows; i++)
            {
                sb.Append(CultureInfo.InvariantCulture, $"name{i:D5},{i}\n");
            }
            return Encoding.UTF8.GetBytes(sb.ToString());
        }

        private static async Task<List<Row>> DrainAsync(IAsyncEnumerable<Row> source)
        {
            var list = new List<Row>();
            await foreach (Row row in source)
            {
                list.Add(row);
            }
            return list;
        }

        [Fact]
        public async Task MemoryOverloadMatchesTheSequentialParser()
        {
            byte[] csv = BuildCsv(20_000);
            using var reader = Excel.FromCsv(csv);
            List<Row> expected = [.. ExcelParser.FromAttributes<Row>().Parse(reader.FirstSheet)];

            using CsvReader source = Excel.FromCsv(csv);
            List<Row> actual = await DrainAsync(SheetParallel.ParseAsync(source.FirstSheet, ExcelParser.FromAttributes<Row>(), options: new SheetParallelOptions { DegreeOfParallelism = 4 }, ct: TestContext.Current.CancellationToken));

            Assert.Equal(expected.Count, actual.Count);
            for (int i = 0; i < expected.Count; i++)
            {
                Assert.Equal(expected[i].Name, actual[i].Name);
                Assert.Equal(expected[i].Age, actual[i].Age);
            }
        }

        [Fact]
        public async Task AFailureAfterBlankLinesReportsTheSameRowAsTheSequentialParser()
        {
            var sb = new StringBuilder("Name,Age\n");
            for (int i = 0; i < 30_000; i++)
            {
                sb.Append(i == 25_000 ? "bad,x\n" : $"name{i},{i}\n");
                if (i % 997 == 0)
                {
                    sb.Append('\n');
                }
            }
            byte[] csv = Encoding.UTF8.GetBytes(sb.ToString());
            var options = new ExcelParserOptions { ThrowOnParseFailure = true };

            using CsvReader sequentialSource = Excel.FromCsv(csv);
            ExcelParseException sequential = Assert.Throws<ExcelParseException>(
                () => ExcelParser.FromAttributes<Row>(options).Parse(sequentialSource.FirstSheet).ToList());
            using CsvReader parallelSource = Excel.FromCsv(csv);
            ExcelParseException parallel = await Assert.ThrowsAsync<ExcelParseException>(
                () => DrainAsync(SheetParallel.ParseAsync(parallelSource.FirstSheet, ExcelParser.FromAttributes<Row>(options), options: new SheetParallelOptions { DegreeOfParallelism = 4 }, ct: TestContext.Current.CancellationToken)));

            Assert.Equal(sequential.Row, parallel.Row);
        }

        private static ExcelParser<Row> IndexParser(ExcelParserOptions? options = null)
        {
            return ExcelParser.Build<Row>(static b => b
                .Factory(static () => new Row())
                .PropertyAt(0, ExcelCellReaders.String, static (ref m, v) => m.Name = v)
                .PropertyAt(1, ExcelCellReaders.Parsable, static (ref Row m, int v) => m.Age = v), options);
        }

        private static byte[] BuildHeaderlessCsv(int rows, int badRow = -1)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < rows; i++)
            {
                sb.Append(i == badRow ? "bad,x\n" : $"name{i:D5},{i}\n");
            }
            return Encoding.UTF8.GetBytes(sb.ToString());
        }

        [Fact]
        public async Task AnIndexBasedMapReadsEveryRowAsDataLikeTheSequentialParser()
        {
            byte[] csv = BuildHeaderlessCsv(40_000);

            using CsvReader source = Excel.FromCsv(csv);
            List<Row> actual = await DrainAsync(SheetParallel.ParseAsync(source.FirstSheet, IndexParser(), options: new SheetParallelOptions { DegreeOfParallelism = 4 }, ct: TestContext.Current.CancellationToken));

            Assert.Equal(40_000, actual.Count);
            for (int i = 0; i < actual.Count; i++)
            {
                Assert.Equal($"name{i:D5}", actual[i].Name);
                Assert.Equal(i, actual[i].Age);
            }
        }

        [Fact]
        public async Task AnIndexBasedMapReportsAFailureAtTheSequentialRow()
        {
            byte[] csv = BuildHeaderlessCsv(40_000, badRow: 30_000);
            var options = new ExcelParserOptions { ThrowOnParseFailure = true };

            using CsvReader sequentialSource = Excel.FromCsv(csv);
            ExcelParseException sequential = Assert.Throws<ExcelParseException>(
                () => IndexParser(options).Parse(sequentialSource.FirstSheet).ToList());
            using CsvReader parallelSource = Excel.FromCsv(csv);
            ExcelParseException parallel = await Assert.ThrowsAsync<ExcelParseException>(
                () => DrainAsync(SheetParallel.ParseAsync(parallelSource.FirstSheet, IndexParser(options), options: new SheetParallelOptions { DegreeOfParallelism = 4 }, ct: TestContext.Current.CancellationToken)));

            Assert.Equal(30_001, sequential.Row);
            Assert.Equal(sequential.Row, parallel.Row);
        }

        [Fact]
        public async Task EnumeratingTheSameResultTwiceYieldsEveryRowBothTimes()
        {
            using CsvReader source = Excel.FromCsv(BuildCsv(20_000));
            IAsyncEnumerable<Row> rows = SheetParallel.ParseAsync(source.FirstSheet, ExcelParser.FromAttributes<Row>(), options: new SheetParallelOptions { DegreeOfParallelism = 4 }, ct: TestContext.Current.CancellationToken);

            Assert.Equal(20_000, (await DrainAsync(rows).WaitAsync(TimeSpan.FromSeconds(30), TestContext.Current.CancellationToken)).Count);
            Assert.Equal(20_000, (await DrainAsync(rows).WaitAsync(TimeSpan.FromSeconds(30), TestContext.Current.CancellationToken)).Count);
        }

        [Fact]
        public async Task FileOverloadMatchesTheSequentialParser()
        {
            byte[] csv = BuildCsv(20_000);
            string path = Path.Combine(Path.GetTempPath(), $"exr-{Guid.NewGuid():N}.csv");
            await File.WriteAllBytesAsync(path, csv, TestContext.Current.CancellationToken);
            try
            {
                using var reader = Excel.FromCsvFile(path);
                List<Row> expected = [.. ExcelParser.FromAttributes<Row>().Parse(reader.FirstSheet)];

                using CsvReader source = Excel.FromCsvFile(path);
                List<Row> actual = await DrainAsync(SheetParallel.ParseAsync(source.FirstSheet, ExcelParser.FromAttributes<Row>(), options: new SheetParallelOptions { DegreeOfParallelism = 8 }, ct: TestContext.Current.CancellationToken));

                Assert.Equal(expected.Count, actual.Count);
                Assert.Equal(expected[^1].Name, actual[^1].Name);
                Assert.Equal(expected[0].Age, actual[0].Age);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public async Task ATinySourceStillProducesTheRightRowsThroughTheSequentialFallback()
        {
            byte[] csv = BuildCsv(3);

            using CsvReader source = Excel.FromCsv(csv);
            List<Row> actual = await DrainAsync(SheetParallel.ParseAsync(source.FirstSheet, ExcelParser.FromAttributes<Row>(), options: new SheetParallelOptions { DegreeOfParallelism = 8 }, ct: TestContext.Current.CancellationToken));

            Assert.Equal(3, actual.Count);
            Assert.Equal("name00000", actual[0].Name);
            Assert.Equal(2, actual[2].Age);
        }

        [Fact]
        public async Task ANonUtf8EncodingFallsBackAndStillDecodesCorrectly()
        {
            Encoding latin1 = Encoding.Latin1;
            byte[] csv = latin1.GetBytes("Name,Age\nJosé,30\nRenée,41\n");
            var options = new CsvReaderOptions { Encoding = latin1 };

            using CsvReader source = Excel.FromCsv(csv, new ExcelReaderOptions { Csv = options });

            List<Row> actual = await DrainAsync(
                SheetParallel.ParseAsync(source.FirstSheet, ExcelParser.FromAttributes<Row>(), options: new SheetParallelOptions { DegreeOfParallelism = 8 }, ct: TestContext.Current.CancellationToken));

            Assert.Equal(2, actual.Count);
            Assert.Equal("José", actual[0].Name);
            Assert.Equal("Renée", actual[1].Name);
        }

        [Fact]
        public async Task ADegreeOfOneIsTheSequentialPath()
        {
            byte[] csv = BuildCsv(5_000);

            using CsvReader source = Excel.FromCsv(csv);
            List<Row> actual = await DrainAsync(SheetParallel.ParseAsync(source.FirstSheet, ExcelParser.FromAttributes<Row>(), options: new SheetParallelOptions { DegreeOfParallelism = 1 }, ct: TestContext.Current.CancellationToken));

            Assert.Equal(5_000, actual.Count);
            Assert.Equal(4_999, actual[^1].Age);
        }

        [Fact]
        public void ANegativeDegreeOfParallelismThrows()
        {
            byte[] csv = BuildCsv(10);

            using CsvReader source = Excel.FromCsv(csv);

            Assert.Throws<ArgumentOutOfRangeException>(
                () => SheetParallel.ParseAsync(source.FirstSheet, ExcelParser.FromAttributes<Row>(), options: new SheetParallelOptions { DegreeOfParallelism = -1 }, ct: TestContext.Current.CancellationToken));
        }

        [Fact]
        public void AHeaderRowOnTheParallelOptionsThrowsBecauseTypedParsingReadsItFromTheParserConfig()
        {
            byte[] csv = BuildCsv(10);

            using CsvReader source = Excel.FromCsv(csv);

            ArgumentException exception = Assert.Throws<ArgumentException>(
                () => SheetParallel.ParseAsync(source.FirstSheet, ExcelParser.FromAttributes<Row>(), options: new SheetParallelOptions { HeaderRow = 2 }, ct: TestContext.Current.CancellationToken));
            Assert.Contains(nameof(ExcelParserOptions.HeaderRow), exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public async Task AMemoryStreamWithAnExposedBufferIsUnwrappedAndParallelized()
        {
            byte[] csv = BuildCsv(20_000);
            using var ms = new MemoryStream(csv, 0, csv.Length, writable: false, publiclyVisible: true);

            using CsvReader source = Excel.FromCsv(ms, leaveOpen: true);
            List<Row> actual = await DrainAsync(SheetParallel.ParseAsync(source.FirstSheet, ExcelParser.FromAttributes<Row>(), options: new SheetParallelOptions { DegreeOfParallelism = 4 }, ct: TestContext.Current.CancellationToken));

            Assert.Equal(20_000, actual.Count);
            Assert.Equal("name00000", actual[0].Name);
            Assert.Equal(19_999, actual[^1].Age);
        }

        [Fact]
        public async Task AMemoryStreamWithoutAnExposedBufferStillProducesTheRightRows()
        {
            byte[] csv = BuildCsv(20_000);
            using var ms = new MemoryStream(csv, 0, csv.Length, writable: false, publiclyVisible: false);

            using CsvReader source = Excel.FromCsv(ms, leaveOpen: true);
            List<Row> actual = await DrainAsync(SheetParallel.ParseAsync(source.FirstSheet, ExcelParser.FromAttributes<Row>(), options: new SheetParallelOptions { DegreeOfParallelism = 4 }, ct: TestContext.Current.CancellationToken));

            Assert.Equal(20_000, actual.Count);
            Assert.Equal(19_999, actual[^1].Age);
        }

        [Fact]
        public async Task AFileStreamIsUnwrappedToItsHandleAndParallelized()
        {
            byte[] csv = BuildCsv(20_000);
            string path = Path.Combine(Path.GetTempPath(), $"exr-{Guid.NewGuid():N}.csv");
            await File.WriteAllBytesAsync(path, csv, TestContext.Current.CancellationToken);
            try
            {
                await using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);

                using CsvReader source = Excel.FromCsv(fs, leaveOpen: true);
                List<Row> actual = await DrainAsync(SheetParallel.ParseAsync(source.FirstSheet, ExcelParser.FromAttributes<Row>(), options: new SheetParallelOptions { DegreeOfParallelism = 8 }, ct: TestContext.Current.CancellationToken));

                Assert.Equal(20_000, actual.Count);
                Assert.Equal(19_999, actual[^1].Age);
                Assert.False(fs.SafeFileHandle.IsClosed);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public async Task AnUnwrappableStreamFallsBackWithoutThrowing()
        {
            byte[] csv = BuildCsv(20_000);
            string path = Path.Combine(Path.GetTempPath(), $"exr-{Guid.NewGuid():N}.csv");
            await File.WriteAllBytesAsync(path, csv, TestContext.Current.CancellationToken);
            try
            {
                await using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
                await using var buffered = new BufferedStream(fs);

                using CsvReader source = Excel.FromCsv(buffered, leaveOpen: true);
                List<Row> actual = await DrainAsync(SheetParallel.ParseAsync(source.FirstSheet, ExcelParser.FromAttributes<Row>(), options: new SheetParallelOptions { DegreeOfParallelism = 8 }, ct: TestContext.Current.CancellationToken));

                Assert.Equal(20_000, actual.Count);
                Assert.Equal(19_999, actual[^1].Age);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public async Task ANonSeekableStreamFallsBackWithoutThrowing()
        {
            byte[] csv = BuildCsv(20_000);
            await using var pipe = new NonSeekableStream(csv);

            using CsvReader source = Excel.FromCsv(pipe, leaveOpen: true);
            List<Row> actual = await DrainAsync(SheetParallel.ParseAsync(source.FirstSheet, ExcelParser.FromAttributes<Row>(), options: new SheetParallelOptions { DegreeOfParallelism = 8 }, ct: TestContext.Current.CancellationToken));

            Assert.Equal(20_000, actual.Count);
            Assert.Equal(19_999, actual[^1].Age);
        }
    }
}
