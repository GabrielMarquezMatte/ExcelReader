using System.Globalization;
using System.Text;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
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
            List<Row> expected = [.. new ExcelParser<Row>().Parse(reader)];

            List<Row> actual = await DrainAsync(Excel.ParseCsvParallelAsync<Row>(csv.AsMemory(), degreeOfParallelism: 4, ct: TestContext.Current.CancellationToken));

            Assert.Equal(expected.Count, actual.Count);
            for (int i = 0; i < expected.Count; i++)
            {
                Assert.Equal(expected[i].Name, actual[i].Name);
                Assert.Equal(expected[i].Age, actual[i].Age);
            }
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
                List<Row> expected = [.. new ExcelParser<Row>().Parse(reader)];

                List<Row> actual = await DrainAsync(Excel.ParseCsvParallelAsync<Row>(path, degreeOfParallelism: 8, ct: TestContext.Current.CancellationToken));

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

            List<Row> actual = await DrainAsync(Excel.ParseCsvParallelAsync<Row>(csv.AsMemory(), degreeOfParallelism: 8, ct: TestContext.Current.CancellationToken));

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

            List<Row> actual = await DrainAsync(
                Excel.ParseCsvParallelAsync<Row>(csv.AsMemory(), degreeOfParallelism: 8, readerOptions: options, ct: TestContext.Current.CancellationToken));

            Assert.Equal(2, actual.Count);
            Assert.Equal("José", actual[0].Name);
            Assert.Equal("Renée", actual[1].Name);
        }

        [Fact]
        public async Task ADegreeOfOneIsTheSequentialPath()
        {
            byte[] csv = BuildCsv(5_000);

            List<Row> actual = await DrainAsync(Excel.ParseCsvParallelAsync<Row>(csv.AsMemory(), degreeOfParallelism: 1, ct: TestContext.Current.CancellationToken));

            Assert.Equal(5_000, actual.Count);
            Assert.Equal(4_999, actual[^1].Age);
        }

        [Fact]
        public void ANegativeDegreeOfParallelismThrows()
        {
            byte[] csv = BuildCsv(10);

            Assert.Throws<ArgumentOutOfRangeException>(
                () => Excel.ParseCsvParallelAsync<Row>(csv.AsMemory(), degreeOfParallelism: -1, ct: TestContext.Current.CancellationToken));
        }

        [Fact]
        public async Task AMemoryStreamWithAnExposedBufferIsUnwrappedAndParallelized()
        {
            byte[] csv = BuildCsv(20_000);
            // The three-arg ctor exposes the buffer, so TryGetBuffer succeeds.
            using var ms = new MemoryStream(csv, 0, csv.Length, writable: false, publiclyVisible: true);

            List<Row> actual = await DrainAsync(Excel.ParseCsvParallelAsync<Row>(ms, degreeOfParallelism: 4, ct: TestContext.Current.CancellationToken));

            Assert.Equal(20_000, actual.Count);
            Assert.Equal("name00000", actual[0].Name);
            Assert.Equal(19_999, actual[^1].Age);
        }

        [Fact]
        public async Task AMemoryStreamWithoutAnExposedBufferStillProducesTheRightRows()
        {
            byte[] csv = BuildCsv(20_000);
            // publiclyVisible: false — TryGetBuffer fails, so this must take the sequential path
            // rather than silently copying the whole source.
            using var ms = new MemoryStream(csv, 0, csv.Length, writable: false, publiclyVisible: false);

            List<Row> actual = await DrainAsync(Excel.ParseCsvParallelAsync<Row>(ms, degreeOfParallelism: 4, ct: TestContext.Current.CancellationToken));

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

                List<Row> actual = await DrainAsync(Excel.ParseCsvParallelAsync<Row>(fs, degreeOfParallelism: 8, ct: TestContext.Current.CancellationToken));

                Assert.Equal(20_000, actual.Count);
                Assert.Equal(19_999, actual[^1].Age);
                // The borrowed handle must outlive enumeration: the caller still owns this stream.
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
                // A BufferedStream over a FileStream is seekable with a known Length, but it is
                // neither a FileStream nor a MemoryStream — the resolver must decline it and fall
                // back rather than reach through to something it does not understand.
                await using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
                await using var buffered = new BufferedStream(fs);

                List<Row> actual = await DrainAsync(Excel.ParseCsvParallelAsync<Row>(buffered, degreeOfParallelism: 8, ct: TestContext.Current.CancellationToken));

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
            // Shared fixture (TestUtils.cs) — hides seekability from a stream that has it, so the
            // fallback path can be exercised without depending on a real pipe or socket.
            await using var pipe = new NonSeekableStream(csv);

            List<Row> actual = await DrainAsync(Excel.ParseCsvParallelAsync<Row>(pipe, degreeOfParallelism: 8, ct: TestContext.Current.CancellationToken));

            Assert.Equal(20_000, actual.Count);
            Assert.Equal(19_999, actual[^1].Age);
        }
    }
}
