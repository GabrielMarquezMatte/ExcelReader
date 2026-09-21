using System.Globalization;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Tests
{
    public class ForEachCsvParallelTests
    {
        public readonly ref struct Sale : ICsvRecord<Sale>
        {
            private Sale(int units)
            {
                Units = units;
            }

            public int Units { get; }

            public static bool TryParse(Row row, out Sale record)
            {
                record = default;
                if (row.ColumnCount < 2)
                {
                    return false;
                }
                if (!row[1].TryParse(CultureInfo.InvariantCulture, out int units))
                {
                    return false;
                }
                record = new Sale(units);
                return true;
            }
        }

        private static byte[] SalesCsv(int rows)
        {
            var sb = new StringBuilder();
            sb.Append("Name,Units\n");
            for (int i = 0; i < rows; i++)
            {
                sb.Append(CultureInfo.InvariantCulture, $"name{i},{i + 1}\n");
            }
            return Encoding.UTF8.GetBytes(sb.ToString());
        }

        private static long ExpectedUnits(int rows)
        {
            return (long)rows * (rows + 1) / 2;
        }

        [Theory]
        [InlineData(1)]
        [InlineData(4)]
        [InlineData(16)]
        public async Task ForEachRecord_Memory_SeesEveryRecordExactlyOnce(int dop)
        {
            const int rows = 20_000;
            byte[] csv = SalesCsv(rows);
            long total = 0;
            long count = 0;
            var options = new CsvParallelOptions { DegreeOfParallelism = dop, HeaderRow = 1 };

            await Excel.ForEachCsvParallelAsync<Sale>(
                csv.AsMemory(),
                sale =>
                {
                    Interlocked.Add(ref total, sale.Units);
                    Interlocked.Increment(ref count);
                },
                options,
                TestContext.Current.CancellationToken);

            Assert.Equal(rows, count);
            Assert.Equal(ExpectedUnits(rows), total);
        }

        [Fact]
        public async Task ForEachRecord_Path_MatchesMemory()
        {
            const int rows = 20_000;
            string path = Path.Combine(Path.GetTempPath(), $"foreach-{Guid.NewGuid():N}.csv");
            await File.WriteAllBytesAsync(path, SalesCsv(rows), TestContext.Current.CancellationToken);
            long total = 0;
            try
            {
                await Excel.ForEachCsvParallelAsync<Sale>(
                    path,
                    sale => Interlocked.Add(ref total, sale.Units),
                    new CsvParallelOptions { DegreeOfParallelism = 4, HeaderRow = 1 },
                    TestContext.Current.CancellationToken);
            }
            finally
            {
                File.Delete(path);
            }

            Assert.Equal(ExpectedUnits(rows), total);
        }

        [Fact]
        public async Task ForEachRecord_Stream_MatchesMemory()
        {
            const int rows = 20_000;
            using var stream = new MemoryStream(SalesCsv(rows), writable: false);
            long total = 0;

            await Excel.ForEachCsvParallelAsync<Sale>(
                stream,
                sale => Interlocked.Add(ref total, sale.Units),
                new CsvParallelOptions { DegreeOfParallelism = 4, HeaderRow = 1 },
                TestContext.Current.CancellationToken);

            Assert.Equal(ExpectedUnits(rows), total);
        }

        [Fact]
        public async Task ForEachRecord_TinySource_FallsBackToSequentialInOrder()
        {
            byte[] csv = SalesCsv(5);
            var seen = new List<int>();

            await Excel.ForEachCsvParallelAsync<Sale>(
                csv.AsMemory(),
                sale => seen.Add(sale.Units),
                new CsvParallelOptions { DegreeOfParallelism = 8, HeaderRow = 1 },
                TestContext.Current.CancellationToken);

            Assert.Equal([1, 2, 3, 4, 5], seen);
        }

        [Fact]
        public async Task ForEachRecord_NullBody_Throws()
        {
            byte[] csv = SalesCsv(4);
            await Assert.ThrowsAsync<ArgumentNullException>(
                () => Excel.ForEachCsvParallelAsync<Sale>(csv.AsMemory(), null!, null, TestContext.Current.CancellationToken));
        }

        [Fact]
        public async Task ForEachRecord_BodyThrows_SurfacesFromTask()
        {
            byte[] csv = SalesCsv(20_000);

            await Assert.ThrowsAsync<InvalidOperationException>(
                () => Excel.ForEachCsvParallelAsync<Sale>(
                    csv.AsMemory(),
                    static _ => throw new InvalidOperationException("boom"),
                    new CsvParallelOptions { DegreeOfParallelism = 4, HeaderRow = 1 },
                    TestContext.Current.CancellationToken));
        }

        [Fact]
        public async Task ForEachRecord_Cancelled_Throws()
        {
            byte[] csv = SalesCsv(200_000);
            using var cts = new CancellationTokenSource();

            await Assert.ThrowsAnyAsync<OperationCanceledException>(
                () => Excel.ForEachCsvParallelAsync<Sale>(
                    csv.AsMemory(),
                    _ => cts.Cancel(),
                    new CsvParallelOptions { DegreeOfParallelism = 4, HeaderRow = 1 },
                    cts.Token));
        }
    }
}
