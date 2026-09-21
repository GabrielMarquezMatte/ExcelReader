using System.Collections.Concurrent;
using System.Globalization;
using System.Text;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Parser.Internal;
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

        public ref struct Order
        {
            [ExcelColumn("Region")]
            public ReadOnlySpan<byte> Region { get; set; }

            [ExcelColumn("Units")]
            public int Units { get; set; }
        }

        private static byte[] OrdersCsv(int rows)
        {
            var sb = new StringBuilder();
            sb.Append("Region,Units\n");
            for (int i = 0; i < rows; i++)
            {
                sb.Append(CultureInfo.InvariantCulture, $"region{i % 4},{i + 1}\n");
            }
            return Encoding.UTF8.GetBytes(sb.ToString());
        }

        [Theory]
        [InlineData(1)]
        [InlineData(4)]
        [InlineData(16)]
        public async Task ForEachMapped_BindsEveryColumn(int dop)
        {
            const int rows = 20_000;
            byte[] csv = OrdersCsv(rows);
            long units = 0;
            long regionBytes = 0;
            long count = 0;

            await Excel.ForEachCsvParallelAsync(
                csv.AsMemory(),
                CsvModelMap.FromAttributes<Order>(),
                order =>
                {
                    Interlocked.Add(ref units, order.Units);
                    Interlocked.Add(ref regionBytes, order.Region.Length);
                    Interlocked.Increment(ref count);
                },
                new CsvParallelOptions { DegreeOfParallelism = dop, HeaderRow = 1 },
                TestContext.Current.CancellationToken);

            Assert.Equal(rows, count);
            Assert.Equal(ExpectedUnits(rows), units);
            Assert.Equal(rows * 7L, regionBytes);
        }

        [Fact]
        public async Task ForEachMapped_MatchesSequentialParser()
        {
            const int rows = 20_000;
            byte[] csv = OrdersCsv(rows);
            long parallel = 0;

            await Excel.ForEachCsvParallelAsync(
                csv.AsMemory(),
                CsvModelMap.FromAttributes<Order>(),
                order => Interlocked.Add(ref parallel, order.Units),
                new CsvParallelOptions { DegreeOfParallelism = 8, HeaderRow = 1 },
                TestContext.Current.CancellationToken);

            long sequential = 0;
            await Excel.ForEachCsvParallelAsync(
                csv.AsMemory(),
                CsvModelMap.FromAttributes<Order>(),
                order => sequential += order.Units,
                new CsvParallelOptions { DegreeOfParallelism = 1, HeaderRow = 1 },
                TestContext.Current.CancellationToken);

            Assert.Equal(sequential, parallel);
            Assert.Equal(ExpectedUnits(rows), parallel);
        }

        [Fact]
        public async Task ForEachMapped_HeaderRowZeroWithNamedMap_Throws()
        {
            byte[] csv = OrdersCsv(8);

            await Assert.ThrowsAsync<ArgumentException>(
                () => Excel.ForEachCsvParallelAsync(
                    csv.AsMemory(),
                    CsvModelMap.FromAttributes<Order>(),
                    static _ => { },
                    new CsvParallelOptions { DegreeOfParallelism = 4, HeaderRow = 0 },
                    TestContext.Current.CancellationToken));
        }

        [Fact]
        public async Task ForEachMapped_NullMap_Throws()
        {
            byte[] csv = OrdersCsv(8);

            await Assert.ThrowsAsync<ArgumentNullException>(
                () => Excel.ForEachCsvParallelAsync<Order>(
                    csv.AsMemory(),
                    null!,
                    static _ => { },
                    null,
                    TestContext.Current.CancellationToken));
        }

        private sealed class ChunkProbe
        {
            private int _chunks;

            internal int Chunks => Volatile.Read(ref _chunks);

            internal CsvAggregation<byte> Wrap(CsvAggregation<byte> inner)
            {
                Func<byte> seed = inner.Seed;
                return new CsvAggregation<byte>
                {
                    Seed = () =>
                    {
                        Interlocked.Increment(ref _chunks);
                        return seed();
                    },
                    Accumulate = inner.Accumulate,
                    Combine = inner.Combine,
                };
            }
        }

        private static Task<byte> ChunkedRecord(byte[] csv, CsvRecordAction<Sale> body, ChunkProbe probe, int dop, int chunkSize)
        {
            return ParallelCsvProcessor.RunWithChunkSizeAsync(
                csv.AsMemory(),
                probe.Wrap(RecordCallback<Sale>.For(body)),
                null,
                new CsvParallelOptions { DegreeOfParallelism = dop, HeaderRow = 1 },
                chunkSize,
                TestContext.Current.CancellationToken);
        }

        private static Task<byte> ChunkedMapped(byte[] csv, CsvRecordAction<Order> body, ChunkProbe probe, int dop, int chunkSize)
        {
            return ParallelCsvProcessor.RunWithChunkSizeAsync(
                csv.AsMemory(),
                probe.Wrap(MappedCallback<Order>.Unbound),
                MappedCallback<Order>.Binder(CsvModelMap.FromAttributes<Order>(), 1, body),
                new CsvParallelOptions { DegreeOfParallelism = dop, HeaderRow = 1 },
                chunkSize,
                TestContext.Current.CancellationToken);
        }

        private const int AlignedRowBytes = 16;

        private static byte[] AlignedOrdersCsv(int rows, string terminator)
        {
            var sb = new StringBuilder();
            sb.Append("Region,Units").Append(terminator);
            int digits = AlignedRowBytes - 3 - terminator.Length;
            for (int i = 0; i < rows; i++)
            {
                sb.Append(CultureInfo.InvariantCulture, $"r{i % 4},{(i + 1).ToString($"D{digits}", CultureInfo.InvariantCulture)}{terminator}");
            }
            return Encoding.UTF8.GetBytes(sb.ToString());
        }

        [Theory]
        [InlineData(1)]
        [InlineData(4)]
        [InlineData(16)]
        public async Task ForEachRecord_Chunked_AlignedBoundaries_DeliversExactlyOnce(int dop)
        {
            const int rows = 8192;
            byte[] csv = AlignedOrdersCsv(rows, "\n");
            var probe = new ChunkProbe();
            long total = 0;
            long count = 0;

            await ChunkedRecord(
                csv,
                sale =>
                {
                    Interlocked.Add(ref total, sale.Units);
                    Interlocked.Increment(ref count);
                },
                probe,
                dop,
                chunkSize: AlignedRowBytes * 256);

            Assert.True(probe.Chunks > 1, $"expected the source to partition, parsed {probe.Chunks} chunk(s)");
            Assert.Equal(rows, count);
            Assert.Equal(ExpectedUnits(rows), total);
        }

        [Theory]
        [InlineData(1)]
        [InlineData(4)]
        [InlineData(16)]
        public async Task ForEachMapped_Chunked_AlignedBoundaries_DeliversExactlyOnce(int dop)
        {
            const int rows = 8192;
            byte[] csv = AlignedOrdersCsv(rows, "\n");
            var probe = new ChunkProbe();
            long total = 0;
            long count = 0;

            await ChunkedMapped(
                csv,
                order =>
                {
                    Interlocked.Add(ref total, order.Units);
                    Interlocked.Increment(ref count);
                },
                probe,
                dop,
                chunkSize: AlignedRowBytes * 256);

            Assert.True(probe.Chunks > 1, $"expected the source to partition, parsed {probe.Chunks} chunk(s)");
            Assert.Equal(rows, count);
            Assert.Equal(ExpectedUnits(rows), total);
        }

        [Fact]
        public async Task ForEachMapped_Chunked_CrLfBoundarySplit_DeliversExactlyOnce()
        {
            const int rows = 8192;
            byte[] csv = AlignedOrdersCsv(rows, "\r\n");
            var probe = new ChunkProbe();
            long count = 0;

            await ChunkedMapped(
                csv,
                _ => Interlocked.Increment(ref count),
                probe,
                dop: 4,
                chunkSize: (AlignedRowBytes * 256) - 1);

            Assert.True(probe.Chunks > 1, $"expected the source to partition, parsed {probe.Chunks} chunk(s)");
            Assert.Equal(rows, count);
        }

        [Fact]
        public async Task ForEachMapped_Chunked_VariableLengthRows_DeliversExactlyOnce()
        {
            const int rows = 20_000;
            byte[] csv = OrdersCsv(rows);
            var probe = new ChunkProbe();
            long total = 0;
            long count = 0;

            await ChunkedMapped(
                csv,
                order =>
                {
                    Interlocked.Add(ref total, order.Units);
                    Interlocked.Increment(ref count);
                },
                probe,
                dop: 4,
                chunkSize: 3000);

            Assert.True(probe.Chunks > 1, $"expected the source to partition, parsed {probe.Chunks} chunk(s)");
            Assert.Equal(rows, count);
            Assert.Equal(ExpectedUnits(rows), total);
        }

        [Fact]
        public async Task ForEachMapped_Chunked_RaggedQuoteFreeRows_MatchesSinglePartition()
        {
            const int rows = 20_000;
            var sb = new StringBuilder();
            sb.Append("Region,Units\r\n");
            for (int i = 0; i < rows; i++)
            {
                sb.Append(CultureInfo.InvariantCulture, $"region{i % 4},{i + 1}");
                sb.Append(i % 3 == 0 ? "\n" : "\r\n");
                if (i % 5 == 0)
                {
                    sb.Append('\n');
                }
            }
            byte[] csv = Encoding.UTF8.GetBytes(sb.ToString());

            long chunked = 0;
            var probe = new ChunkProbe();
            await ChunkedMapped(csv, _ => Interlocked.Increment(ref chunked), probe, dop: 4, chunkSize: 2048);

            long single = 0;
            var whole = new ChunkProbe();
            await ChunkedMapped(csv, _ => single++, whole, dop: 1, chunkSize: csv.Length);

            Assert.True(probe.Chunks > 1, $"expected the source to partition, parsed {probe.Chunks} chunk(s)");
            Assert.Equal(1, whole.Chunks);
            Assert.Equal(single, chunked);
        }

        [Fact]
        public async Task ForEachMapped_Chunked_QuotedFields_DeliversSupersetOfSequential()
        {
            const int rows = 4000;
            var sb = new StringBuilder();
            sb.Append("Region,Units\n");
            for (int i = 0; i < rows; i++)
            {
                sb.Append(CultureInfo.InvariantCulture, $"\"north,\nsouth {i}\",{i + 1}\n");
            }
            byte[] csv = Encoding.UTF8.GetBytes(sb.ToString());
            var probe = new ChunkProbe();
            var delivered = new ConcurrentBag<string>();

            await ChunkedMapped(
                csv,
                order => delivered.Add($"{Encoding.UTF8.GetString(order.Region)}|{order.Units}"),
                probe,
                dop: 4,
                chunkSize: 2048);

            Assert.True(probe.Chunks > 1, $"expected the source to partition, parsed {probe.Chunks} chunk(s)");
            var seen = new HashSet<string>(delivered, StringComparer.Ordinal);
            for (int i = 0; i < rows; i++)
            {
                string expected = string.Create(CultureInfo.InvariantCulture, $"north,\nsouth {i}|{i + 1}");
                Assert.Contains(expected, seen);
            }
        }
    }
}
