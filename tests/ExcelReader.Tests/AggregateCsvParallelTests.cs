using System.Collections.Concurrent;
using System.Globalization;
using System.Text;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Tests
{
    public class AggregateCsvParallelTests
    {
        private readonly struct RowText : ICsvRecord<RowText>
        {
            public string Text { get; init; }

            public static bool TryParse(Row row, out RowText record)
            {
                record = new RowText { Text = Render(row) };
                return true;
            }
        }

        private ref struct NumberedRow : ICsvRecord<NumberedRow>
        {
            public ReadOnlySpan<byte> Name;
            public int Id;

            public static bool TryParse(Row row, out NumberedRow record)
            {
                record = default;
                if (!row[1].TryParse(CultureInfo.InvariantCulture, out int id))
                {
                    return false;
                }
                record = new NumberedRow { Name = row[0].Value, Id = id };
                return true;
            }
        }

        private sealed class RowLog : ICsvAccumulator<RowLog, RowText>, ICsvAccumulator<RowLog, NumberedRow>
        {
            public List<string> Rows { get; } = [];

            public void Add(RowText model)
            {
                Rows.Add(model.Text);
            }

            public void Add(NumberedRow model)
            {
                Rows.Add($"{Encoding.UTF8.GetString(model.Name)}#{model.Id}");
            }

            public void Merge(RowLog following)
            {
                Rows.AddRange(following.Rows);
            }
        }

        private static readonly CsvAggregation<List<string>> Collecting = new()
        {
            Seed = static () => [],
            Accumulate = static (ref List<string> rows, Row row) => rows.Add(Render(row)),
            Combine = static (left, right) =>
            {
                left.AddRange(right);
                return left;
            },
        };

        private static string Render(Row row)
        {
            var sb = new StringBuilder();
            foreach (RowCell cell in row.Cells)
            {
                sb.Append(cell.ColumnIndex).Append(':').Append(cell.Value.GetString()).Append('|');
            }
            return sb.ToString();
        }

        private static List<string> Sequential(byte[] csv, int headerRow = 0)
        {
            using CsvReader reader = Excel.FromCsv(csv);
            var rows = new List<string>();
            foreach (Row row in reader)
            {
                rows.Add(Render(row));
            }
            return rows.Count > headerRow ? rows[headerRow..] : [];
        }

        private static Task<List<string>> WithChunkSize(byte[] csv, int dop, int chunkSize, int headerRow = 0, CsvAggregation<List<string>>? aggregation = null)
        {
            return ParallelCsvProcessor.RunWithChunkSizeAsync(
                csv.AsMemory(),
                aggregation ?? Collecting,
                null,
                new CsvParallelOptions { DegreeOfParallelism = dop, HeaderRow = headerRow },
                chunkSize,
                TestContext.Current.CancellationToken);
        }

        public static TheoryData<string> Corpus()
        {
            return new TheoryData<string>
            {
                "a,b,c\nAda,36,x\nBob,41,y\nCid,7,z\n",
                "a,b,c\nAda,36,x\nBob,41,y",
                "a,b,c\r\nAda,36,x\r\nBob,41,y\r\n",
                "a,b,c\n\"multi\nline\",1,x\n\"another\none\",2,y\n",
                "a,b,c\n\"say \"\"hi\"\"\",1,x\n\"and \"\"bye\"\"\",2,y\n",
                "a,b,c\n\"a,b\",1,x\n\"c,d\",2,y\n",
                "a,b,c\n,1,\n,2,\n",
                "a,b,c\n\"" + new string('q', 400) + "\",1,x\n",
                "a,b,c\nAda,36,x\n\nBob,41,y\n",
                "a,b,c\rAda,36,x\rBob,41,y\r",
            };
        }

        [Theory]
        [MemberData(nameof(Corpus))]
        public async Task MatchesSequentialAcrossEveryDegreeChunkSizeAndHeaderRow(string text)
        {
            byte[] csv = Encoding.UTF8.GetBytes(text);
            foreach (int headerRow in new[] { 0, 1, 2 })
            {
                List<string> expected = Sequential(csv, headerRow);
                foreach (int dop in new[] { 2, 3, 4, 8 })
                {
                    foreach (int chunkSize in new[] { 1, 2, 3, 7, 16, 64 })
                    {
                        Assert.Equal(expected, await WithChunkSize(csv, dop, chunkSize, headerRow));
                    }
                }
            }
        }

        [Fact]
        public async Task SweepsAChunkBoundaryAcrossEveryByteOffsetOfAQuotedNewlineFixture()
        {
            byte[] csv = Encoding.UTF8.GetBytes("\"a\nb\",1,\"c\nd\"\n\"e\nf\",2,\"g\nh\"\n\"i\nj\",3,\"k\nl\"\n");
            List<string> expected = Sequential(csv);

            for (int chunkSize = 1; chunkSize <= csv.Length; chunkSize++)
            {
                Assert.Equal(expected, await WithChunkSize(csv, dop: 4, chunkSize));
            }
        }

        [Fact]
        public async Task DiscardsExceptionsThrownWhileReadingAMisguessedPartition()
        {
            byte[] csv = Encoding.UTF8.GetBytes("ok,\"one\nBOOM,x\"\nok,two\nok,\"three\nBOOM,y\"\nok,four\n");
            List<string> expected = Sequential(csv);
            int misguessedDeliveries = 0;
            var throwOnBoom = new CsvAggregation<List<string>>
            {
                Seed = Collecting.Seed,
                Accumulate = (ref List<string> rows, Row row) =>
                {
                    if (row[0].Value.SequenceEqual("BOOM"u8))
                    {
                        Interlocked.Increment(ref misguessedDeliveries);
                        throw new InvalidOperationException("delivered a record that does not exist");
                    }
                    rows.Add(Render(row));
                },
                Combine = Collecting.Combine,
            };

            for (int chunkSize = 1; chunkSize <= csv.Length; chunkSize++)
            {
                Assert.Equal(expected, await WithChunkSize(csv, dop: 4, chunkSize, aggregation: throwOnBoom));
            }
            Assert.True(misguessedDeliveries > 0, "No chunk size produced a misguessed partition, so nothing was tested.");
        }

        [Fact]
        public async Task SurfacesAnExceptionThrownForARealRecord()
        {
            byte[] csv = Encoding.UTF8.GetBytes(string.Concat(Enumerable.Range(0, 200).Select(i => $"r{i},v\n")));
            var failAt150 = new CsvAggregation<int>
            {
                Seed = static () => 0,
                Accumulate = static (ref int count, Row row) =>
                {
                    if (row[0].Value.SequenceEqual("r150"u8))
                    {
                        throw new InvalidOperationException("r150");
                    }
                    count++;
                },
                Combine = static (left, right) => left + right,
            };

            var ex = await Assert.ThrowsAsync<InvalidOperationException>(() => ParallelCsvProcessor.RunWithChunkSizeAsync(
                csv.AsMemory(), failAt150, null, new CsvParallelOptions { DegreeOfParallelism = 4 }, 64, TestContext.Current.CancellationToken));

            Assert.Equal("r150", ex.Message);
        }

        [Fact]
        public async Task EmptyAndHeaderOnlyInputsReturnAFreshState()
        {
            foreach (string text in new[] { "", "only,a,header\n" })
            {
                Assert.Empty(await WithChunkSize(Encoding.UTF8.GetBytes(text), dop: 4, chunkSize: 4, headerRow: 1));
            }
        }

        [Fact]
        public async Task EveryPublicOverloadMatchesSequentialOnAPartitionableSource()
        {
            byte[] csv = LargeCsv(rows: 150_000);
            List<string> expected = Sequential(csv, headerRow: 1);
            var options = new CsvParallelOptions { DegreeOfParallelism = 4, HeaderRow = 1 };
            CancellationToken ct = TestContext.Current.CancellationToken;
            string path = Path.Combine(Path.GetTempPath(), $"exr-aggregate-{Guid.NewGuid():N}.csv");
            await File.WriteAllBytesAsync(path, csv, ct);
            try
            {
                Assert.Equal(expected, await Excel.AggregateCsvParallelAsync(csv.AsMemory(), Collecting, options, ct));
                Assert.Equal(expected, (await Excel.AggregateCsvParallelAsync<RowLog, RowText>(csv.AsMemory(), options, ct)).Rows);
                Assert.Equal(expected, await Excel.AggregateCsvParallelAsync(path, Collecting, options, ct));
                Assert.Equal(expected, (await Excel.AggregateCsvParallelAsync<RowLog, RowText>(path, options, ct)).Rows);

                await using (var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true))
                {
                    Assert.Equal(expected, await Excel.AggregateCsvParallelAsync(file, Collecting, options, ct));
                }
                await using (var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true))
                {
                    Assert.Equal(expected, (await Excel.AggregateCsvParallelAsync<RowLog, RowText>(file, options, ct)).Rows);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public async Task FallsBackToOneSequentialPassWithTheSameResult()
        {
            byte[] csv = LargeCsv(rows: 20_000);
            List<string> expected = Sequential(csv, headerRow: 1);
            CancellationToken ct = TestContext.Current.CancellationToken;

            var sequential = new CsvParallelOptions { DegreeOfParallelism = 1, HeaderRow = 1 };
            Assert.Equal(expected, await Excel.AggregateCsvParallelAsync(csv.AsMemory(), Collecting, sequential, ct));

            using var unpartitionable = new BufferedStream(new MemoryStream(csv, writable: false));
            Assert.Equal(expected, (await Excel.AggregateCsvParallelAsync<RowLog, RowText>(unpartitionable, new CsvParallelOptions { HeaderRow = 1 }, ct)).Rows);
        }

        [Fact]
        public async Task RejectsInvalidOptionsAndAggregations()
        {
            byte[] csv = "a\n"u8.ToArray();
            await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => Excel.AggregateCsvParallelAsync<RowLog, RowText>(csv, new CsvParallelOptions { DegreeOfParallelism = -1 }, TestContext.Current.CancellationToken));
            await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => Excel.AggregateCsvParallelAsync<RowLog, RowText>(csv, new CsvParallelOptions { HeaderRow = -1 }, TestContext.Current.CancellationToken));
            await Assert.ThrowsAsync<ArgumentNullException>(() => Excel.AggregateCsvParallelAsync(csv, (CsvAggregation<int>)null!, null, TestContext.Current.CancellationToken));
            await Assert.ThrowsAsync<ArgumentNullException>(() => Excel.AggregateCsvParallelAsync(csv, new CsvAggregation<int> { Seed = null!, Accumulate = static (ref int s, Row r) => s++, Combine = static (a, b) => a + b }, null, TestContext.Current.CancellationToken));
        }

        [Fact]
        public async Task ACanceledTokenStopsProcessing()
        {
            byte[] csv = LargeCsv(rows: 150_000);
            using var cts = new CancellationTokenSource();
            await cts.CancelAsync();

            await Assert.ThrowsAnyAsync<OperationCanceledException>(
                () => Excel.AggregateCsvParallelAsync<RowLog, RowText>(csv, new CsvParallelOptions { DegreeOfParallelism = 4 }, cts.Token));
        }

        [Fact]
        public async Task CallsTheFactoryOnceWithTheHeaderRecord()
        {
            byte[] csv = LargeCsv(rows: 2_000);
            List<string> expected = Sequential(csv, headerRow: 1);
            CancellationToken ct = TestContext.Current.CancellationToken;

            foreach (bool partitioned in new[] { true, false })
            {
                var calls = new ConcurrentBag<(string Header, bool Sequential)>();
                CsvAccumulateFactory<List<string>> factory = (Row header, bool sequential) =>
                {
                    calls.Add((Render(header), sequential));
                    return Collecting.Accumulate;
                };
                List<string> rows = partitioned
                    ? await ParallelCsvProcessor.RunWithChunkSizeAsync(csv.AsMemory(), Collecting, factory, new CsvParallelOptions { DegreeOfParallelism = 4, HeaderRow = 1 }, 4096, ct)
                    : await ParallelCsvProcessor.RunAsync(csv.AsMemory(), Collecting, factory, new CsvParallelOptions { DegreeOfParallelism = 1, HeaderRow = 1 }, ct);

                Assert.Equal(expected, rows);
                (string header, bool sequentialCall) = Assert.Single(calls);
                Assert.Equal("0:name|1:id|2:note|", header);
                Assert.Equal(!partitioned, sequentialCall);
            }
        }

        [Fact]
        public async Task CallsTheFactoryWithAnEmptyHeaderWhenThereIsNoHeaderRow()
        {
            byte[] csv = LargeCsv(rows: 500);
            CancellationToken ct = TestContext.Current.CancellationToken;
            var columnCounts = new ConcurrentBag<int>();
            CsvAccumulateFactory<List<string>> factory = (Row header, bool sequential) =>
            {
                columnCounts.Add(header.ColumnCount);
                return Collecting.Accumulate;
            };

            List<string> rows = await ParallelCsvProcessor.RunWithChunkSizeAsync(csv.AsMemory(), Collecting, factory, new CsvParallelOptions { DegreeOfParallelism = 4 }, 1024, ct);

            Assert.Equal(Sequential(csv), rows);
            Assert.Equal(0, Assert.Single(columnCounts));
        }

        [Fact]
        public async Task DoesNotCallTheFactoryWhenTheInputEndsBeforeTheHeader()
        {
            int calls = 0;
            CsvAccumulateFactory<List<string>> factory = (Row header, bool sequential) =>
            {
                Interlocked.Increment(ref calls);
                return Collecting.Accumulate;
            };

            Assert.Empty(await ParallelCsvProcessor.RunWithChunkSizeAsync(ReadOnlyMemory<byte>.Empty, Collecting, factory, new CsvParallelOptions { DegreeOfParallelism = 4, HeaderRow = 1 }, 4, TestContext.Current.CancellationToken));
            Assert.Equal(0, calls);
        }

        [Fact]
        public async Task RecordAccumulatorMatchesSequentialAcrossChunkSizes()
        {
            byte[] csv = LargeCsv(rows: 400);
            List<string> expected = Sequential(csv, headerRow: 1);
            CancellationToken ct = TestContext.Current.CancellationToken;
            foreach (int dop in new[] { 2, 4, 8 })
            {
                foreach (int chunkSize in new[] { 1, 7, 64, 1024 })
                {
                    RowLog log = await ParallelCsvProcessor.RunWithChunkSizeAsync(
                        csv.AsMemory(), RecordAggregation<RowLog, RowText>.Instance, null,
                        new CsvParallelOptions { DegreeOfParallelism = dop, HeaderRow = 1 }, chunkSize, ct);
                    Assert.Equal(expected, log.Rows);
                }
            }
        }

        [Fact]
        public async Task SkipsRecordsWhoseTryParseReturnsFalseAndKeepsSpansValidDuringAdd()
        {
            byte[] csv = "name,id\nada,1\nbob,oops\n\"cid\nmulti\",3\n"u8.ToArray();
            CancellationToken ct = TestContext.Current.CancellationToken;
            foreach (int chunkSize in new[] { 1, 2, 5, 64 })
            {
                RowLog log = await ParallelCsvProcessor.RunWithChunkSizeAsync(
                    csv.AsMemory(), RecordAggregation<RowLog, NumberedRow>.Instance, null,
                    new CsvParallelOptions { DegreeOfParallelism = 4, HeaderRow = 1 }, chunkSize, ct);
                Assert.Equal(["ada#1", "cid\nmulti#3"], log.Rows);
            }
        }

        private static byte[] LargeCsv(int rows)
        {
            var sb = new StringBuilder("name,id,note\n");
            for (int i = 0; i < rows; i++)
            {
                string name = i % 7 == 0 ? $"\"multi\nline {i}, quoted\"" : $"name{i}";
                sb.Append(CultureInfo.InvariantCulture, $"{name},{i},note number {i}\r\n");
            }
            return Encoding.UTF8.GetBytes(sb.ToString());
        }
    }
}
