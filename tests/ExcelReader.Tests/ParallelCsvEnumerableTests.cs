using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class ParallelCsvEnumerableTests
    {
        private sealed class Row
        {
            public string? Name { get; set; }
            public int Age { get; set; }
        }

        [SuppressMessage("Usage", "VSTHRD200:Use \"Async\" suffix for async methods",
            Justification = "Builds the enumerable synchronously; nothing is awaited here, so an Async suffix would misdescribe it.")]
        private static ParallelCsvEnumerable<Row> Build(byte[] csv, int dop, int chunkSizeOverride)
        {
            using var headerReader = Excel.FromCsv(csv);
            CsvBoundColumnMap<Row> map = CsvHeaderBinder.Bind<Row>(
                headerReader, new ExcelParserConfig(), TypeMapper<Row>.GetCsvInfo(), out long dataStart);

            var source = new CsvChunkSource(csv.AsMemory());
            CsvChunkPlan plan = CsvChunkPlan.Create(dataStart, csv.Length - dataStart, dop, chunkSizeOverride);
            return new ParallelCsvEnumerable<Row>(
                source, plan, dataStart, map, TypeMapper<Row>.GetCsvInfo(),
                CsvReaderOptions.Default, new ExcelParserConfig(), dop);
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

        [Fact]
        public async Task EmitsEveryRowInFileOrder()
        {
            byte[] csv = BuildCsv(500);
            var seen = new List<Row>();

            await foreach (Row row in Build(csv, dop: 4, chunkSizeOverride: 64))
            {
                seen.Add(row);
            }

            Assert.Equal(500, seen.Count);
            for (int i = 0; i < 500; i++)
            {
                Assert.Equal($"name{i:D5}", seen[i].Name);
                Assert.Equal(i, seen[i].Age);
            }
        }

        [Fact]
        public async Task CorrectsAChunkThatGuessedItsStartInsideAQuotedField()
        {
            // Every row's Name is a quoted field carrying a newline, so a chunk boundary landing
            // inside one guesses wrong and must be corrected by its predecessor.
            var sb = new StringBuilder("Name,Age\n");
            for (int i = 0; i < 200; i++)
            {
                sb.Append(CultureInfo.InvariantCulture, $"\"line\nbreak{i:D4}\",{i}\n");
            }
            byte[] csv = Encoding.UTF8.GetBytes(sb.ToString());
            var seen = new List<Row>();

            await foreach (Row row in Build(csv, dop: 4, chunkSizeOverride: 64))
            {
                seen.Add(row);
            }

            Assert.Equal(200, seen.Count);
            for (int i = 0; i < 200; i++)
            {
                Assert.Equal($"line\nbreak{i:D4}", seen[i].Name);
                Assert.Equal(i, seen[i].Age);
            }
        }

        [Fact]
        public async Task AwaitsWorkerCompletionWhenTheConsumerBreaksEarly()
        {
            byte[] csv = BuildCsv(5000);
            int taken = 0;

            IAsyncEnumerator<Row> e = Build(csv, dop: 4, chunkSizeOverride: 256)
                .GetAsyncEnumerator(TestContext.Current.CancellationToken);
            await using (e.ConfigureAwait(false))
            {
                while (await e.MoveNextAsync())
                {
                    taken++;
                    if (taken == 10)
                    {
                        break;
                    }
                }
            }

            Assert.Equal(10, taken);
            // Reaching here without a hang or an ObjectDisposedException from a still-running worker
            // is the assertion: DisposeAsync must have cancelled and joined every worker.
        }

        [Fact]
        public async Task PropagatesCancellation()
        {
            byte[] csv = BuildCsv(20000);
            using var cts = new CancellationTokenSource();
            int taken = 0;

            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
            {
                await foreach (Row row in Build(csv, dop: 4, chunkSizeOverride: 128).WithCancellation(cts.Token))
                {
                    taken++;
                    if (taken == 5)
                    {
                        await cts.CancelAsync();
                    }
                }
            });
        }

        [Fact]
        public async Task ProducesTheSameSequenceAtEveryDegreeOfParallelism()
        {
            byte[] csv = BuildCsv(1000);
            List<string> baseline = [];
            await foreach (Row row in Build(csv, dop: 1, chunkSizeOverride: 64))
            {
                baseline.Add($"{row.Name}|{row.Age}");
            }

            foreach (int dop in new[] { 2, 3, 5, 8 })
            {
                List<string> actual = [];
                await foreach (Row row in Build(csv, dop, chunkSizeOverride: 64))
                {
                    actual.Add($"{row.Name}|{row.Age}");
                }
                Assert.Equal(baseline, actual);
            }
        }
    }
}
