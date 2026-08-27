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

        private sealed class RequiredRow
        {
            [ExcelRequired]
            public int Id { get; set; }

            public string? Note { get; set; }
        }

        [SuppressMessage("Usage", "VSTHRD200:Use \"Async\" suffix for async methods",
            Justification = "Builds the enumerable synchronously; nothing is awaited here, so an Async suffix would misdescribe it.")]
        private static ParallelCsvEnumerable<Row> Build(byte[] csv, int dop, int chunkSizeOverride)
        {
            return Build<Row>(csv, dop, chunkSizeOverride, new ExcelParserConfig());
        }

        [SuppressMessage("Usage", "VSTHRD200:Use \"Async\" suffix for async methods",
            Justification = "Builds the enumerable synchronously; nothing is awaited here, so an Async suffix would misdescribe it.")]
        private static ParallelCsvEnumerable<TRow> Build<TRow>(byte[] csv, int dop, int chunkSizeOverride, ExcelParserConfig config)
        {
            using var headerReader = Excel.FromCsv(csv);
            CsvBoundColumnMap<TRow> map = CsvHeaderBinder.Bind<TRow>(
                headerReader, config, TypeMapper<TRow>.GetCsvInfo(), out long dataStart);

            var source = new CsvChunkSource(csv.AsMemory());
            CsvChunkPlan plan = CsvChunkPlan.Create(dataStart, csv.Length - dataStart, dop, chunkSizeOverride);
            return new ParallelCsvEnumerable<TRow>(
                source, plan, dataStart, map, TypeMapper<TRow>.GetCsvInfo(),
                CsvReaderOptions.Default, config, dop);
        }

        // The parallel path's chunk projectors number rows from 1 within their own chunk, so
        // ParallelCsvEnumerable has to renumber a carried failure to the row the sequential path would
        // have reported. Nothing but a differential assertion can pin that: it depends on the header
        // offset and on picking the right ExcelParseException constructor for the failure's shape.
        private static async Task AssertSameFailureAsSequentialAsync<TRow>(byte[] csv, ExcelParserConfig config)
        {
            using CsvReader sequentialReader = Excel.FromCsv(csv);
            ExcelParseException expected = Assert.Throws<ExcelParseException>(
                () => new ExcelParser<TRow>(config).Parse(sequentialReader).ToList());

            ExcelParseException actual = await Assert.ThrowsAsync<ExcelParseException>(async () =>
            {
                await foreach (TRow row in Build<TRow>(csv, dop: 4, chunkSizeOverride: 64, config))
                {
                    _ = row;
                }
            });

            Assert.Equal(expected.Row, actual.Row);
            Assert.Equal(expected.ColumnName, actual.ColumnName);
            Assert.Equal(expected.Message, actual.Message);
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
        public Task ReportsAParseFailureAtTheRowTheSequentialPathWouldReport()
        {
            var sb = new StringBuilder("Name,Age\n");
            for (int i = 0; i < 40; i++)
            {
                // Data record 18 (file row 19) carries an Age that cannot be parsed.
                string age = i == 17 ? "not-a-number" : i.ToString(CultureInfo.InvariantCulture);
                sb.Append(CultureInfo.InvariantCulture, $"name{i:D5},{age}\n");
            }

            return AssertSameFailureAsSequentialAsync<Row>(
                Encoding.UTF8.GetBytes(sb.ToString()),
                new ExcelParserConfig { ThrowOnParseFailure = true });
        }

        [Fact]
        public Task ReportsABlankRequiredCellTheWayTheSequentialPathDoes()
        {
            // A blank required cell is a *different* ExcelParseException shape from a parse failure —
            // different constructor, different message — and the renumbering has to preserve which.
            var sb = new StringBuilder("Id,Note\n");
            for (int i = 0; i < 40; i++)
            {
                string id = i == 22 ? string.Empty : i.ToString(CultureInfo.InvariantCulture);
                sb.Append(CultureInfo.InvariantCulture, $"{id},note{i:D4}\n");
            }

            return AssertSameFailureAsSequentialAsync<RequiredRow>(
                Encoding.UTF8.GetBytes(sb.ToString()), new ExcelParserConfig());
        }

        [Fact]
        public Task ReportsAParseFailureRowRelativeToANonDefaultHeaderRow()
        {
            // HeaderRow 3 shifts every data row's number by two, which pins the header offset in the
            // renumbering rather than only the +1 for the failing record itself.
            var sb = new StringBuilder("skip me\nskip me too\nName,Age\n");
            for (int i = 0; i < 40; i++)
            {
                string age = i == 11 ? "still-not-a-number" : i.ToString(CultureInfo.InvariantCulture);
                sb.Append(CultureInfo.InvariantCulture, $"name{i:D5},{age}\n");
            }

            return AssertSameFailureAsSequentialAsync<Row>(
                Encoding.UTF8.GetBytes(sb.ToString()),
                new ExcelParserConfig { HeaderRow = 3, ThrowOnParseFailure = true });
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
