using System.Collections.Concurrent;
using System.Globalization;
using System.Text;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Parser.ParallelCsv;
using ExcelReader.Core.Parser.SheetPipeline;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsx;

namespace ExcelReader.Tests.Parser
{
    public class SheetParallelXlsxTests
    {
        private sealed class Sale
        {
            [ExcelColumn("Name")]
            public string? Name { get; set; }

            [ExcelColumn("Qty")]
            public int Qty { get; set; }

            [ExcelColumn("Day")]
            public DateTime Day { get; set; }
        }

        private readonly struct NameOnly : IRowRecord<NameOnly>
        {
            public string Name { get; init; }

            public static bool TryParse(Row row, out NameOnly record)
            {
                record = new NameOnly { Name = row[0].GetString() };
                return true;
            }
        }

        private sealed class Names : IRowAccumulator<Names, NameOnly>, IRowAccumulator<Names, Sale>
        {
            public List<string?> Items { get; } = [];

            public void Add(NameOnly model)
            {
                Items.Add(model.Name);
            }

            public void Add(Sale model)
            {
                Items.Add(model.Name);
            }

            public void Merge(Names following)
            {
                Items.AddRange(following.Items);
            }
        }

        private static async Task<MemoryStream> BuildAsync(int rows)
        {
            object?[][] data = new object?[rows + 1][];
            data[0] = ["Name", "Qty", "Day"];
            for (int i = 0; i < rows; i++)
            {
                data[i + 1] = [string.Create(CultureInfo.InvariantCulture, $"item{i:D6}"), i, new DateTime(2024, 1, 1).AddDays(i % 400)];
            }
            return await TypedWorkbook.BuildAsync(data);
        }

        private static IEnumerable<(string?, int, DateTime)> Keys(IEnumerable<Sale> sales)
        {
            return sales.Select(s => (s.Name, s.Qty, s.Day));
        }

        [Theory]
        [InlineData(2)]
        [InlineData(4)]
        [InlineData(8)]
        [InlineData(0)]
        public async Task EveryFamilyMatchesSequential(int dop)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using MemoryStream ms = await BuildAsync(30_000);
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);
            Assert.NotNull(SheetPipelineSource.TryCreate(workbook.FirstSheet, 2));
            ExcelParser<Sale> parser = ExcelParser.FromAttributes<Sale>();
            List<Sale> expected = [.. parser.Parse(workbook.FirstSheet)];
            var withHeader = new SheetParallelOptions { DegreeOfParallelism = dop, HeaderRow = 1 };

            var parsed = new List<Sale>();
            await foreach (Sale sale in SheetParallel.ParseAsync(workbook.FirstSheet, parser, new SheetParallelOptions { DegreeOfParallelism = dop }, ct))
            {
                parsed.Add(sale);
            }
            Assert.Equal(Keys(expected), Keys(parsed));

            Names byRecord = await SheetParallel.AggregateAsync<Names, NameOnly>(workbook.FirstSheet, withHeader, ct);
            Assert.Equal(expected.Select(s => s.Name), byRecord.Items, StringComparer.Ordinal);

            Names byModel = await SheetParallel.AggregateAsync<Names, Sale>(workbook.FirstSheet, parser, withHeader, ct);
            Assert.Equal(expected.Select(s => s.Name), byModel.Items, StringComparer.Ordinal);

            var each = new ConcurrentBag<(string?, int, DateTime)>();
            await SheetParallel.ForEachAsync(workbook.FirstSheet, parser, (Sale s) => each.Add((s.Name, s.Qty, s.Day)), withHeader, ct);
            Assert.Equal(Keys(expected).Order(), each.Order());

            var eachRecord = new ConcurrentBag<string>();
            await SheetParallel.ForEachAsync<NameOnly>(workbook.FirstSheet, r => eachRecord.Add(r.Name), withHeader, ct);
            Assert.Equal(expected.Select(s => s.Name).Order(StringComparer.Ordinal), eachRecord.Order(StringComparer.Ordinal));
        }

        [Fact]
        public async Task ParseAtSmallChunksMatchesSequentialIncludingEmptyRows()
        {
            var sb = new StringBuilder("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>Qty</t></is></c></row>");
            for (int r = 2; r <= 400; r++)
            {
                sb.Append(r % 7 == 0
                    ? string.Create(CultureInfo.InvariantCulture, $"<row r=\"{r}\"/>")
                    : string.Create(CultureInfo.InvariantCulture, $"<row r=\"{r}\"><c r=\"A{r}\" t=\"inlineStr\"><is><t>n{r}</t></is></c><c r=\"B{r}\"><v>{r}</v></c></row>"));
            }
            using MemoryStream ms = WorkbookBuilder.Build(sb.ToString());
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);
            ExcelParser<Sale> parser = ExcelParser.FromAttributes<Sale>();
            List<Sale> expected = [.. parser.Parse(workbook.FirstSheet)];
            SheetPipelineSource source = SheetPipelineSource.Create(workbook.FirstSheet)!;

            var actual = new List<Sale>();
            await foreach (Sale sale in SheetRowProjection<Sale>.ParseAsync(source, parser, isDate1904: false, 4, 96, TestContext.Current.CancellationToken))
            {
                actual.Add(sale);
            }

            Assert.Equal(399, expected.Count);
            Assert.Equal(Keys(expected), Keys(actual));
        }

        [Fact]
        public async Task AParseFailureReportsTheSheetRow()
        {
            var sb = new StringBuilder("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>Qty</t></is></c></row>");
            for (int r = 2; r <= 300; r++)
            {
                string qty = r == 211 ? "<c r=\"B211\" t=\"inlineStr\"><is><t>oops</t></is></c>" : string.Create(CultureInfo.InvariantCulture, $"<c r=\"B{r}\"><v>{r}</v></c>");
                sb.Append(CultureInfo.InvariantCulture, $"<row r=\"{r}\"><c r=\"A{r}\" t=\"inlineStr\"><is><t>n{r}</t></is></c>{qty}</row>");
            }
            using MemoryStream ms = WorkbookBuilder.Build(sb.ToString());
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);
            ExcelParser<Sale> parser = ExcelParser.FromAttributes<Sale>(new ExcelParserOptions { ThrowOnParseFailure = true });
            ExcelParseException sequential = Assert.Throws<ExcelParseException>(() => parser.Parse(workbook.FirstSheet).ToList());
            SheetPipelineSource source = SheetPipelineSource.Create(workbook.FirstSheet)!;

            ExcelParseException parallel = await Assert.ThrowsAsync<ExcelParseException>(async () =>
            {
                await foreach (Sale _ in SheetRowProjection<Sale>.ParseAsync(source, parser, isDate1904: false, 4, 128, TestContext.Current.CancellationToken))
                {
                }
            });

            Assert.Equal(211, sequential.Row);
            Assert.Equal(sequential.Row, parallel.Row);
            Assert.Equal(sequential.ColumnName, parallel.ColumnName);
        }

        [Fact]
        public async Task AParseFailureAfterPipelineRecoveryReportsTheSheetRow()
        {
            var sb = new StringBuilder("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>Qty</t></is></c></row>");
            for (int r = 2; r <= 300; r++)
            {
                string name = r == 20 ? "<![CDATA[ <row r=\"9\"> ]]>" : string.Create(CultureInfo.InvariantCulture, $"n{r}");
                string qty = r == 211 ? "<c r=\"B211\" t=\"inlineStr\"><is><t>oops</t></is></c>" : string.Create(CultureInfo.InvariantCulture, $"<c r=\"B{r}\"><v>{r}</v></c>");
                sb.Append(CultureInfo.InvariantCulture, $"<row r=\"{r}\"><c r=\"A{r}\" t=\"inlineStr\"><is><t>{name}</t></is></c>{qty}</row>");
            }
            using MemoryStream ms = WorkbookBuilder.Build(sb.ToString());
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);
            ExcelParser<Sale> parser = ExcelParser.FromAttributes<Sale>(new ExcelParserOptions { ThrowOnParseFailure = true });
            ExcelParseException sequential = Assert.Throws<ExcelParseException>(() => parser.Parse(workbook.FirstSheet).ToList());
            SheetPipelineSource source = SheetPipelineSource.Create(workbook.FirstSheet)!;
            Assert.Equal(211, sequential.Row);

            foreach (int chunkBytes in new[] { 64, 90, 96, 150, 400 })
            {
                ExcelParseException parallel = await Assert.ThrowsAsync<ExcelParseException>(async () =>
                {
                    await foreach (Sale _ in SheetRowProjection<Sale>.ParseAsync(source, parser, isDate1904: false, 4, chunkBytes, TestContext.Current.CancellationToken))
                    {
                    }
                });
                Assert.Equal(sequential.Row, parallel.Row);
                Assert.Equal(sequential.ColumnName, parallel.ColumnName);
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task APartitionedAggregateFailureReportsTheSheetRow(bool needsRecovery)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var sb = new StringBuilder("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>Qty</t></is></c></row>");
            for (int r = 2; r <= 300; r++)
            {
                string name = needsRecovery && r == 20 ? "<![CDATA[ <row r=\"9\"> ]]>" : string.Create(CultureInfo.InvariantCulture, $"n{r}");
                string qty = r == 211 ? "<c r=\"B211\" t=\"inlineStr\"><is><t>oops</t></is></c>" : string.Create(CultureInfo.InvariantCulture, $"<c r=\"B{r}\"><v>{r}</v></c>");
                sb.Append(CultureInfo.InvariantCulture, $"<row r=\"{r}\"><c r=\"A{r}\" t=\"inlineStr\"><is><t>{name}</t></is></c>{qty}</row>");
            }
            using MemoryStream ms = WorkbookBuilder.Build(sb.ToString());
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);
            ExcelParser<Sale> parser = ExcelParser.FromAttributes<Sale>(new ExcelParserOptions { ThrowOnParseFailure = true });
            ExcelParseException sequential = await Assert.ThrowsAsync<ExcelParseException>(() => SheetParallel.AggregateAsync<Names, Sale>(
                workbook.FirstSheet, parser, new SheetParallelOptions { DegreeOfParallelism = 1, HeaderRow = 1 }, ct));
            Assert.Equal(211, sequential.Row);
            SheetPipelineSource source = SheetPipelineSource.Create(workbook.FirstSheet)!;

            foreach (int chunkBytes in new[] { 64, 90, 96, 150, 400 })
            {
                ExcelParseException parallel = await Assert.ThrowsAsync<ExcelParseException>(() => SheetPipelineRunner.AggregateAsync(
                    source, MappedAggregation<Names, Sale>.Unbound, MappedAggregation<Names, Sale>.Binder(parser, parser.Info, false, 1), 1, 4, chunkBytes, ct));
                Assert.Equal(sequential.Row, parallel.Row);
                Assert.Equal(sequential.ColumnName, parallel.ColumnName);
            }
        }

        [Fact]
        public async Task ACallbackExceptionReachesTheCaller()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using MemoryStream ms = await BuildAsync(30_000);
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);
            ExcelParser<Sale> parser = ExcelParser.FromAttributes<Sale>();

            InvalidOperationException ex = await Assert.ThrowsAsync<InvalidOperationException>(() => SheetParallel.ForEachAsync(
                workbook.FirstSheet, parser,
                (Sale s) =>
                {
                    if (s.Qty == 12_345)
                    {
                        throw new InvalidOperationException("boom");
                    }
                },
                new SheetParallelOptions { DegreeOfParallelism = 4, HeaderRow = 1 }, ct));

            Assert.Equal("boom", ex.Message);
        }

        [Fact]
        public async Task LeavingTheEnumerationEarlyStopsThePipeline()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var tracked = new TrackingStream();
            using (MemoryStream built = await BuildAsync(30_000))
            {
                await built.CopyToAsync(tracked, ct);
            }
            tracked.Position = 0;
            XlsxWorkbook workbook = Excel.FromXlsx(tracked, leaveOpen: false);
            ExcelParser<Sale> parser = ExcelParser.FromAttributes<Sale>();

            await foreach (Sale _ in SheetParallel.ParseAsync(workbook.FirstSheet, parser, new SheetParallelOptions { DegreeOfParallelism = 4 }, ct))
            {
                break;
            }
            workbook.Dispose();

            Assert.True(tracked.Disposed);
        }

        [Fact]
        public async Task DisposingTheWorkbookMidReadIsSafe()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using MemoryStream ms = await BuildAsync(30_000);
            XlsxWorkbook workbook = Excel.FromXlsx(ms);
            ExcelParser<Sale> parser = ExcelParser.FromAttributes<Sale>();
            IExcelSheet sheet = workbook.FirstSheet;
            List<Sale> expected = [.. parser.Parse(workbook.FirstSheet)];

            var actual = new List<Sale>();
            await foreach (Sale sale in SheetParallel.ParseAsync(sheet, parser, new SheetParallelOptions { DegreeOfParallelism = 4 }, ct))
            {
                actual.Add(sale);
                if (actual.Count == 1)
                {
                    workbook.Dispose();
                }
            }
            Assert.Equal(Keys(expected), Keys(actual));

            await Assert.ThrowsAsync<ObjectDisposedException>(async () =>
            {
                await foreach (Sale _ in SheetParallel.ParseAsync(sheet, parser, new SheetParallelOptions { DegreeOfParallelism = 4 }, ct))
                {
                }
            });
        }

        private static readonly RowAggregation<int> Counting = new()
        {
            Seed = static () => 0,
            Accumulate = static (ref int count, Row _) => count++,
            Combine = static (left, right) => left + right,
        };

        private static string Outcome(Func<int> read)
        {
            try
            {
                return string.Create(CultureInfo.InvariantCulture, $"rows={read()}");
            }
            catch (Exception ex) when (ex is InvalidDataException or ExcelLimitExceededException)
            {
                return ex.GetType().Name;
            }
        }

        private static int CountSequentially(XlsxWorkbook workbook)
        {
            int count = 0;
            foreach (Row _ in workbook.FirstSheet)
            {
                count++;
            }
            return count;
        }

        [Fact]
        public async Task ACorruptSheetEntryEndsLikeTheSequentialRead()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using MemoryStream ms = await BuildAsync(30_000);
            byte[] bytes = ms.ToArray();
            int entry = bytes.AsSpan().IndexOf("xl/worksheets/sheet1.xml"u8);
            Array.Fill(bytes, (byte)0xFF, entry + 300_000, 512);
            using XlsxWorkbook sequentialWorkbook = Excel.FromXlsx(bytes);
            using XlsxWorkbook parallelWorkbook = Excel.FromXlsx(bytes);
            var options = new SheetParallelOptions { DegreeOfParallelism = 4 };

            string sequential = Outcome(() => CountSequentially(sequentialWorkbook));
            string parallel = Outcome(() => SheetParallel.AggregateAsync(parallelWorkbook.FirstSheet, Counting, options, ct).GetAwaiter().GetResult());

            Assert.Equal(sequential, parallel);
            Assert.Equal(nameof(InvalidDataException), sequential);
        }

        [Fact]
        public async Task ExceedingTheDecompressedLimitMidSheetThrows()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using MemoryStream ms = await BuildAsync(30_000);
            byte[] bytes = ms.ToArray();
            var limits = new ExcelReaderOptions { MaxTotalDecompressedBytes = 600 * 1024 };
            using XlsxWorkbook sequentialWorkbook = Excel.FromXlsx(bytes, limits);
            using XlsxWorkbook parallelWorkbook = Excel.FromXlsx(bytes, limits);

            Assert.Throws<ExcelLimitExceededException>(() => CountSequentially(sequentialWorkbook));
            await Assert.ThrowsAsync<ExcelLimitExceededException>(
                () => SheetParallel.AggregateAsync(parallelWorkbook.FirstSheet, Counting, new SheetParallelOptions { DegreeOfParallelism = 4 }, ct));
        }
    }
}
