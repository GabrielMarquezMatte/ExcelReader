using System.Globalization;
using System.Text;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Parser.SheetPipeline;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Writer.Xlsb;

namespace ExcelReader.Tests.Parser.SheetPipeline
{
    public class XlsbPipelineTests
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

        private static readonly RowAggregation<List<string>> Collecting = new()
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

        private static async Task<MemoryStream> BuildAsync(int rows)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var ms = new MemoryStream();
            await using (XlsbWorkbookWriter wb = XlsbWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                await using (XlsbRowWriter header = await sheet.StartRowAsync(ct))
                {
                    header.Write("Name");
                    header.Write("Qty");
                    header.Write("Day");
                }
                for (int i = 0; i < rows; i++)
                {
                    await using XlsbRowWriter row = await sheet.StartRowAsync(ct);
                    row.Write(i % 3 == 0 ? new string('x', 1 + (i % 40)) : string.Create(CultureInfo.InvariantCulture, $"item{i:D6}"));
                    row.Write(i);
                    row.Write(new DateTime(2024, 1, 1).AddDays(i % 400));
                }
                await sheet.EndAsync(ct);
                await wb.EndAsync(ct);
            }
            ms.Position = 0;
            return ms;
        }

        private static List<string> Sequential(XlsbWorkbook workbook, int headerRow)
        {
            var rows = new List<string>();
            foreach (Row row in workbook.FirstSheet)
            {
                rows.Add(Render(row));
            }
            return rows.Count > headerRow ? rows[headerRow..] : [];
        }

        [Theory]
        [InlineData(2, 96)]
        [InlineData(4, 96)]
        [InlineData(4, 700)]
        [InlineData(8, 4096)]
        [InlineData(4, 1 << 20)]
        public async Task MatchesSequential(int dop, int chunkBytes)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using MemoryStream ms = await BuildAsync(2_000);
            using XlsbWorkbook workbook = Excel.FromXlsb(ms);
            SheetPipelineSource source = SheetPipelineSource.Create(workbook.FirstSheet)!;

            foreach (int headerRow in new[] { 0, 1, 4 })
            {
                Assert.Equal(Sequential(workbook, headerRow), await SheetPipelineRunner.AggregateAsync(source, Collecting, null, headerRow, dop, chunkBytes, ct));
            }
        }

        [Fact]
        public async Task ShortSheetsYieldNothing()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using MemoryStream ms = await BuildAsync(0);
            using XlsbWorkbook workbook = Excel.FromXlsb(ms);
            SheetPipelineSource source = SheetPipelineSource.Create(workbook.FirstSheet)!;

            Assert.Empty(await SheetPipelineRunner.AggregateAsync(source, Collecting, null, 1, 4, 96, ct));
            Assert.Empty(await SheetPipelineRunner.AggregateAsync(source, Collecting, null, 3, 4, 96, ct));
        }

        [Theory]
        [InlineData(2)]
        [InlineData(8)]
        public async Task ParseAndForEachMatchSequentialThroughThePublicApi(int dop)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using MemoryStream ms = await BuildAsync(40_000);
            using XlsbWorkbook workbook = Excel.FromXlsb(ms);
            Assert.NotNull(SheetPipelineSource.TryCreate(workbook.FirstSheet, 2));
            ExcelParser<Sale> parser = ExcelParser.FromAttributes<Sale>();
            IEnumerable<Sale> sequential = parser.Parse(workbook.FirstSheet);
            List<(string?, int, DateTime)> expected = [.. sequential.Select(s => (s.Name, s.Qty, s.Day))];

            var parsed = new List<(string?, int, DateTime)>();
            await foreach (Sale sale in SheetParallel.ParseAsync(workbook.FirstSheet, parser, new SheetParallelOptions { DegreeOfParallelism = dop }, ct))
            {
                parsed.Add((sale.Name, sale.Qty, sale.Day));
            }
            Assert.Equal(expected, parsed);

            var each = new System.Collections.Concurrent.ConcurrentBag<(string?, int, DateTime)>();
            await SheetParallel.ForEachAsync(workbook.FirstSheet, parser, (Sale s) => each.Add((s.Name, s.Qty, s.Day)), new SheetParallelOptions { DegreeOfParallelism = dop, HeaderRow = 1 }, ct);
            Assert.Equal(expected.Count, each.Count);
            Assert.Equal(expected.Order(), each.Order());
        }
    }
}
