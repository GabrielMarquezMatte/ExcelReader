using System.Collections.Concurrent;
using System.Text;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Parser
{
    public class SkipEmptyRowsTests
    {
        private sealed class Row
        {
            [ExcelRequired]
            public int Id { get; set; }
        }

        private static readonly ExcelParserOptions Skip = new() { SkipEmptyRows = true };

        private static CancellationToken Ct => TestContext.Current.CancellationToken;

        // Row 3 holds only styled blank cells, row 4 only a value in an unmapped column.
        private static async Task<MemoryStream> BuildAsync(string lastId)
        {
            (string? Id, string? Note)[] rows = [("Id", "Note"), ("1", null), (null, null), (null, "x"), (lastId, "y")];
            var stream = new MemoryStream();
            await using (var writer = XlsxWorkbookWriter.Create(stream, leaveOpen: true))
            {
                int bold = writer.AddStyle(new CellStyle { Bold = true });
                await using var sheet = writer.AddSheet("S");
                await sheet.WriteRecordsAsync(rows, (XlsxRowWriter row, (string? Id, string? Note) item) =>
                {
                    row.Style(bold).Write(item.Id);
                    row.Style(bold).Write(item.Note);
                }, Ct);
            }
            stream.Position = 0;
            return stream;
        }

        private static int[] Ids(IEnumerable<Row> rows)
        {
            return [.. rows.Select(row => row.Id)];
        }

        private static MemoryStream Csv()
        {
            return new MemoryStream(Encoding.UTF8.GetBytes("Id,Note\n1,\n,\n,x\n2,y\n"));
        }

        [Fact]
        public async Task XlsxSkipsRowsWithoutAMappedValue()
        {
            await using var ms = await BuildAsync("2");
            await using var reader = await Excel.FromXlsxAsync(ms, ct: Ct);

            int[] ids = Ids(ExcelParser.FromAttributes<Row>(Skip).Parse(reader.FirstSheet));

            Assert.Equal([1, 2], ids);
        }

        [Fact]
        public async Task SkippedRowsStillCountTowardTheReportedRowNumber()
        {
            await using var ms = await BuildAsync("abc");
            await using var reader = await Excel.FromXlsxAsync(ms, ct: Ct);
            var parser = ExcelParser.FromAttributes<Row>(new ExcelParserOptions { SkipEmptyRows = true, ThrowOnParseFailure = true });

            ExcelParseException ex = Assert.Throws<ExcelParseException>(() => parser.Parse(reader.FirstSheet).ToArray());

            Assert.Contains("row 5", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public async Task EmptyRowsAreStillRefusedByDefault()
        {
            await using var ms = await BuildAsync("2");
            await using var reader = await Excel.FromXlsxAsync(ms, ct: Ct);

            ExcelParseException ex = Assert.Throws<ExcelParseException>(
                () => ExcelParser.FromAttributes<Row>().Parse(reader.FirstSheet).ToArray());

            Assert.Contains("row 3", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public async Task CsvSkipsRowsWithoutAMappedValue()
        {
            await using var reader = Excel.FromCsv(Csv());

            int[] ids = Ids(ExcelParser.FromAttributes<Row>(Skip).Parse(reader.FirstSheet));

            Assert.Equal([1, 2], ids);
        }

        [Fact]
        public async Task XlsxSheetParallelSkipsRowsWithoutAMappedValue()
        {
            await using var ms = await BuildAsync("2");
            await using var reader = await Excel.FromXlsxAsync(ms, ct: Ct);
            List<int> ids = [];

            await foreach (Row row in SheetParallel.ParseAsync(reader.FirstSheet, ExcelParser.FromAttributes<Row>(Skip), ct: Ct))
            {
                ids.Add(row.Id);
            }

            Assert.Equal([1, 2], ids);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task SheetParallelForEachSkipsRowsWithoutAMappedValue(bool csv)
        {
            await using var ms = csv ? Csv() : await BuildAsync("2");
            await using IExcelWorkbook reader = csv ? Excel.FromCsv(ms) : await Excel.FromXlsxAsync(ms, ct: Ct);
            ConcurrentBag<int> ids = [];

            await SheetParallel.ForEachAsync(
                reader.FirstSheet, ExcelParser.FromAttributes<Row>(Skip), row => ids.Add(row.Id), new SheetParallelOptions { HeaderRow = 1 }, Ct);

            Assert.Equal([1, 2], ids.Order());
        }
    }
}
