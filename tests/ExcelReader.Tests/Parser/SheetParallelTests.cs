using System.Globalization;
using System.Text;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Tests.Reader.Tables;

namespace ExcelReader.Tests.Parser
{
    public class SheetParallelTests
    {
        private sealed class Person
        {
            public string? Name { get; set; }

            public int Age { get; set; }
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

        private static List<string> Sequential(IExcelSheet sheet, int headerRow)
        {
            var rows = new List<string>();
            using IExcelRowEnumerator e = sheet.GetEnumerator();
            while (e.MoveNext())
            {
                rows.Add(Render(e.Current));
            }
            return rows.Count > headerRow ? rows[headerRow..] : [];
        }

        private static byte[] Csv(int rows)
        {
            var sb = new StringBuilder("Name,Age\n");
            for (int i = 0; i < rows; i++)
            {
                sb.Append(CultureInfo.InvariantCulture, $"name{i:D5},{i}\n");
            }
            return Encoding.UTF8.GetBytes(sb.ToString());
        }

        private static async Task<MemoryStream> XlsxAsync(int rows)
        {
            object?[][] data = new object?[rows + 1][];
            data[0] = ["Name", "Age"];
            for (int i = 0; i < rows; i++)
            {
                data[i + 1] = [string.Create(CultureInfo.InvariantCulture, $"name{i:D5}"), i];
            }
            return await TypedWorkbook.BuildAsync(data);
        }

        [Fact]
        public async Task CsvAggregateMatchesSequential()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using CsvReader reader = Excel.FromCsv(Csv(40_000));
            var options = new SheetParallelOptions { DegreeOfParallelism = 4, HeaderRow = 1 };

            List<string> actual = await SheetParallel.AggregateAsync(reader.FirstSheet, Collecting, options, ct);

            Assert.Equal(Sequential(reader.FirstSheet, 1), actual);
        }

        [Fact]
        public async Task CsvParseMatchesTheSequentialParser()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using CsvReader reader = Excel.FromCsv(Csv(40_000));
            ExcelParser<Person> parser = ExcelParser.FromAttributes<Person>();
            List<Person> expected = [.. parser.Parse(reader.FirstSheet)];

            var actual = new List<Person>();
            await foreach (Person person in SheetParallel.ParseAsync(reader.FirstSheet, parser, new SheetParallelOptions { DegreeOfParallelism = 4 }, ct))
            {
                actual.Add(person);
            }

            Assert.Equal(expected.Select(p => (p.Name, p.Age)), actual.Select(p => (p.Name, p.Age)));
        }

        [Fact]
        public async Task ACsvOverANonSeekableStreamFallsBackToSequential()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] csv = Csv(5_000);
            await using var pipe = new NonSeekableStream(csv);
            using CsvReader reader = Excel.FromCsv(pipe);
            using CsvReader reference = Excel.FromCsv(csv);

            List<string> actual = await SheetParallel.AggregateAsync(reader.FirstSheet, Collecting, new SheetParallelOptions { DegreeOfParallelism = 4, HeaderRow = 1 }, ct);

            Assert.Equal(Sequential(reference.FirstSheet, 1), actual);
        }

        [Fact]
        public async Task AnXlsSheetReadsSequentially()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var ms = new MemoryStream();
            await using (XlsWorkbookWriter wb = XlsWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsSheetWriter sheet = wb.AddSheet("S1");
                for (int r = 0; r < 200; r++)
                {
                    using XlsRowWriter row = sheet.StartRow();
                    row.Write("name-" + r.ToString(CultureInfo.InvariantCulture));
                    row.Write(r);
                }
                sheet.End();
                await wb.EndAsync(ct);
            }
            ms.Position = 0;
            using IExcelWorkbook workbook = Excel.Open(ms);

            List<string> actual = await SheetParallel.AggregateAsync(workbook.FirstSheet, Collecting, new SheetParallelOptions { DegreeOfParallelism = 4 }, ct);

            Assert.Equal(Sequential(workbook.FirstSheet, 0), actual);
        }

        [Theory]
        [MemberData(nameof(TableMetadataTests.Fixtures), MemberType = typeof(TableMetadataTests))]
        public async Task ATableViewReadsSequentially(string file)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using IExcelWorkbook workbook = Excel.Open(TableMetadataTests.FixturePath(file));
            Assert.True(workbook.TryGetTable("Sales", out ExcelTable? table));
            IExcelSheet view = table.AsSheet();

            List<string> actual = await SheetParallel.AggregateAsync(view, Collecting, new SheetParallelOptions { DegreeOfParallelism = 4 }, ct);

            Assert.Equal(Sequential(view, 0), actual);
        }

        [Fact]
        public async Task XlsxFamiliesMatchSequential()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using MemoryStream ms = await XlsxAsync(20_000);
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);
            ExcelParser<Person> parser = ExcelParser.FromAttributes<Person>();
            List<Person> expected = [.. parser.Parse(workbook.FirstSheet)];
            var options = new SheetParallelOptions { DegreeOfParallelism = 4, HeaderRow = 1 };

            Assert.Equal(Sequential(workbook.FirstSheet, 1), await SheetParallel.AggregateAsync(workbook.FirstSheet, Collecting, options, ct));

            var parsed = new List<Person>();
            await foreach (Person person in SheetParallel.ParseAsync(workbook.FirstSheet, parser, new SheetParallelOptions { DegreeOfParallelism = 4 }, ct))
            {
                parsed.Add(person);
            }
            Assert.Equal(expected.Select(p => (p.Name, p.Age)), parsed.Select(p => (p.Name, p.Age)));

            var seen = new System.Collections.Concurrent.ConcurrentBag<(string?, int)>();
            await SheetParallel.ForEachAsync(workbook.FirstSheet, parser, (Person p) => seen.Add((p.Name, p.Age)), options, ct);
            Assert.Equal(expected.Select(p => (p.Name, p.Age)).Order(), seen.Order());
        }

        [Fact]
        public async Task ArgumentsAreValidated()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using CsvReader reader = Excel.FromCsv(Csv(10));
            ExcelParser<Person> parser = ExcelParser.FromAttributes<Person>();

            await Assert.ThrowsAsync<ArgumentNullException>(() => SheetParallel.AggregateAsync(null!, Collecting, null, ct));
            await Assert.ThrowsAsync<ArgumentNullException>(() => SheetParallel.AggregateAsync(reader.FirstSheet, (RowAggregation<int>)null!, null, ct));
            await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => SheetParallel.AggregateAsync(reader.FirstSheet, Collecting, new SheetParallelOptions { DegreeOfParallelism = -1 }, ct));
            await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => SheetParallel.AggregateAsync(reader.FirstSheet, Collecting, new SheetParallelOptions { HeaderRow = -1 }, ct));
            Assert.Throws<ArgumentException>(() => SheetParallel.ParseAsync(reader.FirstSheet, parser, new SheetParallelOptions { HeaderRow = 2 }, ct));
            await Assert.ThrowsAsync<ArgumentException>(() => SheetParallel.ForEachAsync(reader.FirstSheet, parser, (Person _) => { }, new SheetParallelOptions { HeaderRow = 0 }, ct));
        }
    }
}
