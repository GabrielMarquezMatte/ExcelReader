using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    public class WorkbookWriterTests
    {
        // --- Helpers ---

        private sealed class StringRow
        {
            public string? Name { get; set; }
        }

        private sealed class PrimitivesRow
        {
            public int Age { get; set; }
            public double Score { get; set; }
            public decimal Balance { get; set; }
            public bool Active { get; set; }
            public DateTime BirthDate { get; set; }
        }

        private sealed class NullableRow
        {
            public int? Quantity { get; set; }
            public DateTime? EventDate { get; set; }
        }

        private sealed class SparseRow
        {
            public string? A { get; set; }
            public string? C { get; set; }
        }

        private sealed class MultiSheetRow
        {
            public string? Value { get; set; }
        }

        private static async Task<MemoryStream> WriteWorkbookAsync(
            Func<WorkbookWriter, Task> body)
        {
            var ms = new MemoryStream();
            await using var wb = await WorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken).ConfigureAwait(true);
            await wb.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            await body(wb).ConfigureAwait(true);
            await wb.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            ms.Position = 0;
            return ms;
        }

        // --- Basic type round-trips ---

        [Fact]
        public async Task StringCellRoundTrip()
        {
            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

                await using (var header = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    header.Write("Name");
                }

                await using (var row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    row.Write("Alice");
                }

                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.From(ms);
            var rows = new ExcelParser<StringRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Equal("Alice", rows[0].Name);
        }

        [Fact]
        public async Task AllPrimitiveTypesRoundTrip()
        {
            var birth = new DateTime(1990, 6, 15, 0, 0, 0, DateTimeKind.Unspecified);

            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

                await using (RowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    header.Write("Age");
                    header.Write("Score");
                    header.Write("Balance");
                    header.Write("Active");
                    header.Write("BirthDate");
                }

                await using (RowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    row.Write(42);
                    row.Write(3.14);
                    row.Write(9999.99m);
                    row.Write(true);
                    row.Write(birth);
                }

                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.From(ms);
            var rows = new ExcelParser<PrimitivesRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Equal(42, rows[0].Age);
            Assert.Equal(3.14, rows[0].Score, precision: 10);
            Assert.Equal(9999.99m, rows[0].Balance);
            Assert.True(rows[0].Active);
            Assert.Equal(birth.Date, rows[0].BirthDate.Date);
        }

        [Fact]
        public async Task BoolFalseRoundTrip()
        {
            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

                await using (RowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    header.Write("Active");
                }

                await using (RowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    row.Write(false);
                }

                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.From(ms);
            var rows = new ExcelParser<PrimitivesRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.False(rows[0].Active);
        }

        [Fact]
        public async Task NullStringWritesEmptyCell()
        {
            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

                await using (RowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    header.Write("Name");
                }

                await using (RowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    row.Write((string?)null);
                }

                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.From(ms);
            var rows = new ExcelParser<StringRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Null(rows[0].Name);
        }

        [Fact]
        public async Task NullableIntFilledRoundTrip()
        {
            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

                await using (RowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    header.Write("Quantity");
                    header.Write("EventDate");
                }

                await using (RowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    row.Write((int?)77);
                    row.Write((DateTime?)null);
                }

                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.From(ms);
            var rows = new ExcelParser<NullableRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Equal(77, rows[0].Quantity);
            Assert.Null(rows[0].EventDate);
        }

        [Fact]
        public async Task NullableDateTimeRoundTrip()
        {
            var dt = new DateTime(2024, 3, 20, 0, 0, 0, DateTimeKind.Unspecified);

            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

                await using (RowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    header.Write("Quantity");
                    header.Write("EventDate");
                }

                await using (RowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    row.Write((int?)null);
                    row.Write((DateTime?)dt);
                }

                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.From(ms);
            var rows = new ExcelParser<NullableRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Null(rows[0].Quantity);
            Assert.NotNull(rows[0].EventDate);
            Assert.Equal(dt.Date, rows[0].EventDate!.Value.Date);
        }

        private static readonly string[] stringArray = ["Alice", "Bob", "Carol", "Dave", "Eve"];

        // --- Multiple rows ---

        [Fact]
        public async Task MultipleRowsRoundTrip()
        {
            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

                await using (RowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    header.Write("Name");
                }

                foreach (string name in stringArray)
                {
                    await using RowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
                    row.Write(name);
                }

                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.From(ms);
            var rows = new ExcelParser<StringRow>().Parse(reader).ToList();
            Assert.Equal(5, rows.Count);
            Assert.Equal("Alice", rows[0].Name);
            Assert.Equal("Bob", rows[1].Name);
            Assert.Equal("Carol", rows[2].Name);
            Assert.Equal("Dave", rows[3].Name);
            Assert.Equal("Eve", rows[4].Name);
        }

        // --- Multiple sheets ---

        [Fact]
        public async Task MultipleSheetsRoundTrip()
        {
            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet1 = wb.AddSheet("Alpha");
                await sheet1.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
                await using (RowWriter h1 = await sheet1.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    h1.Write("Value");
                }
                await using (RowWriter r1 = await sheet1.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    r1.Write("FromAlpha");
                }
                await sheet1.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

                SheetWriter sheet2 = wb.AddSheet("Beta");
                await sheet2.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
                await using (RowWriter h2 = await sheet2.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    h2.Write("Value");
                }
                await using (RowWriter r2 = await sheet2.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    r2.Write("FromBeta");
                }
                await sheet2.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.From(ms);
            var parser = new ExcelParser<MultiSheetRow>();
            var rowsSheet1 = parser.Parse(reader).ToList();
            Assert.Single(rowsSheet1);
            Assert.Equal("FromAlpha", rowsSheet1[0].Value);

            await using var reader2 = Excel.From(ms);
            reader2.MoveToSheet(1);
            var rowsSheet2 = parser.Parse(reader2).ToList();
            Assert.Single(rowsSheet2);
            Assert.Equal("FromBeta", rowsSheet2[0].Value);
        }

        // --- Skip (column gaps) ---

        [Fact]
        public async Task SkipCreatesColumnGap()
        {
            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

                await using (RowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    header.Write("A");
                    header.Skip(1);
                    header.Write("C");
                }

                await using (RowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    row.Write("aaa");
                    row.Skip(1);
                    row.Write("ccc");
                }

                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.From(ms);
            var rows = new ExcelParser<SparseRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Equal("aaa", rows[0].A);
            Assert.Equal("ccc", rows[0].C);
        }

        // --- XML special characters in strings ---

        [Fact]
        public async Task XmlSpecialCharsAreEscaped()
        {
            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

                await using (RowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    header.Write("Name");
                }

                await using (RowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    row.Write("<Alice & \"Bob\">");
                }

                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.From(ms);
            var rows = new ExcelParser<StringRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Equal("<Alice & \"Bob\">", rows[0].Name);
        }

        // --- DisposeAsync auto-closes ---

        [Fact]
        public async Task DisposeAsyncWithoutEndAsyncProducesReadableWorkbook()
        {
            await using var ms = new MemoryStream();
            WorkbookWriter wb = await WorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken).ConfigureAwait(true);
            await wb.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

            SheetWriter sheet = wb.AddSheet("Sheet1");
            await sheet.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

            await using (RowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
            {
                header.Write("Name");
            }

            await using (RowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
            {
                row.Write("Auto");
            }

            // Do NOT call EndAsync — rely on DisposeAsync chain
            await wb.DisposeAsync().ConfigureAwait(true);
            ms.Position = 0;

            await using var reader = Excel.From(ms);
            var rows = new ExcelParser<StringRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Equal("Auto", rows[0].Name);
        }

        // --- HeaderRow config ---

        [Fact]
        public async Task HeaderRowTwoConfig()
        {
            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

                await using (RowWriter skip = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    skip.Write("ignored");
                }

                await using (RowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    header.Write("Name");
                }

                await using (RowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    row.Write("HeaderRow2");
                }

                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.From(ms);
            var config = new ExcelParserConfig { HeaderRow = 2 };
            var rows = new ExcelParser<StringRow>(config).Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Equal("HeaderRow2", rows[0].Name);
        }

        // --- State machine violations ---

        [Fact]
        public async Task AddSheetBeforeStartAsyncThrows()
        {
            await using WorkbookWriter wb = await WorkbookWriter.CreateAsync(new MemoryStream(), ct: TestContext.Current.CancellationToken).ConfigureAwait(true);
            Assert.Throws<InvalidOperationException>(() => wb.AddSheet("Sheet1"));
        }

        [Fact]
        public async Task AddSheetWhileSheetActiveThrows()
        {
            await using WorkbookWriter wb = await WorkbookWriter.CreateAsync(new MemoryStream(), ct: TestContext.Current.CancellationToken).ConfigureAwait(true);
            await wb.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

            SheetWriter sheet = wb.AddSheet("Sheet1");
            await sheet.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

            Assert.Throws<InvalidOperationException>(() => wb.AddSheet("Sheet2"));
            await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
        }

        [Fact]
        public async Task StartRowBeforeSheetStartAsyncThrows()
        {
            await using WorkbookWriter wb = await WorkbookWriter.CreateAsync(new MemoryStream(), ct: TestContext.Current.CancellationToken).ConfigureAwait(true);
            await wb.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

            SheetWriter sheet = wb.AddSheet("Sheet1");

            await Assert.ThrowsAsync<InvalidOperationException>(() => sheet.StartRowAsync(TestContext.Current.CancellationToken).AsTask()).ConfigureAwait(true);
        }

        [Fact]
        public async Task StartRowWhileRowActiveThrows()
        {
            await using WorkbookWriter wb = await WorkbookWriter.CreateAsync(new MemoryStream(), ct: TestContext.Current.CancellationToken).ConfigureAwait(true);
            await wb.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

            SheetWriter sheet = wb.AddSheet("Sheet1");
            await sheet.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

            RowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

            await Assert.ThrowsAsync<InvalidOperationException>(() => sheet.StartRowAsync(TestContext.Current.CancellationToken).AsTask()).ConfigureAwait(true);

            await row.DisposeAsync().ConfigureAwait(true);
            await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
        }

        [Fact]
        public async Task WriteAfterRowDisposedThrows()
        {
            await using WorkbookWriter wb = await WorkbookWriter.CreateAsync(new MemoryStream(), ct: TestContext.Current.CancellationToken).ConfigureAwait(true);
            await wb.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

            SheetWriter sheet = wb.AddSheet("Sheet1");
            await sheet.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

            RowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            await row.DisposeAsync().ConfigureAwait(true);

            Assert.Throws<ObjectDisposedException>(() => row.Write("late"));

            await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
        }

        // --- Large workbook ---

        [Fact]
        public async Task LargeWorkbookRoundTrip()
        {
            const int rowCount = 1000;

            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

                await using (RowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    header.Write("Age");
                }

                for (int i = 0; i < rowCount; i++)
                {
                    await using RowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
                    row.Write(i);
                }

                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.From(ms);
            var rows = new ExcelParser<PrimitivesRow>().Parse(reader).ToList();
            Assert.Equal(rowCount, rows.Count);
            for (int i = 0; i < rowCount; i++)
            {
                Assert.Equal(i, rows[i].Age);
            }
        }

        // --- FlushAsync ---

        [Fact]
        public async Task FlushAsyncDoesNotThrow()
        {
            await using var ms = new MemoryStream();
            await using WorkbookWriter wb = await WorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken).ConfigureAwait(true);
            await wb.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            await wb.FlushAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            Assert.True(ms.Length >= 0);
        }

        // --- DisposeAsync is idempotent ---

        [Fact]
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Intentionally testing manual double-dispose for idempotency contract.")]
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP016:Don't use disposed instance",
            Justification = "Second DisposeAsync call is the subject under test for idempotency.")]
        public async Task DisposeAsyncIsIdempotent()
        {
            await using var ms = new MemoryStream();
            WorkbookWriter wb = await WorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken).ConfigureAwait(true);
            await wb.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

            SheetWriter sheet = wb.AddSheet("Sheet1");
            await sheet.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

            await wb.DisposeAsync().ConfigureAwait(true);
            await wb.DisposeAsync().ConfigureAwait(true);
        }
    }
}
