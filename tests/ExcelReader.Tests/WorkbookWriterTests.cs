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
            await using WorkbookWriter wb = await WorkbookWriter.CreateAsync(ms, leaveOpen: true);
            await wb.StartAsync();
            await body(wb);
            await wb.EndAsync();
            ms.Position = 0;
            return ms;
        }

        // --- Basic type round-trips ---

        [Fact]
        public async Task StringCellRoundTrip()
        {
            using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync();

                await using (RowWriter header = await sheet.StartRowAsync())
                {
                    await header.WriteAsync("Name");
                }

                await using (RowWriter row = await sheet.StartRowAsync())
                {
                    await row.WriteAsync("Alice");
                }

                await sheet.EndAsync();
            });

            using var reader = Excel.From(ms);
            var rows = new ExcelParser<StringRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Equal("Alice", rows[0].Name);
        }

        [Fact]
        public async Task AllPrimitiveTypesRoundTrip()
        {
            var birth = new DateTime(1990, 6, 15, 0, 0, 0, DateTimeKind.Unspecified);

            using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync();

                await using (RowWriter header = await sheet.StartRowAsync())
                {
                    await header.WriteAsync("Age");
                    await header.WriteAsync("Score");
                    await header.WriteAsync("Balance");
                    await header.WriteAsync("Active");
                    await header.WriteAsync("BirthDate");
                }

                await using (RowWriter row = await sheet.StartRowAsync())
                {
                    await row.WriteAsync(42);
                    await row.WriteAsync(3.14);
                    await row.WriteAsync(9999.99m);
                    await row.WriteAsync(true);
                    await row.WriteAsync(birth);
                }

                await sheet.EndAsync();
            });

            using var reader = Excel.From(ms);
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
            using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync();

                await using (RowWriter header = await sheet.StartRowAsync())
                {
                    await header.WriteAsync("Active");
                }

                await using (RowWriter row = await sheet.StartRowAsync())
                {
                    await row.WriteAsync(false);
                }

                await sheet.EndAsync();
            });

            using var reader = Excel.From(ms);
            var rows = new ExcelParser<PrimitivesRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.False(rows[0].Active);
        }

        [Fact]
        public async Task NullStringWritesEmptyCell()
        {
            using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync();

                await using (RowWriter header = await sheet.StartRowAsync())
                {
                    await header.WriteAsync("Name");
                }

                await using (RowWriter row = await sheet.StartRowAsync())
                {
                    await row.WriteAsync((string?)null);
                }

                await sheet.EndAsync();
            });

            using var reader = Excel.From(ms);
            var rows = new ExcelParser<StringRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Null(rows[0].Name);
        }

        [Fact]
        public async Task NullableIntFilledRoundTrip()
        {
            using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync();

                await using (RowWriter header = await sheet.StartRowAsync())
                {
                    await header.WriteAsync("Quantity");
                    await header.WriteAsync("EventDate");
                }

                await using (RowWriter row = await sheet.StartRowAsync())
                {
                    await row.WriteAsync((int?)77);
                    await row.WriteAsync((DateTime?)null);
                }

                await sheet.EndAsync();
            });

            using var reader = Excel.From(ms);
            var rows = new ExcelParser<NullableRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Equal(77, rows[0].Quantity);
            Assert.Null(rows[0].EventDate);
        }

        [Fact]
        public async Task NullableDateTimeRoundTrip()
        {
            var dt = new DateTime(2024, 3, 20, 0, 0, 0, DateTimeKind.Unspecified);

            using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync();

                await using (RowWriter header = await sheet.StartRowAsync())
                {
                    await header.WriteAsync("Quantity");
                    await header.WriteAsync("EventDate");
                }

                await using (RowWriter row = await sheet.StartRowAsync())
                {
                    await row.WriteAsync((int?)null);
                    await row.WriteAsync((DateTime?)dt);
                }

                await sheet.EndAsync();
            });

            using var reader = Excel.From(ms);
            var rows = new ExcelParser<NullableRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Null(rows[0].Quantity);
            Assert.NotNull(rows[0].EventDate);
            Assert.Equal(dt.Date, rows[0].EventDate!.Value.Date);
        }

        // --- Multiple rows ---

        [Fact]
        public async Task MultipleRowsRoundTrip()
        {
            using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync();

                await using (RowWriter header = await sheet.StartRowAsync())
                {
                    await header.WriteAsync("Name");
                }

                foreach (string name in new[] { "Alice", "Bob", "Carol", "Dave", "Eve" })
                {
                    await using RowWriter row = await sheet.StartRowAsync();
                    await row.WriteAsync(name);
                }

                await sheet.EndAsync();
            });

            using var reader = Excel.From(ms);
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
            using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet1 = wb.AddSheet("Alpha");
                await sheet1.StartAsync();
                await using (RowWriter h1 = await sheet1.StartRowAsync())
                {
                    await h1.WriteAsync("Value");
                }
                await using (RowWriter r1 = await sheet1.StartRowAsync())
                {
                    await r1.WriteAsync("FromAlpha");
                }
                await sheet1.EndAsync();

                SheetWriter sheet2 = wb.AddSheet("Beta");
                await sheet2.StartAsync();
                await using (RowWriter h2 = await sheet2.StartRowAsync())
                {
                    await h2.WriteAsync("Value");
                }
                await using (RowWriter r2 = await sheet2.StartRowAsync())
                {
                    await r2.WriteAsync("FromBeta");
                }
                await sheet2.EndAsync();
            });

            using var reader = Excel.From(ms);
            var parser = new ExcelParser<MultiSheetRow>();
            var rowsSheet1 = parser.Parse(reader).ToList();
            Assert.Single(rowsSheet1);
            Assert.Equal("FromAlpha", rowsSheet1[0].Value);

            using var reader2 = Excel.From(ms);
            reader2.MoveToSheet(1);
            var rowsSheet2 = parser.Parse(reader2).ToList();
            Assert.Single(rowsSheet2);
            Assert.Equal("FromBeta", rowsSheet2[0].Value);
        }

        // --- Skip (column gaps) ---

        [Fact]
        public async Task SkipCreatesColumnGap()
        {
            using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync();

                await using (RowWriter header = await sheet.StartRowAsync())
                {
                    await header.WriteAsync("A");
                    header.Skip(1);
                    await header.WriteAsync("C");
                }

                await using (RowWriter row = await sheet.StartRowAsync())
                {
                    await row.WriteAsync("aaa");
                    row.Skip(1);
                    await row.WriteAsync("ccc");
                }

                await sheet.EndAsync();
            });

            using var reader = Excel.From(ms);
            var rows = new ExcelParser<SparseRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Equal("aaa", rows[0].A);
            Assert.Equal("ccc", rows[0].C);
        }

        // --- XML special characters in strings ---

        [Fact]
        public async Task XmlSpecialCharsAreEscaped()
        {
            using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync();

                await using (RowWriter header = await sheet.StartRowAsync())
                {
                    await header.WriteAsync("Name");
                }

                await using (RowWriter row = await sheet.StartRowAsync())
                {
                    await row.WriteAsync("<Alice & \"Bob\">");
                }

                await sheet.EndAsync();
            });

            using var reader = Excel.From(ms);
            var rows = new ExcelParser<StringRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Equal("<Alice & \"Bob\">", rows[0].Name);
        }

        // --- DisposeAsync auto-closes ---

        [Fact]
        public async Task DisposeAsyncWithoutEndAsyncProducesReadableWorkbook()
        {
            var ms = new MemoryStream();
            WorkbookWriter wb = await WorkbookWriter.CreateAsync(ms, leaveOpen: true);
            await wb.StartAsync();

            SheetWriter sheet = wb.AddSheet("Sheet1");
            await sheet.StartAsync();

            await using (RowWriter header = await sheet.StartRowAsync())
            {
                await header.WriteAsync("Name");
            }

            await using (RowWriter row = await sheet.StartRowAsync())
            {
                await row.WriteAsync("Auto");
            }

            // Do NOT call EndAsync — rely on DisposeAsync chain
            await wb.DisposeAsync();
            ms.Position = 0;

            using var reader = Excel.From(ms);
            var rows = new ExcelParser<StringRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Equal("Auto", rows[0].Name);
        }

        // --- HeaderRow config ---

        [Fact]
        public async Task HeaderRowTwoConfig()
        {
            using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync();

                await using (RowWriter skip = await sheet.StartRowAsync())
                {
                    await skip.WriteAsync("ignored");
                }

                await using (RowWriter header = await sheet.StartRowAsync())
                {
                    await header.WriteAsync("Name");
                }

                await using (RowWriter row = await sheet.StartRowAsync())
                {
                    await row.WriteAsync("HeaderRow2");
                }

                await sheet.EndAsync();
            });

            using var reader = Excel.From(ms);
            var config = new ExcelParserConfig { HeaderRow = 2 };
            var rows = new ExcelParser<StringRow>(config).Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Equal("HeaderRow2", rows[0].Name);
        }

        // --- State machine violations ---

        [Fact]
        public async Task AddSheetBeforeStartAsyncThrows()
        {
            await using WorkbookWriter wb = await WorkbookWriter.CreateAsync(new MemoryStream());
            Assert.Throws<InvalidOperationException>(() => wb.AddSheet("Sheet1"));
        }

        [Fact]
        public async Task AddSheetWhileSheetActiveThrows()
        {
            await using WorkbookWriter wb = await WorkbookWriter.CreateAsync(new MemoryStream());
            await wb.StartAsync();

            SheetWriter sheet = wb.AddSheet("Sheet1");
            await sheet.StartAsync();

            Assert.Throws<InvalidOperationException>(() => wb.AddSheet("Sheet2"));
            await sheet.EndAsync();
        }

        [Fact]
        public async Task StartRowBeforeSheetStartAsyncThrows()
        {
            await using WorkbookWriter wb = await WorkbookWriter.CreateAsync(new MemoryStream());
            await wb.StartAsync();

            SheetWriter sheet = wb.AddSheet("Sheet1");

            await Assert.ThrowsAsync<InvalidOperationException>(() => sheet.StartRowAsync().AsTask());
        }

        [Fact]
        public async Task StartRowWhileRowActiveThrows()
        {
            await using WorkbookWriter wb = await WorkbookWriter.CreateAsync(new MemoryStream());
            await wb.StartAsync();

            SheetWriter sheet = wb.AddSheet("Sheet1");
            await sheet.StartAsync();

            RowWriter row = await sheet.StartRowAsync();

            await Assert.ThrowsAsync<InvalidOperationException>(() => sheet.StartRowAsync().AsTask());

            await row.DisposeAsync();
            await sheet.EndAsync();
        }

        [Fact]
        public async Task WriteAfterRowDisposedThrows()
        {
            await using WorkbookWriter wb = await WorkbookWriter.CreateAsync(new MemoryStream());
            await wb.StartAsync();

            SheetWriter sheet = wb.AddSheet("Sheet1");
            await sheet.StartAsync();

            RowWriter row = await sheet.StartRowAsync();
            await row.DisposeAsync();

            await Assert.ThrowsAsync<ObjectDisposedException>(() => row.WriteAsync("late").AsTask());

            await sheet.EndAsync();
        }

        // --- Large workbook ---

        [Fact]
        public async Task LargeWorkbookRoundTrip()
        {
            const int rowCount = 1000;

            using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync();

                await using (RowWriter header = await sheet.StartRowAsync())
                {
                    await header.WriteAsync("Age");
                }

                for (int i = 0; i < rowCount; i++)
                {
                    await using RowWriter row = await sheet.StartRowAsync();
                    await row.WriteAsync(i);
                }

                await sheet.EndAsync();
            });

            using var reader = Excel.From(ms);
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
            var ms = new MemoryStream();
            await using WorkbookWriter wb = await WorkbookWriter.CreateAsync(ms, leaveOpen: true);
            await wb.StartAsync();
            await wb.FlushAsync();
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
            var ms = new MemoryStream();
            WorkbookWriter wb = await WorkbookWriter.CreateAsync(ms, leaveOpen: true);
            await wb.StartAsync();

            SheetWriter sheet = wb.AddSheet("Sheet1");
            await sheet.StartAsync();
            await sheet.EndAsync();

            await wb.DisposeAsync();
            await wb.DisposeAsync();
        }
    }
}
