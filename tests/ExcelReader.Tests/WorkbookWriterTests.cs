using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Enums;
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
            public DateOnly BirthDate { get; set; }
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

        private sealed class AliasedRow
        {
            [ExcelColumn("Full Name")]
            public string? Name { get; set; }
        }

        public enum Priority { Low, Medium, High }

        // Non-numeric, non-primitive properties: exercise the record writer's ToString() fallback
        // (they must land in text cells, not corrupt number cells). Both round-trip via ExcelParser.
        private sealed class FallbackRow
        {
            public Priority Priority { get; set; }
            public Guid Id { get; set; }
        }

        private static async IAsyncEnumerable<T> ToAsync<T>(IEnumerable<T> items)
        {
            foreach (T item in items)
            {
                yield return item;
                await Task.CompletedTask.ConfigureAwait(false);
            }
        }

        public enum RecordFormat { Xlsx, Xlsb, Xls }

        // Runs body against a record writer for the given format, returning the finished stream.
        private static async Task<MemoryStream> WriteRecordsAsync<T>(
            RecordFormat format, Func<Func<string, IEnumerable<T>, ValueTask>, ValueTask> body)
        {
            var ms = new MemoryStream();
            var ct = TestContext.Current.CancellationToken;
            switch (format)
            {
                case RecordFormat.Xlsx:
                    await using (var w = await RecordWriter.CreateXlsxAsync(ms, leaveOpen: true, ct: ct).ConfigureAwait(true))
                    {
                        await body((n, r) => w.WriteSheetAsync(n, r, ct)).ConfigureAwait(true);
                    }
                    break;
                case RecordFormat.Xlsb:
                    await using (var w = await RecordWriter.CreateXlsbAsync(ms, leaveOpen: true, ct: ct).ConfigureAwait(true))
                    {
                        await body((n, r) => w.WriteSheetAsync(n, r, ct)).ConfigureAwait(true);
                    }
                    break;
                default:
                    await using (var w = await RecordWriter.CreateXlsAsync(ms, leaveOpen: true, ct: ct).ConfigureAwait(true))
                    {
                        await body((n, r) => w.WriteSheetAsync(n, r, ct)).ConfigureAwait(true);
                    }
                    break;
            }
            ms.Position = 0;
            return ms;
        }

        [Theory]
        [InlineData(RecordFormat.Xlsx)]
        [InlineData(RecordFormat.Xlsb)]
        [InlineData(RecordFormat.Xls)]
        public async Task RecordWriterRoundTripsAllFormats(RecordFormat format)
        {
            var people = new[]
            {
                new PrimitivesRow { Age = 1, Score = 1.5, Balance = 10m, Active = true, BirthDate = new DateOnly(2000, 1, 2) },
                new PrimitivesRow { Age = 2, Score = 2.5, Balance = 20m, Active = false, BirthDate = new DateOnly(2001, 3, 4) },
            };

            await using var ms = await WriteRecordsAsync<PrimitivesRow>(format,
                write => write("People", people)).ConfigureAwait(true);

            await using var reader = Excel.Open(ms);
            var rows = new ExcelParser<PrimitivesRow>().Parse(reader).ToList();
            Assert.Equal(2, rows.Count);
            Assert.Equal(1, rows[0].Age);
            Assert.Equal(2.5, rows[1].Score);
            Assert.Equal(20m, rows[1].Balance);
            Assert.False(rows[1].Active);
            Assert.Equal(new DateOnly(2001, 3, 4), rows[1].BirthDate);
        }

        [Theory]
        [InlineData(RecordFormat.Xlsx)]
        [InlineData(RecordFormat.Xlsb)]
        [InlineData(RecordFormat.Xls)]
        public async Task RecordWriterRoundTripsEnumAndGuidAsText(RecordFormat format)
        {
            var id = new Guid("11112222-3333-4444-5555-666677778888");
            var rows = new[] { new FallbackRow { Priority = Priority.High, Id = id } };

            await using var ms = await WriteRecordsAsync<FallbackRow>(format,
                write => write("Data", rows)).ConfigureAwait(true);

            await using var reader = Excel.Open(ms);
            var parsed = new ExcelParser<FallbackRow>().Parse(reader).ToList();
            var row = Assert.Single(parsed);
            // If the writer had routed these through the numeric Write<U> path, the cells would be
            // corrupt numbers and neither would parse back — so a clean round-trip proves the text fallback.
            Assert.Equal(Priority.High, row.Priority);
            Assert.Equal(id, row.Id);
        }

        [Fact]
        public async Task RecordWriterWritesHeaderOnlyForEmptyRecords()
        {
            var ms = new MemoryStream();
            await using (var writer = await RecordWriter.CreateXlsxAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken).ConfigureAwait(true))
            {
                await writer.WriteSheetAsync("Sheet1", Array.Empty<StringRow>(), TestContext.Current.CancellationToken).ConfigureAwait(true);
            }
            ms.Position = 0;

            // Header row is written even with no records; the parser then yields zero data rows.
            await using var reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("Name", e.Current[0].GetString());
            Assert.False(e.MoveNext());

            await using var reader2 = Excel.From(ms);
            var parsed = new ExcelParser<StringRow>().Parse(reader2).ToList();
            Assert.Empty(parsed);
        }

        [Fact]
        public async Task RecordWriterWritesHeaderAndDataAcrossSheets()
        {
            var people = new[]
            {
                new PrimitivesRow { Age = 1, Score = 1.5, Balance = 10m, Active = true, BirthDate = new DateOnly(2000, 1, 2) },
                new PrimitivesRow { Age = 2, Score = 2.5, Balance = 20m, Active = false, BirthDate = new DateOnly(2001, 3, 4) },
            };
            var names = new[] { new StringRow { Name = "Alice" }, new StringRow { Name = "Bob" } };

            var ms = new MemoryStream();
            await using (var writer = await RecordWriter.CreateXlsxAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken).ConfigureAwait(true))
            {
                await writer.WriteSheetAsync("People", people, TestContext.Current.CancellationToken).ConfigureAwait(true);
                await writer.WriteSheetAsync("Names", names, TestContext.Current.CancellationToken).ConfigureAwait(true);
            }
            ms.Position = 0;

            await using var reader = Excel.From(ms);
            var primitives = new ExcelParser<PrimitivesRow>().Parse(reader).ToList();
            Assert.Equal(2, primitives.Count);
            Assert.Equal(1, primitives[0].Age);
            Assert.Equal(20m, primitives[1].Balance);
            Assert.False(primitives[1].Active);

            await using var reader2 = Excel.From(ms);
            reader2.MoveToSheet(1);
            var strings = new ExcelParser<StringRow>().Parse(reader2).ToList();
            Assert.Equal(["Alice", "Bob"], strings.Select(r => r.Name));
        }

        [Fact]
        public async Task RecordWriterAsyncEnumerableAndAliasHeader()
        {
            var names = new[] { new AliasedRow { Name = "Alice" }, new AliasedRow { Name = "Bob" } };

            var ms = new MemoryStream();
            await using (var writer = await RecordWriter.CreateXlsxAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken).ConfigureAwait(true))
            {
                await writer.WriteSheetAsync("Sheet1", ToAsync(names), TestContext.Current.CancellationToken).ConfigureAwait(true);
            }
            ms.Position = 0;

            await using var reader = Excel.From(ms);
            var rows = new ExcelParser<AliasedRow>().Parse(reader).ToList();
            Assert.Equal(["Alice", "Bob"], rows.Select(r => r.Name));
        }

        [Fact]
        public async Task RecordWriterHandlesNullableColumns()
        {
            var rows = new[]
            {
                new NullableRow { Quantity = 7, EventDate = new DateTime(2020, 5, 6, 0, 0, 0, DateTimeKind.Unspecified) },
                new NullableRow { Quantity = null, EventDate = null },
            };

            var ms = new MemoryStream();
            await using (var writer = await RecordWriter.CreateXlsxAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken).ConfigureAwait(true))
            {
                await writer.WriteSheetAsync("Sheet1", rows, TestContext.Current.CancellationToken).ConfigureAwait(true);
            }
            ms.Position = 0;

            await using var reader = Excel.Open(ms);
            var parsed = new ExcelParser<NullableRow>().Parse(reader).ToList();
            Assert.Equal(2, parsed.Count);
            Assert.Equal(7, parsed[0].Quantity);
            Assert.Equal(new DateTime(2020, 5, 6, 0, 0, 0, DateTimeKind.Unspecified), parsed[0].EventDate);
            Assert.Null(parsed[1].Quantity);
            Assert.Null(parsed[1].EventDate);
        }

        [Fact]
        public async Task RecordWriterRejectsDuplicateSheetName()
        {
            var ms = new MemoryStream();
            await using var writer = await RecordWriter.CreateXlsxAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken).ConfigureAwait(true);
            await writer.WriteSheetAsync("Sheet1", new[] { new StringRow { Name = "Alice" } }, TestContext.Current.CancellationToken).ConfigureAwait(true);
            await Assert.ThrowsAsync<InvalidOperationException>(async () =>
                await writer.WriteSheetAsync("Sheet1", new[] { new StringRow { Name = "Bob" } }, TestContext.Current.CancellationToken).ConfigureAwait(true)).ConfigureAwait(true);
        }

        [Fact]
        public async Task AllPrimitiveTypesRoundTrip()
        {
            var birth = new DateOnly(1990, 6, 15);

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
            Assert.Equal(birth, rows[0].BirthDate);
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

        [Fact]
        public async Task NullableAndGenericNumericOverloadsRoundTrip()
        {
            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                SheetWriter sheet = wb.AddSheet("Numbers");
                await sheet.StartAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

                await using (RowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    row.Write((long?)1234567890123L);
                    row.Write((long?)null);
                    row.Write((double?)2.5d);
                    row.Write((double?)null);
                    row.Write((decimal?)3.75m);
                    row.Write((decimal?)null);
                    row.Write<short>(4);
                    row.Write((short?)5);
                    row.Write<short>(null);
                    row.Skip(0);
                    row.Write(6);
                }

                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("1234567890123", e.Current[0].GetString());
            Assert.Equal(CellType.Empty, e.Current[1].Type);
            Assert.Equal("2.5", e.Current[2].GetString());
            Assert.Equal(CellType.Empty, e.Current[3].Type);
            Assert.Equal("3.75", e.Current[4].GetString());
            Assert.Equal(CellType.Empty, e.Current[5].Type);
            Assert.Equal("4", e.Current[6].GetString());
            Assert.Equal("5", e.Current[7].GetString());
            Assert.Equal(CellType.Empty, e.Current[8].Type);
            Assert.Equal("6", e.Current[9].GetString());
            Assert.False(e.MoveNext());
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
            Exception? ex = await Record.ExceptionAsync(async () =>
                await wb.DisposeAsync().ConfigureAwait(true)).ConfigureAwait(true);

            Assert.Null(ex);
        }
    }
}
