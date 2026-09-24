using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    public class WorkbookWriterTests
    {

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
            Func<XlsxWorkbookWriter, Task> body)
        {
            var ms = new MemoryStream();
            await using var wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true);
            await body(wb).ConfigureAwait(true);
            await wb.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            ms.Position = 0;
            return ms;
        }


        [Fact]
        public async Task StringCellRoundTrip()
        {
            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                XlsxSheetWriter sheet = wb.AddSheet("Sheet1");

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

            await using var reader = Excel.FromXlsx(ms);
            var rows = ExcelParser.FromAttributes<StringRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Equal("Alice", rows[0].Name);
        }

        private sealed class AliasedRow
        {
            [ExcelColumn("Full Name")]
            public string? Name { get; set; }
        }

        public enum Priority { Low, Medium, High }

        private sealed class FallbackRow
        {
            public Priority Priority { get; set; }
            public Guid Id { get; set; }
        }

        private sealed class TimeRow
        {
            public TimeOnly Open { get; set; }
            public TimeOnly? Close { get; set; }
        }

        private sealed class IgnoreRow
        {
            public string? Name { get; set; }

            [ExcelIgnore]
            public int Computed { get; set; }
        }

        private readonly record struct Money(decimal Amount);

        private sealed class MoneyConverter : IExcelCellConverter<Money>, IExcelCellWriter<Money>
        {
            public bool TryConvert(in Cell cell, bool isDate1904, IFormatProvider provider, out Money value)
            {
                if (cell.TryParse(provider, out decimal amount))
                {
                    value = new Money(amount);
                    return true;
                }
                value = default;
                return false;
            }

            public void Write(IRowWriter row, Money value)
            {
                row.Write(value.Amount);
            }
        }

        private sealed class MoneyRow
        {
            [ExcelConverter(typeof(MoneyConverter))]
            public Money Price { get; set; }
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

        private static async Task<MemoryStream> WriteRecordsAsync<T>(
            RecordFormat format, Func<Func<string, IEnumerable<T>, ValueTask>, ValueTask> body)
        {
            var ms = new MemoryStream();
            var ct = TestContext.Current.CancellationToken;
            switch (format)
            {
                case RecordFormat.Xlsx:
                    await using (var w = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
                    {
                        await body(async (n, r) =>
                        {
                            await using var sheet = w.AddSheet(n);
                            await sheet.WriteRecordsAsync(r, ExcelRecordLayout.FromAttributes<T>(), ct).ConfigureAwait(true);
                        }).ConfigureAwait(true);
                    }
                    break;
                case RecordFormat.Xlsb:
                    await using (var w = XlsbWorkbookWriter.Create(ms, leaveOpen: true))
                    {
                        await body(async (n, r) =>
                        {
                            await using var sheet = w.AddSheet(n);
                            await sheet.WriteRecordsAsync(r, ExcelRecordLayout.FromAttributes<T>(), ct).ConfigureAwait(true);
                        }).ConfigureAwait(true);
                    }
                    break;
                default:
                    await using (var w = XlsWorkbookWriter.Create(ms, leaveOpen: true))
                    {
                        await body(async (n, r) =>
                        {
                            await using var sheet = w.AddSheet(n);
                            await sheet.WriteRecordsAsync(r, ExcelRecordLayout.FromAttributes<T>(), ct).ConfigureAwait(true);
                        }).ConfigureAwait(true);
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
            var rows = ExcelParser.FromAttributes<PrimitivesRow>().Parse(reader).ToList();
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
            var parsed = ExcelParser.FromAttributes<FallbackRow>().Parse(reader).ToList();
            var row = Assert.Single(parsed);
            Assert.Equal(Priority.High, row.Priority);
            Assert.Equal(id, row.Id);
        }

        [Fact]
        public async Task RecordWriterWritesHeaderOnlyForEmptyRecords()
        {
            var ms = new MemoryStream();
            await using (var writer = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
            {
                await using (var sheet = writer.AddSheet("Sheet1"))
                {
                    await sheet.WriteRecordsAsync(Array.Empty<StringRow>(), ExcelRecordLayout.FromAttributes<StringRow>(), TestContext.Current.CancellationToken).ConfigureAwait(true);
                }
            }
            ms.Position = 0;

            await using var reader = Excel.FromXlsx(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("Name", e.Current[0].GetString());
            Assert.False(e.MoveNext());

            await using var reader2 = Excel.FromXlsx(ms);
            var parsed = ExcelParser.FromAttributes<StringRow>().Parse(reader2).ToList();
            Assert.Empty(parsed);
        }

        [Theory]
        [InlineData(RecordFormat.Xlsx)]
        [InlineData(RecordFormat.Xlsb)]
        [InlineData(RecordFormat.Xls)]
        public async Task RecordWriterRoundTripsTimeOnly(RecordFormat format)
        {
            var rows = new[]
            {
                new TimeRow { Open = new TimeOnly(9, 30, 0), Close = new TimeOnly(17, 45, 30) },
                new TimeRow { Open = new TimeOnly(0, 0, 0), Close = null },
            };

            await using var ms = await WriteRecordsAsync<TimeRow>(format,
                write => write("Times", rows)).ConfigureAwait(true);

            await using var reader = Excel.Open(ms);
            var parsed = ExcelParser.FromAttributes<TimeRow>().Parse(reader).ToList();
            Assert.Equal(2, parsed.Count);
            Assert.Equal(new TimeOnly(9, 30, 0), parsed[0].Open);
            Assert.Equal(new TimeOnly(17, 45, 30), parsed[0].Close);
            Assert.Equal(new TimeOnly(0, 0, 0), parsed[1].Open);
            Assert.Null(parsed[1].Close);
        }

        [Fact]
        public async Task RecordWriterSkipsIgnoredProperty()
        {
            var rows = new[] { new IgnoreRow { Name = "Alice", Computed = 999 } };

            var ms = new MemoryStream();
            await using (var writer = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
            {
                await using (var sheet = writer.AddSheet("Sheet1"))
                {
                    await sheet.WriteRecordsAsync(rows, ExcelRecordLayout.FromAttributes<IgnoreRow>(), TestContext.Current.CancellationToken).ConfigureAwait(true);
                }
            }
            ms.Position = 0;

            await using var reader = Excel.FromXlsx(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("Name", e.Current[0].GetString());
            Assert.Equal(CellType.Empty, e.Current[1].Type);

            await using var reader2 = Excel.FromXlsx(ms);
            var parsed = ExcelParser.FromAttributes<IgnoreRow>().Parse(reader2).ToList();
            var row = Assert.Single(parsed);
            Assert.Equal("Alice", row.Name);
            Assert.Equal(0, row.Computed);
        }

        [Theory]
        [InlineData(RecordFormat.Xlsx)]
        [InlineData(RecordFormat.Xlsb)]
        [InlineData(RecordFormat.Xls)]
        public async Task RecordWriterUsesExcelConverterOnWrite(RecordFormat format)
        {
            var rows = new[] { new MoneyRow { Price = new Money(19.5m) }, new MoneyRow { Price = new Money(1234.25m) } };

            await using var ms = await WriteRecordsAsync<MoneyRow>(format,
                write => write("Prices", rows)).ConfigureAwait(true);

            await using var reader = Excel.Open(ms);
            var parsed = ExcelParser.FromAttributes<MoneyRow>().Parse(reader).ToList();
            Assert.Equal(2, parsed.Count);
            Assert.Equal(new Money(19.5m), parsed[0].Price);
            Assert.Equal(new Money(1234.25m), parsed[1].Price);
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
            await using (var writer = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
            {
                await using (var sheet = writer.AddSheet("People"))
                {
                    await sheet.WriteRecordsAsync(people, ExcelRecordLayout.FromAttributes<PrimitivesRow>(), TestContext.Current.CancellationToken).ConfigureAwait(true);
                }
                await using (var sheet = writer.AddSheet("Names"))
                {
                    await sheet.WriteRecordsAsync(names, ExcelRecordLayout.FromAttributes<StringRow>(), TestContext.Current.CancellationToken).ConfigureAwait(true);
                }
            }
            ms.Position = 0;

            await using var reader = Excel.FromXlsx(ms);
            var primitives = ExcelParser.FromAttributes<PrimitivesRow>().Parse(reader).ToList();
            Assert.Equal(2, primitives.Count);
            Assert.Equal(1, primitives[0].Age);
            Assert.Equal(20m, primitives[1].Balance);
            Assert.False(primitives[1].Active);

            await using var reader2 = Excel.FromXlsx(ms);
            reader2.MoveToSheet(1);
            var strings = ExcelParser.FromAttributes<StringRow>().Parse(reader2).ToList();
            Assert.Equal(["Alice", "Bob"], strings.Select(r => r.Name), StringComparer.Ordinal);
        }

        [Fact]
        public async Task RecordWriterAsyncEnumerableAndAliasHeader()
        {
            var names = new[] { new AliasedRow { Name = "Alice" }, new AliasedRow { Name = "Bob" } };

            var ms = new MemoryStream();
            await using (var writer = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
            {
                await using (var sheet = writer.AddSheet("Sheet1"))
                {
                    await sheet.WriteRecordsAsync(ToAsync(names), ExcelRecordLayout.FromAttributes<AliasedRow>(), TestContext.Current.CancellationToken).ConfigureAwait(true);
                }
            }
            ms.Position = 0;

            await using var reader = Excel.FromXlsx(ms);
            var rows = ExcelParser.FromAttributes<AliasedRow>().Parse(reader).ToList();
            Assert.Equal(["Alice", "Bob"], rows.Select(r => r.Name), StringComparer.Ordinal);
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
            await using (var writer = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
            {
                await using (var sheet = writer.AddSheet("Sheet1"))
                {
                    await sheet.WriteRecordsAsync(rows, ExcelRecordLayout.FromAttributes<NullableRow>(), TestContext.Current.CancellationToken).ConfigureAwait(true);
                }
            }
            ms.Position = 0;

            await using var reader = Excel.Open(ms);
            var parsed = ExcelParser.FromAttributes<NullableRow>().Parse(reader).ToList();
            Assert.Equal(2, parsed.Count);
            Assert.Equal(7, parsed[0].Quantity);
            Assert.Equal(new DateTime(2020, 5, 6, 0, 0, 0, DateTimeKind.Unspecified), parsed[0].EventDate);
            Assert.Null(parsed[1].Quantity);
            Assert.Null(parsed[1].EventDate);
        }

        [Fact]
        public async Task WorkbookRejectsDuplicateSheetNameIgnoringCase()
        {
            var ms = new MemoryStream();
            await using var writer = XlsxWorkbookWriter.Create(ms, leaveOpen: true);
            await using (var sheet = writer.AddSheet("Sheet1"))
            {
                await sheet.WriteRecordsAsync(new[] { new StringRow { Name = "Alice" } }, ExcelRecordLayout.FromAttributes<StringRow>(), TestContext.Current.CancellationToken).ConfigureAwait(true);
            }
            await Assert.ThrowsAsync<InvalidOperationException>(async () =>
            {
                await using var duplicate = writer.AddSheet("SHEET1");
            }).ConfigureAwait(true);
        }

        [Fact]
        public async Task AllPrimitiveTypesRoundTrip()
        {
            var birth = new DateOnly(1990, 6, 15);

            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                XlsxSheetWriter sheet = wb.AddSheet("Sheet1");

                await using (XlsxRowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    header.Write("Age");
                    header.Write("Score");
                    header.Write("Balance");
                    header.Write("Active");
                    header.Write("BirthDate");
                }

                await using (XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    row.Write(42);
                    row.Write(3.14);
                    row.Write(9999.99m);
                    row.Write(true);
                    row.Write(birth);
                }

                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.FromXlsx(ms);
            var rows = ExcelParser.FromAttributes<PrimitivesRow>().Parse(reader).ToList();
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
                XlsxSheetWriter sheet = wb.AddSheet("Sheet1");

                await using (XlsxRowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    header.Write("Active");
                }

                await using (XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    row.Write(false);
                }

                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.FromXlsx(ms);
            var rows = ExcelParser.FromAttributes<PrimitivesRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.False(rows[0].Active);
        }

        [Fact]
        public async Task NullStringWritesEmptyCell()
        {
            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                XlsxSheetWriter sheet = wb.AddSheet("Sheet1");

                await using (XlsxRowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    header.Write("Name");
                }

                await using (XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    row.Write((string?)null);
                }

                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.FromXlsx(ms);
            var rows = ExcelParser.FromAttributes<StringRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Null(rows[0].Name);
        }

        [Fact]
        public async Task NullableIntFilledRoundTrip()
        {
            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                XlsxSheetWriter sheet = wb.AddSheet("Sheet1");

                await using (XlsxRowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    header.Write("Quantity");
                    header.Write("EventDate");
                }

                await using (XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    row.Write((int?)77);
                    row.Write((DateTime?)null);
                }

                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.FromXlsx(ms);
            var rows = ExcelParser.FromAttributes<NullableRow>().Parse(reader).ToList();
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
                XlsxSheetWriter sheet = wb.AddSheet("Sheet1");

                await using (XlsxRowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    header.Write("Quantity");
                    header.Write("EventDate");
                }

                await using (XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    row.Write((int?)null);
                    row.Write((DateTime?)dt);
                }

                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.FromXlsx(ms);
            var rows = ExcelParser.FromAttributes<NullableRow>().Parse(reader).ToList();
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
                XlsxSheetWriter sheet = wb.AddSheet("Numbers");

                await using (XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
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

            await using var reader = Excel.FromXlsx(ms);
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


        [Fact]
        public async Task MultipleRowsRoundTrip()
        {
            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                XlsxSheetWriter sheet = wb.AddSheet("Sheet1");

                await using (XlsxRowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    header.Write("Name");
                }

                foreach (string name in stringArray)
                {
                    await using XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
                    row.Write(name);
                }

                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.FromXlsx(ms);
            var rows = ExcelParser.FromAttributes<StringRow>().Parse(reader).ToList();
            Assert.Equal(5, rows.Count);
            Assert.Equal("Alice", rows[0].Name);
            Assert.Equal("Bob", rows[1].Name);
            Assert.Equal("Carol", rows[2].Name);
            Assert.Equal("Dave", rows[3].Name);
            Assert.Equal("Eve", rows[4].Name);
        }


        [Fact]
        public async Task MultipleSheetsRoundTrip()
        {
            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                XlsxSheetWriter sheet1 = wb.AddSheet("Alpha");
                await using (XlsxRowWriter h1 = await sheet1.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    h1.Write("Value");
                }
                await using (XlsxRowWriter r1 = await sheet1.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    r1.Write("FromAlpha");
                }
                await sheet1.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

                XlsxSheetWriter sheet2 = wb.AddSheet("Beta");
                await using (XlsxRowWriter h2 = await sheet2.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    h2.Write("Value");
                }
                await using (XlsxRowWriter r2 = await sheet2.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    r2.Write("FromBeta");
                }
                await sheet2.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.FromXlsx(ms);
            var parser = ExcelParser.FromAttributes<MultiSheetRow>();
            var rowsSheet1 = parser.Parse(reader).ToList();
            Assert.Single(rowsSheet1);
            Assert.Equal("FromAlpha", rowsSheet1[0].Value);

            await using var reader2 = Excel.FromXlsx(ms);
            reader2.MoveToSheet(1);
            var rowsSheet2 = parser.Parse(reader2).ToList();
            Assert.Single(rowsSheet2);
            Assert.Equal("FromBeta", rowsSheet2[0].Value);
        }


        [Fact]
        public async Task SkipCreatesColumnGap()
        {
            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                XlsxSheetWriter sheet = wb.AddSheet("Sheet1");

                await using (XlsxRowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    header.Write("A");
                    header.Skip(1);
                    header.Write("C");
                }

                await using (XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    row.Write("aaa");
                    row.Skip(1);
                    row.Write("ccc");
                }

                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.FromXlsx(ms);
            var rows = ExcelParser.FromAttributes<SparseRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Equal("aaa", rows[0].A);
            Assert.Equal("ccc", rows[0].C);
        }

        [Fact]
        public async Task NegativeSkipThrows()
        {
            await using var ms = new MemoryStream();
            await using var wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true);
            XlsxSheetWriter sheet = wb.AddSheet("Sheet1");
            await using XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

            Assert.Throws<ArgumentOutOfRangeException>(() => row.Skip(-1));
        }

        [Theory]
        [InlineData("")]
        [InlineData("12345678901234567890123456789012")]
        [InlineData("Bad[Name")]
        public async Task InvalidSheetNameThrows(string name)
        {
            await using var ms = new MemoryStream();
            await using var wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true);

            Assert.Throws<ArgumentException>(() => wb.AddSheet(name));
        }


        [Fact]
        public async Task XmlSpecialCharsAreEscaped()
        {
            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                XlsxSheetWriter sheet = wb.AddSheet("Sheet1");

                await using (XlsxRowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    header.Write("Name");
                }

                await using (XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    row.Write("<Alice & \"Bob\">");
                }

                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.FromXlsx(ms);
            var rows = ExcelParser.FromAttributes<StringRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Equal("<Alice & \"Bob\">", rows[0].Name);
        }


        [Fact]
        public async Task DisposeAsyncWithoutEndAsyncProducesReadableWorkbook()
        {
            await using var ms = new MemoryStream();
            XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true);

            XlsxSheetWriter sheet = wb.AddSheet("Sheet1");

            await using (XlsxRowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
            {
                header.Write("Name");
            }

            await using (XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
            {
                row.Write("Auto");
            }

            await wb.DisposeAsync().ConfigureAwait(true);
            ms.Position = 0;

            await using var reader = Excel.FromXlsx(ms);
            var rows = ExcelParser.FromAttributes<StringRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Equal("Auto", rows[0].Name);
        }


        [Fact]
        public async Task HeaderRowTwoConfig()
        {
            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                XlsxSheetWriter sheet = wb.AddSheet("Sheet1");

                await using (XlsxRowWriter skip = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    skip.Write("ignored");
                }

                await using (XlsxRowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    header.Write("Name");
                }

                await using (XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    row.Write("HeaderRow2");
                }

                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.FromXlsx(ms);
            var config = new ExcelParserConfig { HeaderRow = 2 };
            var rows = ExcelParser.FromAttributes<StringRow>(config).Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Equal("HeaderRow2", rows[0].Name);
        }


        [Fact]
        public async Task AddSheetWhileSheetActiveThrows()
        {
            await using XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(new MemoryStream());

            XlsxSheetWriter sheet = wb.AddSheet("Sheet1");

            Assert.Throws<InvalidOperationException>(() => wb.AddSheet("Sheet2"));
            await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
        }

        [Fact]
        public async Task StartRowWhileRowActiveThrows()
        {
            await using XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(new MemoryStream());

            XlsxSheetWriter sheet = wb.AddSheet("Sheet1");

            XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

            await Assert.ThrowsAsync<InvalidOperationException>(() => sheet.StartRowAsync(TestContext.Current.CancellationToken).AsTask()).ConfigureAwait(true);

            await row.DisposeAsync().ConfigureAwait(true);
            await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
        }

        [Fact]
        public async Task WriteAfterRowDisposedThrows()
        {
            await using XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(new MemoryStream());

            XlsxSheetWriter sheet = wb.AddSheet("Sheet1");

            XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            await row.DisposeAsync().ConfigureAwait(true);

            Assert.Throws<ObjectDisposedException>(() => row.Write("late"));

            await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
        }


        [Fact]
        public async Task LargeWorkbookRoundTrip()
        {
            const int rowCount = 1000;

            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                XlsxSheetWriter sheet = wb.AddSheet("Sheet1");

                await using (XlsxRowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    header.Write("Age");
                }

                for (int i = 0; i < rowCount; i++)
                {
                    await using XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
                    row.Write(i);
                }

                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.FromXlsx(ms);
            var rows = ExcelParser.FromAttributes<PrimitivesRow>().Parse(reader).ToList();
            Assert.Equal(rowCount, rows.Count);
            for (int i = 0; i < rowCount; i++)
            {
                Assert.Equal(i, rows[i].Age);
            }
        }


        [Fact]
        public async Task FlushAsyncDoesNotThrow()
        {
            await using var ms = new MemoryStream();
            await using XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true);
            await wb.FlushAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            Assert.True(ms.Length >= 0);
        }


        [Fact]
        public async Task DisposeAsyncIsIdempotent()
        {
            await using var ms = new MemoryStream();
            XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true);

            XlsxSheetWriter sheet = wb.AddSheet("Sheet1");
            await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

            await wb.DisposeAsync().ConfigureAwait(true);
            Exception? ex = await Record.ExceptionAsync(async () =>
                await wb.DisposeAsync().ConfigureAwait(true)).ConfigureAwait(true);

            Assert.Null(ex);
        }

        [Fact]
        public async Task InlineStringWithEdgeWhitespaceUsesXmlSpacePreserve()
        {
            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                XlsxSheetWriter sheet = wb.AddSheet("Sheet1");
                await using (XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    row.Write(" leading and trailing ");
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            using var zip = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: true);
            using var stream = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
            using var text = new StreamReader(stream, Encoding.UTF8);
            string xml = await text.ReadToEndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

            Assert.Contains("<t xml:space=\"preserve\"> leading and trailing </t>", xml, StringComparison.Ordinal);
        }

        [Theory]
        [InlineData(double.NaN)]
        [InlineData(double.PositiveInfinity)]
        public async Task WriterRejectsNonFiniteNumbers(double value)
        {
            await using var ms = new MemoryStream();
            await using XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true);
            XlsxSheetWriter sheet = wb.AddSheet("Sheet1");
            await using XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

            Assert.Throws<ArgumentException>(() => row.Write(value));
        }

        [Fact]
        public async Task SheetWriterExtensionsWriteSynchronousAndAsyncRecords()
        {
            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                XlsxSheetWriter sheet = wb.AddSheet("Numbers");

                await sheet.WriteRecordsAsync([1, 2], static (row, value) => row.Write(value), TestContext.Current.CancellationToken);

                ISheetWriter<XlsxRowWriter> genericSheet = sheet;
                await genericSheet.WriteRecordsAsync(
                    ToAsync([3, 4]),
                    static (row, value) => row.Write(value),
                    TestContext.Current.CancellationToken);

                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            await using var reader = Excel.FromXlsx(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();
            for (int expected = 1; expected <= 4; expected++)
            {
                Assert.True(e.MoveNext());
                Assert.True(e.Current[0].TryParse(null, out int actual));
                Assert.Equal(expected, actual);
            }
            Assert.False(e.MoveNext());
        }

        [Fact]
        public async Task WriterEscapesXmlControlsAndLiteralExcelEscapeSequences()
        {
            await using var ms = await WriteWorkbookAsync(async wb =>
            {
                XlsxSheetWriter sheet = wb.AddSheet("Sheet1");
                await using (XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true))
                {
                    row.Write("a\u0001b _x0041_");
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
            }).ConfigureAwait(true);

            using var zip = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: true);
            using var stream = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
            using var text = new StreamReader(stream, Encoding.UTF8);
            string xml = await text.ReadToEndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

            Assert.Contains("a_x0001_b _x005F_x0041_", xml, StringComparison.Ordinal);
        }

        [Fact]
        public async Task WriterRejectsColumnsBeyondXfd()
        {
            await using var ms = new MemoryStream();
            await using XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true);
            XlsxSheetWriter sheet = wb.AddSheet("Sheet1");
            await using XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

            row.Skip(16_384);
            Assert.Throws<ExcelLimitExceededException>(() => row.Write("beyond XFD"));
        }

        [Fact]
        public async Task XlsxEndAsyncThrowsWhenLastRowStillActive()
        {
            await using var ms = new MemoryStream();
            await using XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true);
            XlsxSheetWriter sheet = wb.AddSheet("Sheet1");
            XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);

            await Assert.ThrowsAsync<InvalidOperationException>(
                () => sheet.EndAsync(TestContext.Current.CancellationToken).AsTask()).ConfigureAwait(true);

            await row.DisposeAsync().ConfigureAwait(true);
            await sheet.EndAsync(TestContext.Current.CancellationToken).ConfigureAwait(true);
        }
    }
}
