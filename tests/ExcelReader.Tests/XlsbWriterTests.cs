using System.Globalization;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    public class XlsbWriterTests
    {
        private sealed class PersonRow
        {
            public string? Name { get; set; }
            public int Age { get; set; }
            public bool Active { get; set; }
            public DateTime BirthDate { get; set; }
        }

        private sealed class SparseRow
        {
            public string? A { get; set; }
            public string? C { get; set; }
        }

        private static async Task<MemoryStream> WriteAsync(Func<XlsbWorkbookWriter, Task> build, bool date1904 = false)
        {
            MemoryStream ms = new();
            await using var wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true, date1904: date1904, ct: TestContext.Current.CancellationToken);
            await wb.StartAsync(TestContext.Current.CancellationToken);
            await build(wb);
            await wb.EndAsync(TestContext.Current.CancellationToken);
            ms.Position = 0;
            return ms;
        }

        [Fact]
        public async Task RoundTripsThroughReaderAndParser()
        {
            var birth = new DateTime(1990, 6, 15, 0, 0, 0, DateTimeKind.Unspecified);
            await using var ms = await WriteAsync(async wb =>
            {
                XlsbSheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsbRowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    header.Write("Name");
                    header.Write("Age");
                    header.Write("Active");
                    header.Write("BirthDate");
                }
                await using (XlsbRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    row.Write("Alice");
                    row.Write(42);
                    row.Write(true);
                    row.Write(birth);
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
            });

            await using var reader = Excel.FromXlsb(ms);
            var rows = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Equal("Alice", rows[0].Name);
            Assert.Equal(42, rows[0].Age);
            Assert.True(rows[0].Active);
            Assert.Equal(birth, rows[0].BirthDate);
        }

        [Fact]
        public async Task OpenDetectsWrittenXlsb()
        {
            await using var ms = await WriteAsync(async wb =>
            {
                XlsbSheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsbRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    row.Write("x");
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
            });

            await using var reader = Excel.Open(ms);
            Assert.IsType<XlsbReader>(reader);
        }

        [Fact]
        public async Task SkipCreatesColumnGap()
        {
            await using var ms = await WriteAsync(async wb =>
            {
                XlsbSheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsbRowWriter header = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    header.Write("A");
                    header.Skip();
                    header.Write("C");
                }
                await using (XlsbRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    row.Write("aaa");
                    row.Skip();
                    row.Write("ccc");
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
            });

            await using var reader = Excel.FromXlsb(ms);
            var rows = new ExcelParser<SparseRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Equal("aaa", rows[0].A);
            Assert.Equal("ccc", rows[0].C);
        }

        [Fact]
        public async Task Date1904RoundTrips()
        {
            var date = new DateTime(1904, 1, 2, 0, 0, 0, DateTimeKind.Unspecified);
            await using var ms = await WriteAsync(async wb =>
            {
                XlsbSheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsbRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    row.Write(date);
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
            }, date1904: true);

            await using var reader = Excel.FromXlsb(ms);
            Assert.True(reader.IsDate1904);
            using XlsbReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.True(e.Current[0].TryGetDateTime(reader.IsDate1904, out DateTime parsed));
            Assert.Equal(date, parsed);
        }

        [Fact]
        public async Task NumericOverloadsRoundTrip()
        {
            await using var ms = await WriteAsync(async wb =>
            {
                XlsbSheetWriter sheet = wb.AddSheet("Numbers");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsbRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    row.Write(123);
                    row.Write(1234567890123L);
                    row.Write(1.5f);
                    row.Write(2.75d);
                    row.Write(12.5m);
                    row.Write((int?)null);
                    row.Write((long?)7L);
                    row.Write((double?)null);
                    row.Write((decimal?)8.25m);
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
            });

            await using var reader = Excel.FromXlsb(ms);
            using XlsbReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.True(e.Current[0].TryParse(null, out int intValue));
            Assert.Equal(123, intValue);
            Assert.True(e.Current[1].TryParse(null, out long longValue));
            Assert.Equal(1234567890123L, longValue);
            Assert.True(e.Current[2].TryParse(null, out double floatValue));
            Assert.Equal(1.5, floatValue);
            Assert.True(e.Current[3].TryParse(null, out double doubleValue));
            Assert.Equal(2.75, doubleValue);
            Assert.True(e.Current[4].TryParse(CultureInfo.InvariantCulture, out decimal decimalValue));
            Assert.Equal(12.5m, decimalValue);
            Assert.Equal(CellType.Empty, e.Current[5].Type);
            Assert.True(e.Current[6].TryParse(null, out long nullableLong));
            Assert.Equal(7L, nullableLong);
            Assert.Equal(CellType.Empty, e.Current[7].Type);
            Assert.True(e.Current[8].TryParse(CultureInfo.InvariantCulture, out decimal nullableDecimal));
            Assert.Equal(8.25m, nullableDecimal);
            Assert.False(e.MoveNext());
        }

        [Fact]
        public async Task BulkXlsbCellRowRoundTrip()
        {
            var birth = new DateTime(1990, 6, 15, 0, 0, 0, DateTimeKind.Unspecified);
            await using var ms = await WriteAsync(async wb =>
            {
                XlsbSheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                XlsbCell[] header =
                [
                    XlsbCell.Create("Name"),
                    XlsbCell.Create("Age"),
                    XlsbCell.Create("Active"),
                    XlsbCell.Create("BirthDate"),
                ];
                XlsbCell[] row =
                [
                    XlsbCell.Create("Alice"),
                    XlsbCell.Create(42),
                    XlsbCell.Create(true),
                    XlsbCell.Create(birth),
                ];
                sheet.WriteRow(header);
                sheet.WriteRow(row);
                await sheet.EndAsync(TestContext.Current.CancellationToken);
            });

            await using var reader = Excel.FromXlsb(ms);
            var rows = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Equal("Alice", rows[0].Name);
            Assert.Equal(42, rows[0].Age);
            Assert.True(rows[0].Active);
            Assert.Equal(birth, rows[0].BirthDate);
        }

        [Fact]
        public async Task BulkXlsbCellRowPreservesEmptyCells()
        {
            await using var ms = await WriteAsync(async wb =>
            {
                XlsbSheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                XlsbCell[] header = [XlsbCell.Create("A"), XlsbCell.Empty, XlsbCell.Create("C")];
                XlsbCell[] row = [XlsbCell.Create("aaa"), XlsbCell.Create((string?)null), XlsbCell.Create("ccc")];
                sheet.WriteRow(header);
                sheet.WriteRow(row);
                await sheet.EndAsync(TestContext.Current.CancellationToken);
            });

            await using var reader = Excel.FromXlsb(ms);
            var rows = new ExcelParser<SparseRow>().Parse(reader).ToList();
            Assert.Single(rows);
            Assert.Equal("aaa", rows[0].A);
            Assert.Equal("ccc", rows[0].C);
        }

        [Fact]
        public async Task BulkXlsbCellDate1904RoundTrips()
        {
            var date = new DateTime(1904, 1, 2, 0, 0, 0, DateTimeKind.Unspecified);
            await using var ms = await WriteAsync(async wb =>
            {
                XlsbSheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                XlsbCell[] row = [XlsbCell.Create(date)];
                sheet.WriteRow(row);
                await sheet.EndAsync(TestContext.Current.CancellationToken);
            }, date1904: true);

            await using var reader = Excel.FromXlsb(ms);
            Assert.True(reader.IsDate1904);
            using XlsbReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.True(e.Current[0].TryGetDateTime(reader.IsDate1904, out DateTime parsed));
            Assert.Equal(date, parsed);
        }
    }
}
