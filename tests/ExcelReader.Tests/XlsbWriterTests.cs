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
    }
}