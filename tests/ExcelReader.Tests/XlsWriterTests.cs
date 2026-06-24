using System.Globalization;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    public class XlsWriterTests
    {
        private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        private static async Task<byte[]> WriteAsync(
            Func<XlsWorkbookWriter, Task> build, bool date1904 = false, CancellationToken ct = default)
        {
            var ms = new MemoryStream();
            await using (var wb = await XlsWorkbookWriter.CreateAsync(ms, leaveOpen: true, date1904: date1904, ct: ct))
            {
                await wb.StartAsync(ct);
                await build(wb);
                await wb.EndAsync(ct);
            }
            return ms.ToArray();
        }

        [Fact]
        public async Task RoundTripsAllCellTypes()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var date = new DateTime(2024, 5, 6, 0, 0, 0, DateTimeKind.Unspecified);
            byte[] bytes = await WriteAsync(async wb =>
            {
                var s = wb.AddSheet("Plan1");
                await s.StartAsync(ct);
                await using (var r = await s.StartRowAsync(ct))
                {
                    r.Write("João");      // compressed Latin-1
                    r.Write(42);          // int via generic
                    r.Write(3.5);         // double
                    r.Write(true);        // bool
                    r.Write(date);        // date
                }
                await s.EndAsync(ct);
            }, ct: ct);

            using var reader = Excel.FromXls(new MemoryStream(bytes));
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            var row = e.Current;
            Assert.Equal("João", row[0].GetString());
            Assert.True(row[1].TryParse<int>(Inv, out int age));
            Assert.Equal(42, age);
            Assert.True(row[2].TryParse<double>(Inv, out double price));
            Assert.Equal(3.5, price);
            Assert.Equal(CellType.Boolean, row[3].Type);
            Assert.Equal("1", row[3].GetString());
            Assert.Equal(CellType.Date, row[4].Type);
            Assert.True(row[4].TryGetDateTime(out DateTime parsed));
            Assert.Equal(date, parsed);
            Assert.False(e.MoveNext());
        }

        [Fact]
        public async Task NullsAndSkipLeaveCellsEmpty()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] bytes = await WriteAsync(async wb =>
            {
                var s = wb.AddSheet("S1");
                await s.StartAsync(ct);
                await using (var r = await s.StartRowAsync(ct))
                {
                    r.Write("A");
                    r.Write((string?)null);  // empty
                    r.Skip();                // empty
                    r.Write("D");
                }
                await s.EndAsync(ct);
            }, ct: ct);

            using var reader = Excel.FromXls(new MemoryStream(bytes));
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            var row = e.Current;
            Assert.Equal("A", row[0].GetString());
            Assert.Equal(CellType.Empty, row[1].Type);
            Assert.Equal(CellType.Empty, row[2].Type);
            Assert.Equal("D", row[3].GetString());
        }

        [Fact]
        public async Task WritesMultipleSheetsWithCorrectOffsets()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] bytes = await WriteAsync(async wb =>
            {
                var first = wb.AddSheet("First");
                await first.StartAsync(ct);
                await using (var r = await first.StartRowAsync(ct)) { r.Write("one"); }
                await first.EndAsync(ct);

                var second = wb.AddSheet("Ωmega");
                await second.StartAsync(ct);
                await using (var r = await second.StartRowAsync(ct)) { r.Write("two"); }
                await second.EndAsync(ct);
            }, ct: ct);

            using var reader = Excel.FromXls(new MemoryStream(bytes));
            Assert.Equal(2, reader.SheetCount);

            reader.MoveToSheet(0);
            Assert.Equal("First", reader.SheetName);
            using (var e = reader.GetEnumerator()) { Assert.True(e.MoveNext()); Assert.Equal("one", e.Current[0].GetString()); }

            Assert.True(reader.TryMoveToSheet("Ωmega"));
            using (var e = reader.GetEnumerator()) { Assert.True(e.MoveNext()); Assert.Equal("two", e.Current[0].GetString()); }
        }

        [Fact]
        public async Task ManyRowsRoundTripWithExactCount()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            const int rows = 1000;
            byte[] bytes = await WriteAsync(async wb =>
            {
                var s = wb.AddSheet("Big");
                await s.StartAsync(ct);
                for (int i = 0; i < rows; i++)
                {
                    await using var r = await s.StartRowAsync(ct);
                    r.Write($"r{i}");
                    r.Write(i);
                    r.Write(i * 0.25);
                }
                await s.EndAsync(ct);
            }, ct: ct);

            Assert.True(bytes.Length > 512 * 4, "workbook should span several OLE sectors");
            using var reader = Excel.FromXls(new MemoryStream(bytes));
            using var e = reader.GetEnumerator();
            int read = 0;
            while (e.MoveNext())
            {
                var row = e.Current;
                Assert.Equal($"r{read}", row[0].GetString());
                Assert.True(row[1].TryParse<int>(Inv, out int n));
                Assert.Equal(read, n);
                Assert.True(row[2].TryParse<double>(Inv, out double d));
                Assert.Equal(read * 0.25, d);
                read++;
            }
            Assert.Equal(rows, read);
        }

        [Fact]
        public async Task Date1904RoundTrips()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var date = new DateTime(1990, 7, 15, 0, 0, 0, DateTimeKind.Unspecified);
            byte[] bytes = await WriteAsync(async wb =>
            {
                var s = wb.AddSheet("S1");
                await s.StartAsync(ct);
                await using (var r = await s.StartRowAsync(ct)) { r.Write(date); }
                await s.EndAsync(ct);
            }, date1904: true, ct: ct);

            using var reader = Excel.FromXls(new MemoryStream(bytes));
            Assert.True(reader.IsDate1904);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.True(e.Current[0].TryGetDateTime(reader.IsDate1904, out DateTime parsed));
            Assert.Equal(date, parsed);
        }

        [Fact]
        public async Task ColumnBeyondLimitThrows()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var ms = new MemoryStream();
            await using var wb = await XlsWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: ct);
            await wb.StartAsync(ct);
            var s = wb.AddSheet("S1");
            await s.StartAsync(ct);
            await using var r = await s.StartRowAsync(ct);
            r.Skip(256); // column index 256 (0-based) is out of range
            Assert.Throws<InvalidOperationException>(() => r.Write("x"));
        }

        [Fact]
        public async Task SheetNameTooLongThrows()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var ms = new MemoryStream();
            await using var wb = await XlsWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: ct);
            await wb.StartAsync(ct);
            Assert.Throws<ArgumentException>(() => wb.AddSheet(new string('x', 32)));
        }

        [Fact]
        public async Task EmptyWorkbookThrows()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var ms = new MemoryStream();
            await using var wb = await XlsWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: ct);
            await wb.StartAsync(ct);
            await Assert.ThrowsAsync<InvalidOperationException>(async () => await wb.EndAsync(ct));
        }
    }
}
