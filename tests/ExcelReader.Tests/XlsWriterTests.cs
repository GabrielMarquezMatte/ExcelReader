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
            Action<XlsWorkbookWriter> build, bool date1904 = false, CancellationToken ct = default)
        {
            var ms = new MemoryStream();
            await using (var wb = XlsWorkbookWriter.Create(ms, leaveOpen: true, date1904: date1904))
            {
                wb.Start();
                build(wb);
                await wb.EndAsync(ct);
            }
            return ms.ToArray();
        }

        [Fact]
        public async Task RoundTripsAllCellTypes()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var date = new DateTime(2024, 5, 6, 0, 0, 0, DateTimeKind.Unspecified);
            byte[] bytes = await WriteAsync(wb =>
            {
                var s = wb.AddSheet("Plan1");
                s.Start();
                using (var r = s.StartRow())
                {
                    r.Write("João");      // compressed Latin-1
                    r.Write(42);          // int via generic
                    r.Write(3.5);         // double
                    r.Write(true);        // bool
                    r.Write(date);        // date
                }
                s.End();
            }, ct: ct);

            using var reader = Excel.FromXls(new MemoryStream(bytes));
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("João", e.Current[0].GetString());
            Assert.True(e.Current[1].TryParse(Inv, out int age));
            Assert.Equal(42, age);
            Assert.True(e.Current[2].TryParse(Inv, out double price));
            Assert.Equal(3.5, price);
            Assert.Equal(CellType.Boolean, e.Current[3].Type);
            Assert.Equal("1", e.Current[3].GetString());
            Assert.Equal(CellType.Date, e.Current[4].Type);
            Assert.True(e.Current[4].TryGetDateTime(out DateTime parsed));
            Assert.Equal(date, parsed);
            Assert.False(e.MoveNext());
        }

        [Fact]
        public async Task NumericOverloadsRoundTrip()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] bytes = await WriteAsync(wb =>
            {
                var s = wb.AddSheet("Numbers");
                s.Start();
                using (var r = s.StartRow())
                {
                    r.Write(123);
                    r.Write(1234567890123L);
                    r.Write(1.5f);
                    r.Write(2.75d);
                    r.Write(12.5m);
                    r.Write((int?)null);
                    r.Write((long?)7L);
                    r.Write((double?)null);
                    r.Write((decimal?)8.25m);
                }
                s.End();
            }, ct: ct);

            using var reader = Excel.FromXls(new MemoryStream(bytes));
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.True(e.Current[0].TryParse(Inv, out int intValue));
            Assert.Equal(123, intValue);
            Assert.True(e.Current[1].TryParse(Inv, out long longValue));
            Assert.Equal(1234567890123L, longValue);
            Assert.True(e.Current[2].TryParse(Inv, out double floatValue));
            Assert.Equal(1.5, floatValue);
            Assert.True(e.Current[3].TryParse(Inv, out double doubleValue));
            Assert.Equal(2.75, doubleValue);
            Assert.True(e.Current[4].TryParse(Inv, out decimal decimalValue));
            Assert.Equal(12.5m, decimalValue);
            Assert.Equal(CellType.Empty, e.Current[5].Type);
            Assert.True(e.Current[6].TryParse(Inv, out long nullableLong));
            Assert.Equal(7L, nullableLong);
            Assert.Equal(CellType.Empty, e.Current[7].Type);
            Assert.True(e.Current[8].TryParse(Inv, out decimal nullableDecimal));
            Assert.Equal(8.25m, nullableDecimal);
            Assert.False(e.MoveNext());
        }

        [Fact]
        public async Task NullsAndSkipLeaveCellsEmpty()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] bytes = await WriteAsync(wb =>
            {
                var s = wb.AddSheet("S1");
                s.Start();
                using (var r = s.StartRow())
                {
                    r.Write("A");
                    r.Write((string?)null);  // empty
                    r.Skip();                // empty
                    r.Write("D");
                }
                s.End();
            }, ct: ct);

            using var reader = Excel.FromXls(new MemoryStream(bytes));
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("A", e.Current[0].GetString());
            Assert.Equal(CellType.Empty, e.Current[1].Type);
            Assert.Equal(CellType.Empty, e.Current[2].Type);
            Assert.Equal("D", e.Current[3].GetString());
        }

        [Fact]
        public async Task WritesMultipleSheetsWithCorrectOffsets()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] bytes = await WriteAsync(wb =>
            {
                var first = wb.AddSheet("First");
                first.Start();
                using (var r = first.StartRow()) { r.Write("one"); }
                first.End();

                var second = wb.AddSheet("Ωmega");
                second.Start();
                using (var r = second.StartRow()) { r.Write("two"); }
                second.End();
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
            byte[] bytes = await WriteAsync(wb =>
            {
                var s = wb.AddSheet("Big");
                s.Start();
                for (int i = 0; i < rows; i++)
                {
                    using var r = s.StartRow();
                    r.Write($"r{i}");
                    r.Write(i);
                    r.Write(i * 0.25);
                }
                s.End();
            }, ct: ct);

            Assert.True(bytes.Length > 512 * 4, "workbook should span several OLE sectors");
            using var reader = Excel.FromXls(new MemoryStream(bytes));
            using var e = reader.GetEnumerator();
            int read = 0;
            while (e.MoveNext())
            {
                Assert.Equal($"r{read}", e.Current[0].GetString());
                Assert.True(e.Current[1].TryParse(Inv, out int n));
                Assert.Equal(read, n);
                Assert.True(e.Current[2].TryParse(Inv, out double d));
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
            byte[] bytes = await WriteAsync(wb =>
            {
                var s = wb.AddSheet("S1");
                s.Start();
                using (var r = s.StartRow()) { r.Write(date); }
                s.End();
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
            await using var wb = XlsWorkbookWriter.Create(ms, leaveOpen: true);
            wb.Start();
            var s = wb.AddSheet("S1");
            s.Start();
            using var r = s.StartRow();
            r.Skip(256); // column index 256 (0-based) is out of range
            Assert.Throws<InvalidOperationException>(() => r.Write("x"));
        }

        [Fact]
        public async Task SheetNameTooLongThrows()
        {
            var ms = new MemoryStream();
            await using var wb = XlsWorkbookWriter.Create(ms, leaveOpen: true);
            wb.Start();
            Assert.Throws<ArgumentException>(() => wb.AddSheet(new string('x', 32)));
        }

        [Fact]
        public async Task EmptyWorkbookThrows()
        {
            var ms = new MemoryStream();
            await using var wb = XlsWorkbookWriter.Create(ms, leaveOpen: true);
            wb.Start();
            await Assert.ThrowsAsync<InvalidOperationException>(async () => await wb.EndAsync(TestContext.Current.CancellationToken));
        }

        [Fact]
        public async Task LargeWorkbookForcesDifatSectors()
        {
            // The OLE container only allocates DIFAT sectors once the FAT exceeds the 109 entries that
            // fit in the header — that needs a workbook stream past ~7.1 MB (>~13,900 512-byte sectors).
            // 64000 rows x 8 NUMBER cells (18 bytes each) ≈ 9.2 MB, comfortably over the threshold.
            CancellationToken ct = TestContext.Current.CancellationToken;
            const int rows = 64000;
            const int cols = 8;
            byte[] bytes = await WriteAsync(wb =>
            {
                var s = wb.AddSheet("Big");
                s.Start();
                for (int i = 0; i < rows; i++)
                {
                    using var r = s.StartRow();
                    r.Write(i);
                    for (int c = 1; c < cols; c++)
                    {
                        r.Write(i + (c * 0.5));
                    }
                }
                s.End();
            }, ct: ct);

            Assert.True(bytes.Length > 7 * 1024 * 1024, "workbook must exceed the DIFAT threshold");

            using var reader = Excel.FromXls(new MemoryStream(bytes));
            using var e = reader.GetEnumerator();
            int read = 0;
            while (e.MoveNext())
            {
                if (read == 0 || read == rows - 1)
                {
                    Assert.True(e.Current[0].TryParse(Inv, out int first));
                    Assert.Equal(read, first);
                    Assert.True(e.Current[cols - 1].TryParse(Inv, out double last));
                    Assert.Equal(read + ((cols - 1) * 0.5), last);
                }
                read++;
            }
            Assert.Equal(rows, read);
        }

        [Fact]
        public async Task RowOverflowAutoSplitsIntoNewSheet()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            const int rows = 65537; // one past the BIFF8 per-sheet cap (65536)
            byte[] bytes = await WriteAsync(wb =>
            {
                var s = wb.AddSheet("S");
                s.Start();
                for (int i = 0; i < rows; i++)
                {
                    using var r = s.StartRow();
                    r.Write(i);
                }
                s.End();
            }, ct: ct);

            using var reader = Excel.FromXls(new MemoryStream(bytes));
            Assert.Equal(2, reader.SheetCount);

            int totalRows = 0;
            for (int sheet = 0; sheet < reader.SheetCount; sheet++)
            {
                reader.MoveToSheet(sheet);
                using var e = reader.GetEnumerator();
                while (e.MoveNext())
                {
                    totalRows++;
                }
            }
            Assert.Equal(rows, totalRows);
        }
    }
}
