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
        public async Task NullableAndGenericOverloadsRoundTrip()
        {
            DateTime date = new(2026, 6, 27, 0, 0, 0, DateTimeKind.Unspecified);
            await using var ms = await WriteAsync(async wb =>
            {
                XlsbSheetWriter sheet = wb.AddSheet("MoreNumbers");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsbRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    row.Write((bool?)true);
                    row.Write((bool?)null);
                    row.Write((DateTime?)date);
                    row.Write((DateTime?)null);
                    row.Write((float?)1.25f);
                    row.Write((float?)null);
                    row.Write<short>(6);
                    row.Write((short?)7);
                    row.Write<short>(null);
                    row.Skip(0);
                    row.Write<byte>(8);
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
            });

            await using var reader = Excel.FromXlsb(ms);
            using XlsbReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Boolean, e.Current[0].Type);
            Assert.Equal(CellType.Empty, e.Current[1].Type);
            Assert.True(e.Current[2].TryGetDateTime(out DateTime parsed));
            Assert.Equal(date, parsed);
            Assert.Equal(CellType.Empty, e.Current[3].Type);
            Assert.True(e.Current[4].TryParse(null, out double nullableFloat));
            Assert.Equal(1.25, nullableFloat);
            Assert.Equal(CellType.Empty, e.Current[5].Type);
            Assert.True(e.Current[6].TryParse(null, out short genericShort));
            Assert.Equal(6, genericShort);
            Assert.True(e.Current[7].TryParse(null, out short nullableShort));
            Assert.Equal(7, nullableShort);
            Assert.Equal(CellType.Empty, e.Current[8].Type);
            Assert.True(e.Current[9].TryParse(null, out byte genericByte));
            Assert.Equal(8, genericByte);
            Assert.False(e.MoveNext());
        }

        [Fact]
        public async Task NegativeSkipThrows()
        {
            await using var ms = await WriteAsync(async wb =>
            {
                XlsbSheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsbRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    Assert.Throws<ArgumentOutOfRangeException>(() => row.Skip(-1));
                    row.Write("ok");
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
            });

            await using var reader = Excel.FromXlsb(ms);
            using XlsbReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("ok", e.Current[0].GetString());
            Assert.False(e.MoveNext());
        }

        [Theory]
        [InlineData("")]
        [InlineData("12345678901234567890123456789012")]
        [InlineData("Bad[Name")]
        public async Task InvalidSheetNameThrows(string name)
        {
            await using var wb = await XlsbWorkbookWriter.CreateAsync(new MemoryStream(), ct: TestContext.Current.CancellationToken);
            await wb.StartAsync(TestContext.Current.CancellationToken);

            Assert.Throws<ArgumentException>(() => wb.AddSheet(name));
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

        [Fact]
        public void XlsbCellCreateOverloadsSetExpectedKindsAndValues()
        {
            DateTime date = new(2026, 6, 27, 0, 0, 0, DateTimeKind.Unspecified);

            Assert.Equal(XlsbCellKind.Empty, XlsbCell.Create((bool?)null).Kind);
            Assert.Equal(XlsbCellKind.Empty, XlsbCell.Create((DateTime?)null).Kind);
            Assert.Equal(XlsbCellKind.Empty, XlsbCell.Create((int?)null).Kind);
            Assert.Equal(XlsbCellKind.Empty, XlsbCell.Create((long?)null).Kind);
            Assert.Equal(XlsbCellKind.Empty, XlsbCell.Create((float?)null).Kind);
            Assert.Equal(XlsbCellKind.Empty, XlsbCell.Create((double?)null).Kind);
            Assert.Equal(XlsbCellKind.Empty, XlsbCell.Create((decimal?)null).Kind);
            Assert.Equal(XlsbCellKind.Empty, XlsbCell.Create<int>(null).Kind);

            XlsbCell text = XlsbCell.Create("text");
            Assert.Equal(XlsbCellKind.String, text.Kind);
            Assert.Equal("text", text.Text);

            XlsbCell boolean = XlsbCell.Create((bool?)true);
            Assert.Equal(XlsbCellKind.Boolean, boolean.Kind);
            Assert.True(boolean.Boolean);

            XlsbCell dateCell = XlsbCell.Create((DateTime?)date);
            Assert.Equal(XlsbCellKind.Date, dateCell.Kind);
            Assert.Equal(date.ToOADate(), dateCell.Number);

            Assert.Equal(1, XlsbCell.Create((int?)1).Number);
            Assert.Equal(2L, XlsbCell.Create((long?)2L).Number);
            Assert.Equal(3.5f, XlsbCell.Create((float?)3.5f).Number);
            Assert.Equal(4.5d, XlsbCell.Create((double?)4.5d).Number);
            Assert.Equal(5.5d, XlsbCell.Create((decimal?)5.5m).Number);
            Assert.Equal(6d, XlsbCell.Create<short>(6).Number);
            Assert.Equal(7d, XlsbCell.Create((short?)7).Number);
        }

        // XlsbRowWriter.Skip had no bound anywhere in the class before this fix — not even
        // downstream at Write time, unlike XLS. BIFF12/.xlsb uses the modern 16,384-column grid.
        [Fact]
        public async Task SkipBeyondColumnLimitThrows()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using MemoryStream ms = new();
            await using XlsbWorkbookWriter wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: ct);
            await wb.StartAsync(ct);
            XlsbSheetWriter sheet = wb.AddSheet("Sheet1");
            await sheet.StartAsync(ct);
            await using XlsbRowWriter row = await sheet.StartRowAsync(ct);

            Assert.Throws<ExcelLimitExceededException>(() => row.Skip(16_385));
        }

        // Write<T> falls back through XlsbRowWriter.ToDouble, whose final fallback used to
        // silently return 0.0 for a T that formats as non-numeric text instead of throwing.
        [Fact]
        public async Task WriteNonNumericFormattableThrowsArgumentException()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using MemoryStream ms = new();
            await using XlsbWorkbookWriter wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: ct);
            await wb.StartAsync(ct);
            XlsbSheetWriter sheet = wb.AddSheet("Sheet1");
            await sheet.StartAsync(ct);
            await using XlsbRowWriter row = await sheet.StartRowAsync(ct);

            Assert.Throws<ArgumentException>(() => row.Write(new NonNumericFormattable()));
        }

        // EndAsync used to flip _state to Ended before the zero-sheet check threw, so a failed
        // EndAsync left DisposeAsync's state check matching neither Started nor Created — _zip was
        // never disposed on this path. After the reorder, _state stays Started when EndAsync throws,
        // so DisposeAsync correctly falls into its Started branch instead of silently skipping _zip.
        [Fact]
        public async Task DisposeAfterFailedZeroSheetEndDoesNotThrow()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using MemoryStream ms = new();
            XlsbWorkbookWriter wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: ct);
            await wb.StartAsync(ct);

            await Assert.ThrowsAsync<InvalidOperationException>(() => wb.EndAsync(ct).AsTask());

            // Before the fix, _state was already Ended here, so DisposeAsync's Started/Created branches
            // both missed and _zip leaked silently — this call completing without throwing or hanging is
            // the observable half of the fix; the other half (that _zip's Dispose actually ran) is
            // internal and not directly assertable from the test.
            await wb.DisposeAsync();
        }
    }
}
