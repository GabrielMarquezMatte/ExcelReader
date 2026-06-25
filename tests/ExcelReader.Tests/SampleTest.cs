using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class SampleTest
    {
        [Fact]
        public void ReadsSharedStringsStylesAndNumbers()
        {
            string path = Path.Combine(AppContext.BaseDirectory, "data", "sample.xlsx");
            using var reader = Excel.FromFile(path);

            int r = 0;
            foreach (var row in reader)
            {
                if (r == 0)
                {
                    // Header row: all shared strings, resolved to text.
                    Assert.Equal("file", row[0].GetString());
                    Assert.Equal("changes", row[1].GetString());
                    Assert.Equal("lines_added", row[2].GetString());
                    Assert.Equal("lines_deleted", row[3].GetString());
                    Assert.Equal(CellType.ExcelString, row[0].Type);
                }
                else if (r == 1)
                {
                    // Styled shared string (s="1") + UTF-8 numeric parse.
                    Assert.Equal("global.json", row[0].GetString());
                    Assert.Equal(1, row[0].StyleIndex);
                    Assert.True(row[1].TryParse(null, out int n));
                    Assert.Equal(2, n);
                    Assert.Equal(CellType.Number, row[1].Type);
                }
                r++;
            }
            Assert.Equal(3, r);
        }

        [Fact]
        public async Task HandlesSparseCellsAndBufferGrowth()
        {
            // A row with a gap (no B), and an inline string far larger than the 64 KB scan buffer
            // to exercise compaction/grow and the cross-boundary </c> search.
            string big = new('x', 100_000);
            await using var ms = await TypedWorkbook.BuildAsync(
                [10, new Gap(), 30],
                [big]);

            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            await using var enumerator = await reader.GetAsyncEnumeratorAsync(TestContext.Current.CancellationToken);
            int r = 0;
            while (await enumerator.MoveNextAsync())
            {
                if (r == 0)
                {
                    Assert.Equal(3, enumerator.Current.ColumnCount);
                    Assert.True(enumerator.Current[0].TryParse(null, out int a));
                    Assert.Equal(10, a);
                    Assert.Equal(CellType.Empty, enumerator.Current[1].Type); // the gap
                    Assert.True(enumerator.Current[2].TryParse(null, out int c));
                    Assert.Equal(30, c);
                }
                else if (r == 1)
                {
                    Assert.Equal(big.Length, enumerator.Current[0].GetString().Length);
                }
                r++;
            }
            Assert.Equal(2, r);
        }

        [Fact]
        public void CellsEnumeratesOnlyPopulatedCellsInColumnOrder()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1"><v>10</v></c><c r="AA1"><v>30</v></c></row>""");

            using var reader = Excel.From(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());

            var row = e.Current;
            Assert.Equal(27, row.ColumnCount);

            var columns = new List<int>();
            var values = new List<int>();
            foreach (var rowCell in row.Cells)
            {
                columns.Add(rowCell.ColumnIndex);
                Assert.True(rowCell.Value.TryParse(null, out int value));
                values.Add(value);
            }

            Assert.Equal([0, 26], columns);
            Assert.Equal([10, 30], values);
        }

        [Fact]
        public void DecodesXmlEntitiesInSharedStrings()
        {
            // Shared strings are a raw-XML feature WorkbookWriter does not emit.
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="s"><v>0</v></c></row>""",
                sharedStrings: "<si><t>a &amp; b &lt;tag&gt; &#65;</t></si>");

            using var reader = Excel.From(ms);
            using var enumerator = reader.GetEnumerator();
            if (!enumerator.MoveNext())
            {
                Assert.Fail("Expected at least one row");
            }
            var row = enumerator.Current;
            Assert.Equal("a & b <tag> A", row[0].GetString());
        }

        [Fact]
        public void DecodePassesThroughLoneAmpersandAndUnknownEntities()
        {
            // Exercises Decode's bulk-copy paths: a '&' with no terminator, and an unrecognized entity.
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>""",
                sharedStrings: "<si><t>a&b</t></si><si><t>x&foo;y</t></si>");

            using var reader = Excel.From(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            var row = e.Current;
            Assert.Equal("a&b", row[0].GetString());
            Assert.Equal("x&foo;y", row[1].GetString());
        }

        [Fact]
        public void DetectsDateStylesAndConvertsSerial()
        {
            // s="1" -> cellXfs[1] -> builtin numFmtId 14 (date); s="2" -> custom 164 (date); s="0" -> General.
            // Custom number formats are a raw-XML feature WorkbookWriter does not emit.
            const string styles =
                """<styleSheet><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm"/></numFmts>""" +
                """<cellXfs count="3"><xf numFmtId="0"/><xf numFmtId="14"/><xf numFmtId="164"/></cellXfs></styleSheet>""";
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" s="1"><v>45292</v></c><c r="B1" s="2"><v>45292.5</v></c><c r="C1" s="0"><v>45292</v></c></row>""",
                styles: styles);

            using var reader = Excel.From(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            var row = e.Current;

            Assert.Equal(CellType.Date, row[0].Type);
            Assert.True(row[0].TryGetDateTime(out var d0));
            Assert.Equal(new DateTime(2024, 1, 1, 0, 0, 0, DateTimeKind.Unspecified), d0);

            Assert.Equal(CellType.Date, row[1].Type);
            Assert.True(row[1].TryGetDateTime(out var d1));
            Assert.Equal(new DateTime(2024, 1, 1, 12, 0, 0, DateTimeKind.Unspecified), d1);

            // Same serial, no date style -> plain number.
            Assert.Equal(CellType.Number, row[2].Type);
        }

        [Fact]
        public async Task AsyncReadsSampleFileLikeSyncPath()
        {
            var ct = TestContext.Current.CancellationToken;
            string path = Path.Combine(AppContext.BaseDirectory, "data", "sample.xlsx");
            await using var reader = await Excel.FromFileAsync(path, ct);
            await using var e = await reader.GetAsyncEnumeratorAsync(ct);

            int r = 0;
            while (await e.MoveNextAsync())
            {
                if (r == 0)
                {
                    Assert.Equal("file", e.Current[0].GetString());
                    Assert.Equal("lines_deleted", e.Current[3].GetString());
                }
                else if (r == 1)
                {
                    Assert.Equal("global.json", e.Current[0].GetString());
                    Assert.True(e.Current[1].TryParse(null, out int n));
                    Assert.Equal(2, n);
                }
                r++;
            }
            Assert.Equal(3, r);
        }

        [Fact]
        public async Task AsyncHandlesBufferGrowthAndDates()
        {
            // Inline string far larger than the 64 KB scan buffer exercises FillAsync compaction/grow and
            // the cross-refill </c> search on the async path; the date cell exercises style detection.
            var ct = TestContext.Current.CancellationToken;
            string big = new('x', 100_000);
            await using var ms = await TypedWorkbook.BuildAsync(
                [DateTime.FromOADate(45292)],
                [big]);

            await using var reader = await Excel.FromAsync(ms, ct: ct);
            await using var e = await reader.GetAsyncEnumeratorAsync(ct);

            Assert.True(await e.MoveNextAsync());
            Assert.Equal(CellType.Date, e.Current[0].Type);
            Assert.True(e.Current[0].TryGetDateTime(out var d));
            Assert.Equal(new DateTime(2024, 1, 1, 0, 0, 0, DateTimeKind.Unspecified), d);

            Assert.True(await e.MoveNextAsync());
            Assert.Equal(big.Length, e.Current[0].GetString().Length);

            Assert.False(await e.MoveNextAsync());
        }
    }
}
