using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    // Regression coverage for two SpreadsheetML dialects non-Excel producers emit:
    //   3.4 — every element carries a namespace prefix (<x:worksheet>/<x:row>/<x:c>/...).
    //   3.7 — ISO-8601 date cells typed t="d" (a bare serial is NOT what the <v> holds).
    public class NamespacePrefixAndIsoDateTests
    {
        // ---- 3.4: namespace-prefixed worksheets ----

        [Fact]
        public void PrefixedNumberAndInlineStringCellsAreRead()
        {
            using MemoryStream ms = WorkbookBuilder.BuildPrefixed("x",
                """<x:row r="1"><x:c r="A1"><x:v>42</x:v></x:c><x:c r="B1" t="inlineStr"><x:is><x:t>hello</x:t></x:is></x:c></x:row>""");
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Number, e.Current[0].Type);
            Assert.Equal("42", e.Current[0].GetString());
            Assert.Equal(CellType.ExcelString, e.Current[1].Type);
            Assert.Equal("hello", e.Current[1].GetString());
            Assert.False(e.MoveNext());
        }

        [Fact]
        public void PrefixedMultipleRowsAreAllEnumerated()
        {
            using MemoryStream ms = WorkbookBuilder.BuildPrefixed("x",
                """<x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row><x:row r="2"><x:c r="A2"><x:v>2</x:v></x:c></x:row><x:row r="3"><x:c r="A3"><x:v>3</x:v></x:c></x:row>""");
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal("1", e.Current[0].GetString());
            Assert.True(e.MoveNext());
            Assert.Equal("2", e.Current[0].GetString());
            Assert.True(e.MoveNext());
            Assert.Equal("3", e.Current[0].GetString());
            Assert.False(e.MoveNext());
        }

        [Fact]
        public void PrefixedSharedStringsAreResolved()
        {
            using MemoryStream ms = WorkbookBuilder.BuildPrefixed("x",
                """<x:row r="1"><x:c r="A1" t="s"><x:v>0</x:v></x:c><x:c r="B1" t="s"><x:v>1</x:v></x:c></x:row>""",
                sharedStrings: "<x:si><x:t>alpha</x:t></x:si><x:si><x:t>beta</x:t></x:si>");
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal("alpha", e.Current[0].GetString());
            Assert.Equal("beta", e.Current[1].GetString());
        }

        [Fact]
        public void PrefixedPhoneticRunsAreSkippedInSharedStrings()
        {
            // Exercises the prefixed <x:rPh> skip in WriteTextRuns (item 1.1 under a namespace prefix).
            using MemoryStream ms = WorkbookBuilder.BuildPrefixed("x",
                """<x:row r="1"><x:c r="A1" t="s"><x:v>0</x:v></x:c></x:row>""",
                sharedStrings: "<x:si><x:t>株式会社</x:t><x:rPh sb=\"0\" eb=\"4\"><x:t>カブシキガイシャ</x:t></x:rPh></x:si>");
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal("株式会社", e.Current[0].GetString());
        }

        [Fact]
        public void PrefixedCustomDateStyleIsClassifiedAsDate()
        {
            // Exercises the prefixed <x:numFmt> + <x:cellXfs>/<x:xf> style parsing.
            using MemoryStream ms = WorkbookBuilder.BuildPrefixed("x",
                """<x:row r="1"><x:c r="A1" s="0"><x:v>45658</x:v></x:c></x:row>""",
                stylesInner: """<x:numFmts count="1"><x:numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></x:numFmts><x:cellXfs count="1"><x:xf numFmtId="164"/></x:cellXfs>""");
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Date, e.Current[0].Type);
            Assert.True(e.Current[0].TryGetDateTime(reader.IsDate1904, out DateTime dt));
            Assert.Equal(new DateTime(2025, 1, 1), dt);
        }

        [Fact]
        public async Task PrefixedWorksheetIsReadAsync()
        {
            await using MemoryStream ms = WorkbookBuilder.BuildPrefixed("x",
                """<x:row r="1"><x:c r="A1"><x:v>7</x:v></x:c><x:c r="B1" t="s"><x:v>0</x:v></x:c></x:row>""",
                sharedStrings: "<x:si><x:t>async</x:t></x:si>");
            await using XlsxReader reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            await using XlsxReader.Enumerator e = await reader.GetAsyncEnumeratorAsync(TestContext.Current.CancellationToken);

            Assert.True(await e.MoveNextAsync());
            Assert.Equal("7", e.Current[0].GetString());
            Assert.Equal("async", e.Current[1].GetString());
            Assert.False(await e.MoveNextAsync());
        }

        [Fact]
        public void UnusualPrefixNameIsSupported()
        {
            // The prefix is detected from the root element, so it need not be the conventional "x".
            using MemoryStream ms = WorkbookBuilder.BuildPrefixed("ss",
                """<ss:row r="1"><ss:c r="A1"><ss:v>99</ss:v></ss:c></ss:row>""");
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal("99", e.Current[0].GetString());
        }

        // ---- 3.7: ISO-8601 date cells (t="d") ----

        [Fact]
        public void IsoDateTimeCellIsParsedAsDate()
        {
            using MemoryStream ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="d"><v>2026-01-02T13:45:00</v></c></row>""");
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Date, e.Current[0].Type);
            Assert.True(e.Current[0].TryGetDateTime(reader.IsDate1904, out DateTime dt));
            Assert.Equal(new DateTime(2026, 1, 2, 13, 45, 0), dt);
        }

        [Fact]
        public void IsoDateOnlyCellIsParsedAsDate()
        {
            using MemoryStream ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="d"><v>2026-01-02</v></c></row>""");
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Date, e.Current[0].Type);
            Assert.True(e.Current[0].TryGetDateTime(reader.IsDate1904, out DateTime dt));
            Assert.Equal(new DateTime(2026, 1, 2), dt);
        }

        [Fact]
        public void UnparseableIsoDateCellIsKeptAsString()
        {
            // Garbage in a t="d" cell must not crash the enumerator or silently vanish.
            using MemoryStream ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="d"><v>not-a-date</v></c></row>""");
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal(CellType.ExcelString, e.Current[0].Type);
            Assert.Equal("not-a-date", e.Current[0].GetString());
        }

        [Fact]
        public async Task IsoDateCellIsParsedAsDateAsync()
        {
            await using MemoryStream ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="d"><v>2026-01-02T00:00:00</v></c></row>""");
            await using XlsxReader reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            await using XlsxReader.Enumerator e = await reader.GetAsyncEnumeratorAsync(TestContext.Current.CancellationToken);

            Assert.True(await e.MoveNextAsync());
            Assert.Equal(CellType.Date, e.Current[0].Type);
            Assert.True(e.Current[0].TryGetDateTime(reader.IsDate1904, out DateTime dt));
            Assert.Equal(new DateTime(2026, 1, 2), dt);
        }

        [Fact]
        public void IsoDateCellCombinedWithPrefix()
        {
            // Both dialects at once: a prefixed worksheet whose cell is also an ISO t="d" date.
            using MemoryStream ms = WorkbookBuilder.BuildPrefixed("x",
                """<x:row r="1"><x:c r="A1" t="d"><x:v>2026-03-04</x:v></x:c></x:row>""");
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Date, e.Current[0].Type);
            Assert.True(e.Current[0].TryGetDateTime(reader.IsDate1904, out DateTime dt));
            Assert.Equal(new DateTime(2026, 3, 4), dt);
        }
    }
}
