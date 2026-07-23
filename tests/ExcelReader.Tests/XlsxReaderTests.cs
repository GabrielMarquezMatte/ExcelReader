using System.Globalization;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    // Focused reader suite for XlsxReader itself — the flagship format's coverage was previously
    // scattered across dialect/corpus/interop test files with no single place asserting the reader's
    // own structural behaviors (F13 in docs/road-to-a.md).
    public class XlsxReaderTests
    {
        [Fact]
        public void MultiSheetSelectionByIndexReadsDistinctData()
        {
            using MemoryStream ms = WorkbookBuilder.BuildMultiSheet(
            [
                ("First", """<row r="1"><c r="A1"><v>1</v></c></row>"""),
                ("Second", """<row r="1"><c r="A1"><v>2</v></c></row>"""),
                ("Third", """<row r="1"><c r="A1"><v>3</v></c></row>"""),
            ]);
            using XlsxReader reader = Excel.From(ms);

            Assert.Equal(3, reader.SheetCount);
            for (int i = 0; i < reader.SheetCount; i++)
            {
                reader.MoveToSheet(i);
                using XlsxReader.Enumerator e = reader.GetEnumerator();
                Assert.True(e.MoveNext());
                Assert.Equal((i + 1).ToString(CultureInfo.InvariantCulture), e.Current[0].GetString());
            }
        }

        [Fact]
        public void MultiSheetSelectionByNameIsCaseInsensitive()
        {
            using MemoryStream ms = WorkbookBuilder.BuildMultiSheet(
            [
                ("Alpha", """<row r="1"><c r="A1"><v>10</v></c></row>"""),
                ("Beta", """<row r="1"><c r="A1"><v>20</v></c></row>"""),
            ]);
            using XlsxReader reader = Excel.From(ms);

            Assert.True(reader.TryMoveToSheet("beta"));
            Assert.Equal("Beta", reader.SheetName);
            using XlsxReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("20", e.Current[0].GetString());
        }

        [Fact]
        public void SameSheetCanBeEnumeratedMoreThanOnce()
        {
            using MemoryStream ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>""");
            using XlsxReader reader = Excel.From(ms);

            using (XlsxReader.Enumerator first = reader.GetEnumerator())
            {
                Assert.True(first.MoveNext());
                Assert.Equal("1", first.Current[0].GetString());
                Assert.True(first.MoveNext());
                Assert.Equal("2", first.Current[0].GetString());
                Assert.False(first.MoveNext());
            }

            // A fresh GetEnumerator() call re-opens the worksheet entry from the start.
            using XlsxReader.Enumerator second = reader.GetEnumerator();
            Assert.True(second.MoveNext());
            Assert.Equal("1", second.Current[0].GetString());
        }

        [Fact]
        public void SwitchingAwayAndBackToASheetReenumeratesFromStart()
        {
            using MemoryStream ms = WorkbookBuilder.BuildMultiSheet(
            [
                ("A", """<row r="1"><c r="A1"><v>11</v></c></row>"""),
                ("B", """<row r="1"><c r="A1"><v>22</v></c></row>"""),
            ]);
            using XlsxReader reader = Excel.From(ms);

            using (XlsxReader.Enumerator e = reader.GetEnumerator())
            {
                Assert.True(e.MoveNext());
                Assert.Equal("11", e.Current[0].GetString());
            }

            reader.MoveToSheet(1);
            using (XlsxReader.Enumerator e = reader.GetEnumerator())
            {
                Assert.True(e.MoveNext());
                Assert.Equal("22", e.Current[0].GetString());
            }

            reader.MoveToSheet(0);
            using XlsxReader.Enumerator back = reader.GetEnumerator();
            Assert.True(back.MoveNext());
            Assert.Equal("11", back.Current[0].GetString());
        }

        [Fact]
        public void EmptySheetYieldsNoRows()
        {
            using MemoryStream ms = WorkbookBuilder.Build(sheetRows: "");
            using XlsxReader reader = Excel.From(ms);

            using XlsxReader.Enumerator e = reader.GetEnumerator();
            Assert.False(e.MoveNext());
        }

        [Fact]
        public void HugeSharedStringTableResolvesEveryIndexIncludingRepeatsViaCache()
        {
            const int uniqueCount = 5_000;
            var sharedStrings = new System.Text.StringBuilder();
            var rows = new System.Text.StringBuilder();
            for (int i = 0; i < uniqueCount; i++)
            {
                sharedStrings.Append(CultureInfo.InvariantCulture, $"<si><t>value{i}</t></si>");
            }
            // Every row references the same repeated index (exercises the shared-string dedup cache)
            // plus a unique index per row (exercises the offset table across a large uniqueCount).
            for (int i = 0; i < 200; i++)
            {
                rows.Append(CultureInfo.InvariantCulture,
                    $"""<row r="{i + 1}"><c r="A{i + 1}" t="s"><v>{i}</v></c><c r="B{i + 1}" t="s"><v>0</v></c></row>""");
            }

            using MemoryStream ms = WorkbookBuilder.Build(rows.ToString(), sharedStrings: sharedStrings.ToString());
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            for (int i = 0; i < 200; i++)
            {
                Assert.True(e.MoveNext());
                Assert.Equal($"value{i}", e.Current[0].GetString());
                Assert.Equal("value0", e.Current[1].GetString());
            }
            Assert.False(e.MoveNext());
        }

        [Fact]
        public void NumericCellIsTypedAsNumber()
        {
            using MemoryStream ms = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>3.14</v></c></row>""");
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Number, e.Current[0].Type);
            Assert.True(e.Current[0].TryGetDouble(out double value));
            Assert.Equal(3.14, value);
        }

        [Fact]
        public void DateStyledNumericCellIsTypedAsDate()
        {
            using MemoryStream ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" s="0"><v>45658</v></c></row>""",
                styles: "<styleSheet><cellXfs count=\"1\"><xf numFmtId=\"14\"/></cellXfs></styleSheet>");
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Date, e.Current[0].Type);
            Assert.True(e.Current[0].TryGetDateTime(reader.IsDate1904, out DateTime dt));
            Assert.Equal(new DateTime(2025, 1, 1), dt);
        }

        [Fact]
        public void IsDate1904DefaultsToFalse()
        {
            using MemoryStream ms = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>1</v></c></row>""");
            using XlsxReader reader = Excel.From(ms);

            Assert.False(reader.IsDate1904);
        }

        [Fact]
        public void IsDate1904IsTrueWhenWorkbookPrDeclaresIt()
        {
            using MemoryStream ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1"><v>1</v></c></row>""", date1904: true);
            using XlsxReader reader = Excel.From(ms);

            Assert.True(reader.IsDate1904);
        }
    }
}
