using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class MultiSheetTests
    {
        [Fact]
        public void SheetCountMatchesWorkbookSheetCount()
        {
            using var ms = WorkbookBuilder.BuildMultiSheet(
                [("Alpha", """<row r="1"><c r="A1"><v>1</v></c></row>"""),
                 ("Beta",  """<row r="1"><c r="A1"><v>2</v></c></row>""")]);
            using var reader = Excel.From(ms);
            Assert.Equal(2, reader.SheetCount);
        }

        [Fact]
        public void SheetNameMatchesCurrentSheet()
        {
            using var ms = WorkbookBuilder.BuildMultiSheet([("MySheet", "")]);
            using var reader = Excel.From(ms);
            Assert.Equal("MySheet", reader.SheetName);
        }

        [Fact]
        public void TryMoveToSheetMatchesCaseInsensitively()
        {
            using var ms = WorkbookBuilder.BuildMultiSheet([("Sheet1", "")]);
            using var reader = Excel.From(ms);
            Assert.True(reader.TryMoveToSheet("sheet1"));
            Assert.Equal("Sheet1", reader.SheetName);
        }

        [Fact]
        public void TryMoveToSheetReturnsFalseWhenNotFound()
        {
            using var ms = WorkbookBuilder.Build("");
            using var reader = Excel.From(ms);
            Assert.False(reader.TryMoveToSheet("DoesNotExist"));
        }

        [Fact]
        public void MoveToSheetByIndexSwitchesCurrentSheet()
        {
            using var ms = WorkbookBuilder.BuildMultiSheet(
                [("First", ""), ("Second", "")]);
            using var reader = Excel.From(ms);
            reader.MoveToSheet(1);
            Assert.Equal("Second", reader.SheetName);
        }

        [Fact]
        public void MoveToSheetNegativeIndexThrows()
        {
            using var ms = WorkbookBuilder.Build("");
            using var reader = Excel.From(ms);
            Assert.Throws<ArgumentOutOfRangeException>(() => reader.MoveToSheet(-1));
        }

        [Fact]
        public void MoveToSheetOutOfRangeIndexThrows()
        {
            using var ms = WorkbookBuilder.Build("");
            using var reader = Excel.From(ms);
            Assert.Throws<ArgumentOutOfRangeException>(() => reader.MoveToSheet(1));
        }

        [Fact]
        public void MultipleSheetsEachHaveDistinctData()
        {
            using var ms = WorkbookBuilder.BuildMultiSheet(
                [("A", """<row r="1"><c r="A1"><v>11</v></c></row>"""),
                 ("B", """<row r="1"><c r="A1"><v>22</v></c></row>""")]);
            using var reader = Excel.From(ms);

            using var e1 = reader.GetEnumerator();
            Assert.True(e1.MoveNext());
            Assert.True(e1.Current[0].TryParse(null, out int v1));
            Assert.Equal(11, v1);

            reader.MoveToSheet(1);

            using var e2 = reader.GetEnumerator();
            Assert.True(e2.MoveNext());
            Assert.True(e2.Current[0].TryParse(null, out int v2));
            Assert.Equal(22, v2);
        }
    }
}
