using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Tests.Writer.Tables
{
    public class TableValidationTests
    {
        [Theory]
        [InlineData("Sales")]
        [InlineData("_private")]
        [InlineData("\\back")]
        [InlineData("Q1.Results_2026")]
        [InlineData("Vendas")]
        [InlineData("ABCD1")]
        [InlineData("Ação")]
        public void ValidNamesPass(string name)
        {
            string[] columns = TableValidation.Validate(name, ["A"], ExcelTableOptions.Default);

            Assert.Equal(["A"], columns);
        }

        [Theory]
        [InlineData("")]
        [InlineData("1abc")]
        [InlineData(".dot")]
        [InlineData("has space")]
        [InlineData("dash-name")]
        [InlineData("A1")]
        [InlineData("abc1")]
        [InlineData("XFD1048576")]
        [InlineData("R1C1")]
        [InlineData("r")]
        [InlineData("C")]
        [InlineData("RC")]
        [InlineData("R12")]
        [InlineData("c3")]
        public void InvalidNamesThrow(string name)
        {
            Assert.Throws<ArgumentException>(() => TableValidation.Validate(name, ["A"], ExcelTableOptions.Default));
        }

        [Fact]
        public void NullAndOverlongNamesThrow()
        {
            Assert.Throws<ArgumentNullException>(() => TableValidation.Validate(null!, ["A"], ExcelTableOptions.Default));
            Assert.Throws<ArgumentException>(() => TableValidation.Validate(new string('a', 256), ["A"], ExcelTableOptions.Default));
            Assert.Equal(["A"], TableValidation.Validate(new string('a', 255), ["A"], ExcelTableOptions.Default));
        }

        [Fact]
        public void InvalidColumnsThrow()
        {
            Assert.Throws<ArgumentNullException>(() => TableValidation.Validate("T", null!, ExcelTableOptions.Default));
            Assert.Throws<ArgumentException>(() => TableValidation.Validate("T", [], ExcelTableOptions.Default));
            Assert.Throws<ArgumentException>(() => TableValidation.Validate("T", ["a", ""], ExcelTableOptions.Default));
            Assert.Throws<ArgumentException>(() => TableValidation.Validate("T", ["a", null!], ExcelTableOptions.Default));
            Assert.Throws<ArgumentException>(() => TableValidation.Validate("T", ["Qty", "qty"], ExcelTableOptions.Default));
            Assert.Throws<ArgumentException>(() => TableValidation.Validate("T", ["a\nb"], ExcelTableOptions.Default));
            Assert.Throws<ArgumentException>(() => TableValidation.Validate("T", [new string('c', 256)], ExcelTableOptions.Default));
        }

        [Fact]
        public void ColumnsMustFitTheSheet()
        {
            string[] two = ["a", "b"];

            Assert.Equal(two, TableValidation.Validate("T", two, ExcelTableOptions.Default with { FirstColumn = 16382 }));
            Assert.Throws<ArgumentException>(() => TableValidation.Validate("T", two, ExcelTableOptions.Default with { FirstColumn = 16383 }));
            Assert.Throws<ArgumentOutOfRangeException>(() => TableValidation.Validate("T", two, ExcelTableOptions.Default with { FirstColumn = -1 }));
        }

        [Theory]
        [InlineData("TableStyleLight1")]
        [InlineData("TableStyleLight21")]
        [InlineData("TableStyleMedium28")]
        [InlineData("TableStyleDark11")]
        [InlineData(null)]
        public void BuiltInStylesPass(string? style)
        {
            Assert.Equal(["A"], TableValidation.Validate("T", ["A"], ExcelTableOptions.Default with { StyleName = style }));
        }

        [Theory]
        [InlineData("TableStyleLight0")]
        [InlineData("TableStyleLight22")]
        [InlineData("TableStyleMedium29")]
        [InlineData("TableStyleDark12")]
        [InlineData("TableStyleLight01")]
        [InlineData("tablestylelight1")]
        [InlineData("MyStyle")]
        [InlineData("")]
        public void UnknownStylesThrow(string style)
        {
            Assert.Throws<ArgumentException>(() => TableValidation.Validate("T", ["A"], ExcelTableOptions.Default with { StyleName = style }));
        }

        [Fact]
        public void ValidateReturnsACopyOfTheColumns()
        {
            List<string> source = ["a", "b"];
            string[] copy = TableValidation.Validate("T", source, ExcelTableOptions.Default);
            source[0] = "changed";

            Assert.Equal(["a", "b"], copy);
        }

        [Fact]
        public void RegistryClaimsNamesCaseInsensitivelyAndNumbersTables()
        {
            TableRegistry registry = new();

            Assert.Equal(1, registry.Claim("Sales"));
            Assert.Equal(2, registry.Claim("Stock"));
            Assert.Throws<ArgumentException>(() => registry.RequireAvailable("sales"));
            Assert.Throws<ArgumentException>(() => registry.Claim("STOCK"));
            Assert.Equal(3, registry.Claim("Other"));
        }

        [Fact]
        public void TrackerGivesHeaderOnlyTablesOneDataRow()
        {
            TableTracker tracker = new();
            tracker.Begin(1, "T", ["A"], ExcelTableOptions.Default, headerRow: 4);
            tracker.End(lastWrittenRow: 4);

            WrittenTable table = Assert.Single(tracker.Finished);
            Assert.Equal((4, 5, "A5:A6"), (table.HeaderRow, table.LastRow, table.Ref));
        }

        [Fact]
        public void TrackerRefusesTwoOpenTablesAndEndWithoutOpen()
        {
            TableTracker tracker = new();
            Assert.Throws<InvalidOperationException>(() => tracker.End(0));

            tracker.Begin(1, "T", ["A", "B"], ExcelTableOptions.Default with { FirstColumn = 2 }, headerRow: 0);
            Assert.Throws<InvalidOperationException>(() => tracker.RequireNoneOpen());
            tracker.CloseOpen(lastWrittenRow: 3);
            tracker.CloseOpen(lastWrittenRow: 9);

            WrittenTable table = Assert.Single(tracker.Finished);
            Assert.Equal((2, 3, "C1:D4"), (table.FirstColumn, table.LastColumn, table.Ref));
        }

        [Fact]
        public void TableNeedsRoomForItsHeaderAndOneDataRow()
        {
            TableTracker.RequireRoom(1_048_574);

            Assert.Throws<ExcelLimitExceededException>(() => TableTracker.RequireRoom(1_048_575));
        }

        [Fact]
        public void SheetRelsListOneRelationshipPerTable()
        {
            WrittenTable first = new(3, "A", ["x"], ExcelTableOptions.Default, 0, 1);
            WrittenTable second = new(4, "B", ["y"], ExcelTableOptions.Default, 3, 4);

            string xml = TablePackage.SheetRelsXml([first, second], "bin");

            Assert.Equal(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/table3.bin\"/>"
                + "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/table4.bin\"/>"
                + "</Relationships>",
                xml);
        }
    }
}
