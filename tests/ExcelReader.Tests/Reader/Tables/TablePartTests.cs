using System.Text;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Tests.Reader.Tables
{
    public class TablePartTests
    {
        [Theory]
        [InlineData("C4:E8", 3, 2, 7, 4)]
        [InlineData("A1", 0, 0, 0, 0)]
        [InlineData("AA10:AB12", 9, 26, 11, 27)]
        [InlineData("XFD1048576:XFD1048576", 1048575, 16383, 1048575, 16383)]
        public void TryParseRangeReadsA1Ranges(string text, int firstRow, int firstColumn, int lastRow, int lastColumn)
        {
            Assert.True(TablePart.TryParseRange(Encoding.ASCII.GetBytes(text), out int r1, out int c1, out int r2, out int c2));
            Assert.Equal((firstRow, firstColumn, lastRow, lastColumn), (r1, c1, r2, c2));
        }

        [Theory]
        [InlineData("")]
        [InlineData("C")]
        [InlineData("4")]
        [InlineData("C0")]
        [InlineData("c4")]
        [InlineData("C4:")]
        [InlineData("C4:E")]
        [InlineData("$C$4:$E$8")]
        [InlineData("XFE1")]
        [InlineData("AAAA1")]
        [InlineData("C4x")]
        [InlineData("C99999999999")]
        public void TryParseRangeRejectsMalformedText(string text)
        {
            Assert.False(TablePart.TryParseRange(Encoding.ASCII.GetBytes(text), out _, out _, out _, out _));
        }

        [Fact]
        public void FormatRangeWritesA1Notation()
        {
            Assert.Equal("C4:E8", TablePart.FormatRange(3, 2, 7, 4));
            Assert.Equal("XFD1048576:XFD1048576", TablePart.FormatRange(1048575, 16383, 1048575, 16383));
        }

        [Fact]
        public void CreateKeepsAWellFormedTable()
        {
            TablePart part = TablePart.Create(1, 7, "Sales", 3, 2, 7, 4, 1, 1, ["Product", "Qty", "Price"], "TableStyleMedium2");

            Assert.Equal(("Sales", 1, 7, "TableStyleMedium2"), (part.Name, part.SheetIndex, part.Id, part.StyleName));
            Assert.Equal((3, 2, 7, 4, 1, 1), (part.FirstRow, part.FirstColumn, part.LastRow, part.LastColumn, part.HeaderRowCount, part.TotalsRowCount));
            Assert.Equal(["Product", "Qty", "Price"], part.Columns);
        }

        [Theory]
        [InlineData("", 0, 0, 1, 0, 1, 0, 1)]
        [InlineData("T", -1, 0, 1, 0, 1, 0, 1)]
        [InlineData("T", 2, 0, 1, 0, 1, 0, 1)]
        [InlineData("T", 0, 0, 1048576, 0, 1, 0, 1)]
        [InlineData("T", 0, 1, 1, 0, 1, 0, 1)]
        [InlineData("T", 0, 0, 1, 16384, 1, 0, 16385)]
        [InlineData("T", 0, 0, 1, 0, 2, 0, 1)]
        [InlineData("T", 0, 0, 1, 0, 1, -1, 1)]
        [InlineData("T", 0, 0, 0, 0, 1, 1, 1)]
        [InlineData("T", 0, 0, 1, 1, 1, 0, 1)]
        public void CreateRejectsInconsistentTables(string name, int firstRow, int firstColumn, int lastRow, int lastColumn,
            int headerRowCount, int totalsRowCount, int columnCount)
        {
            string[] columns = [.. Enumerable.Range(0, columnCount).Select(static i => $"C{i}")];

            Assert.Throws<InvalidDataException>(() => TablePart.Create(
                0, 1, name, firstRow, firstColumn, lastRow, lastColumn, headerRowCount, totalsRowCount, columns, null));
        }
    }
}
