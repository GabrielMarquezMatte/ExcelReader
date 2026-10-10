using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests.Writer.Tables
{
    public class TableWriterTests
    {
        private sealed class StockRow
        {
            public string? Product { get; set; }

            public int Qty { get; set; }
        }

        [Theory]
        [MemberData(nameof(TableBooks.Formats), MemberType = typeof(TableBooks))]
        public void SingleTableRoundTrips(string format)
        {
            ITableBook book = TableBooks.Create(format);
            book.AddSheet("Report");
            book.BeginTable("Stock", ["Product", "Qty"]);
            book.Row("Pen", 100);
            book.Row("Notebook", 40);
            book.EndTable();

            using IExcelWorkbook workbook = Read(book.Finish());

            ExcelTable table = Assert.Single(workbook.Tables);
            Assert.Equal(("Stock", "A1:B3", "TableStyleMedium2"), (table.Name, table.Ref, table.StyleName));
            Assert.Equal((1, 0, "Report"), (table.HeaderRowCount, table.TotalsRowCount, table.Sheet.Name));
            Assert.Equal(["Product", "Qty"], table.ColumnNames);
            Assert.Equal([("Pen", 100), ("Notebook", 40)], ParseStock(table));
        }

        [Theory]
        [MemberData(nameof(TableBooks.Formats), MemberType = typeof(TableBooks))]
        public void TableAtAnOffsetColumnRoundTrips(string format)
        {
            ITableBook book = TableBooks.Create(format);
            book.AddSheet("Report");
            book.BeginTable("Stock", ["Product", "Qty"], ExcelTableOptions.Default with { FirstColumn = 2 });
            book.Row(null, null, "Pen", 100);
            book.Row(null, null, "Notebook", 40);

            using IExcelWorkbook workbook = Read(book.Finish());

            ExcelTable table = Assert.Single(workbook.Tables);
            Assert.Equal("C1:D3", table.Ref);
            Assert.Equal([("Pen", 100), ("Notebook", 40)], ParseStock(table));
        }

        [Theory]
        [MemberData(nameof(TableBooks.Formats), MemberType = typeof(TableBooks))]
        public void TablesInSequenceAndAcrossSheetsGetWorkbookWideIds(string format)
        {
            ITableBook book = TableBooks.Create(format);
            book.AddSheet("One");
            book.BeginTable("First", ["A"]);
            book.Row("x");
            book.EndTable();
            book.Row("between");
            book.BeginTable("Second", ["B"]);
            book.Row("y");
            book.EndTable();
            book.EndSheet();
            book.AddSheet("Two");
            book.BeginTable("Third", ["C"]);
            book.Row("z");
            byte[] package = book.Finish();

            using IExcelWorkbook workbook = Read(package);

            Assert.Equal(
                [("First", 0, "A1:A2"), ("Second", 0, "A4:A5"), ("Third", 1, "A1:A2")],
                workbook.Tables.Select(static t => (t.Name, t.Sheet.Index, t.Ref)));
            string ext = TableBooks.Extension(format);
            IReadOnlyList<string> names = TableBooks.EntryNames(package);
            Assert.Contains($"xl/tables/table1.{ext}", names, StringComparer.Ordinal);
            Assert.Contains($"xl/tables/table2.{ext}", names, StringComparer.Ordinal);
            Assert.Contains($"xl/tables/table3.{ext}", names, StringComparer.Ordinal);
        }

        [Theory]
        [MemberData(nameof(TableBooks.Formats), MemberType = typeof(TableBooks))]
        public void HeaderOnlyTableGetsOneEmptyDataRow(string format)
        {
            ITableBook book = TableBooks.Create(format);
            book.AddSheet("Report");
            book.BeginTable("Empty", ["Only"]);
            book.EndTable();

            using IExcelWorkbook workbook = Read(book.Finish());

            Assert.Equal("A1:A2", Assert.Single(workbook.Tables).Ref);
        }

        [Theory]
        [MemberData(nameof(TableBooks.Formats), MemberType = typeof(TableBooks))]
        public void SheetEndClosesTheOpenTable(string format)
        {
            ITableBook book = TableBooks.Create(format);
            book.AddSheet("Report");
            book.BeginTable("Open", ["H"]);
            book.Row("a");
            book.Row("b");

            using IExcelWorkbook workbook = Read(book.Finish());

            Assert.Equal("A1:A3", Assert.Single(workbook.Tables).Ref);
        }

        [Theory]
        [MemberData(nameof(TableBooks.Formats), MemberType = typeof(TableBooks))]
        public void UnstyledTableHasNoStyleName(string format)
        {
            ITableBook book = TableBooks.Create(format);
            book.AddSheet("Report");
            book.BeginTable("Plain", ["H"], ExcelTableOptions.Default with { StyleName = null });
            book.Row("a");

            using IExcelWorkbook workbook = Read(book.Finish());

            Assert.Null(Assert.Single(workbook.Tables).StyleName);
        }

        [Theory]
        [MemberData(nameof(TableBooks.Formats), MemberType = typeof(TableBooks))]
        public void FailedBeginTableWritesNothingAndClaimsNoName(string format)
        {
            ITableBook book = TableBooks.Create(format);
            book.AddSheet("Report");

            Assert.Throws<ArgumentException>(() => book.BeginTable("A1", ["H"]));
            Assert.Throws<ArgumentException>(() => book.BeginTable("Good", ["H", "h"]));
            book.BeginTable("Good", ["H"]);
            book.Row("a");

            using IExcelWorkbook workbook = Read(book.Finish());

            Assert.Equal("A1:A2", Assert.Single(workbook.Tables).Ref);
        }

        [Theory]
        [MemberData(nameof(TableBooks.Formats), MemberType = typeof(TableBooks))]
        public void DuplicateNameAcrossSheetsIsRejectedBeforeWriting(string format)
        {
            ITableBook book = TableBooks.Create(format);
            book.AddSheet("One");
            book.BeginTable("Sales", ["H"]);
            book.Row("a");
            book.EndSheet();
            book.AddSheet("Two");

            Assert.Throws<ArgumentException>(() => book.BeginTable("sales", ["H"]));
            book.Row("untouched");

            using IExcelWorkbook workbook = Read(book.Finish());

            Assert.Single(workbook.Tables);
            using IExcelRowEnumerator rows = workbook.SheetAt(1).GetEnumerator();
            Assert.True(rows.MoveNext());
            Assert.Equal("untouched", rows.Current[0].GetString());
        }

        [Theory]
        [MemberData(nameof(TableBooks.Formats), MemberType = typeof(TableBooks))]
        public void SecondOpenTableAndEndWithoutOpenAreRejected(string format)
        {
            ITableBook book = TableBooks.Create(format);
            book.AddSheet("Report");

            Assert.Throws<InvalidOperationException>(() => book.EndTable());
            book.BeginTable("First", ["H"]);
            Assert.Throws<InvalidOperationException>(() => book.BeginTable("Second", ["H"]));
            book.Finish();
        }

        [Theory]
        [MemberData(nameof(TableBooks.Formats), MemberType = typeof(TableBooks))]
        public async Task AsyncBeginTableMatchesSync(string format)
        {
            ITableBook sync = TableBooks.Create(format);
            sync.AddSheet("Report");
            sync.BeginTable("Stock", ["Product", "Qty"], ExcelTableOptions.Default with { FirstColumn = 1 });
            sync.Row(null, "Pen", 100);

            ITableBook async = TableBooks.Create(format);
            async.AddSheet("Report");
            await async.BeginTableAsync("Stock", ["Product", "Qty"], ExcelTableOptions.Default with { FirstColumn = 1 });
            async.Row(null, "Pen", 100);

            Dictionary<string, byte[]> expected = TableBooks.Entries(sync.Finish());
            Dictionary<string, byte[]> actual = TableBooks.Entries(async.Finish());
            Assert.Equal(expected.Keys.Order(StringComparer.Ordinal), actual.Keys.Order(StringComparer.Ordinal));
            foreach ((string name, byte[] bytes) in expected)
            {
                Assert.Equal(bytes, actual[name]);
            }
        }

        private static IExcelWorkbook Read(byte[] package)
        {
            return Excel.Open(new ReadOnlyMemory<byte>(package));
        }

        private static List<(string?, int)> ParseStock(ExcelTable table)
        {
            List<(string?, int)> rows = [];
            foreach (StockRow row in ExcelParser.FromAttributes<StockRow>().Parse(table.AsSheet()))
            {
                rows.Add((row.Product, row.Qty));
            }
            return rows;
        }
    }
}
