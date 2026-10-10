using System.Text;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests.Writer.Tables
{
    public class XlsxTableShapeTests
    {
        [Fact]
        public void SheetReferencesItsTablesAndDeclaresTheRelationshipsNamespace()
        {
            byte[] package = TwoTables();

            string sheet = Text(package, "xl/worksheets/sheet1.xml");

            Assert.StartsWith(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">",
                sheet, StringComparison.Ordinal);
            Assert.EndsWith(
                "</sheetData><tableParts count=\"2\"><tablePart r:id=\"rId1\"/><tablePart r:id=\"rId2\"/></tableParts></worksheet>",
                sheet, StringComparison.Ordinal);
        }

        [Fact]
        public void SheetRelsAndContentTypesListTheTables()
        {
            byte[] package = TwoTables();

            Assert.Equal(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/table1.xml\"/>"
                + "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/table2.xml\"/>"
                + "</Relationships>",
                Text(package, "xl/worksheets/_rels/sheet1.xml.rels"));
            string types = Text(package, "[Content_Types].xml");
            Assert.Contains("<Override PartName=\"/xl/tables/table1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>", types, StringComparison.Ordinal);
            Assert.Contains("<Override PartName=\"/xl/tables/table2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>", types, StringComparison.Ordinal);
        }

        [Fact]
        public void TablePartCarriesNameRangeColumnsAndStyleFlags()
        {
            byte[] package = TwoTables();

            Assert.Equal(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"2\" name=\"Prices\" displayName=\"Prices\" ref=\"B5:C6\" totalsRowShown=\"0\">"
                + "<autoFilter ref=\"B5:C6\"/>"
                + "<tableColumns count=\"2\"><tableColumn id=\"1\" name=\"Item\"/><tableColumn id=\"2\" name=\"Qty &amp; &quot;Price&quot;\"/></tableColumns>"
                + "<tableStyleInfo name=\"TableStyleLight9\" showFirstColumn=\"1\" showLastColumn=\"0\" showRowStripes=\"0\" showColumnStripes=\"1\"/>"
                + "</table>",
                Text(package, "xl/tables/table2.xml"));
        }

        [Fact]
        public void UnstyledTableOmitsTheStyleName()
        {
            ITableBook book = TableBooks.Create("xlsx");
            book.AddSheet("S");
            book.BeginTable("Plain", ["H"], ExcelTableOptions.Default with { StyleName = null });

            string table = Text(book.Finish(), "xl/tables/table1.xml");

            Assert.Contains("<tableStyleInfo showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\"/>", table, StringComparison.Ordinal);
        }

        [Fact]
        public void ColumnNameWithAnEscapeLookalikeIsEscapedLikeItsHeaderCell()
        {
            ITableBook book = TableBooks.Create("xlsx");
            book.AddSheet("S");
            book.BeginTable("Codes", ["Code_x0041_"]);
            book.Row("a");
            byte[] package = book.Finish();

            Assert.Contains("<tableColumn id=\"1\" name=\"Code_x005F_x0041_\"/>", Text(package, "xl/tables/table1.xml"), StringComparison.Ordinal);
        }

        private static byte[] TwoTables()
        {
            ITableBook book = TableBooks.Create("xlsx");
            book.AddSheet("S");
            book.BeginTable("Stock", ["Product"]);
            book.Row("Pen");
            book.EndTable();
            book.Row();
            book.Row();
            book.BeginTable("Prices", ["Item", "Qty & \"Price\""], new ExcelTableOptions
            {
                FirstColumn = 1,
                StyleName = "TableStyleLight9",
                ShowRowStripes = false,
                ShowColumnStripes = true,
                ShowFirstColumn = true,
            });
            book.Row(null, "Pen", 2);
            return book.Finish();
        }

        private static string Text(byte[] package, string name)
        {
            return Encoding.UTF8.GetString(TableBooks.Entry(package, name));
        }
    }
}
