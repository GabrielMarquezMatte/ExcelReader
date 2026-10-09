using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using ExcelReader.Core.Reader.Xls;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Tests.Reader.Xls;

namespace ExcelReader.Tests.Reader.Tables
{
    public class WorkbookTablesTests
    {
        [Fact]
        public void XlsxWithoutTablesReportsNone()
        {
            using MemoryStream ms = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>1</v></c></row>""");
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);

            Assert.Empty(workbook.Tables);
            Assert.False(workbook.TryGetTable("Sales", out _));
        }

        [Fact]
        public void XlsReportsNoTables()
        {
            using MemoryStream ms = XlsWorkbookBuilder.Build(false, null, null, ("S1", [["a"]]));
            using XlsWorkbook workbook = Excel.FromXls(ms);

            Assert.Empty(workbook.Tables);
            Assert.False(workbook.TryGetTable("Sales", out _));
        }

        [Fact]
        public void CsvReportsNoTables()
        {
            using CsvReader workbook = Excel.FromCsv("a,b\n1,2\n"u8.ToArray());

            Assert.Empty(workbook.Tables);
            Assert.False(workbook.TryGetTable("Sales", out _));
        }

        [Fact]
        public void TablesThrowOnceTheWorkbookIsDisposed()
        {
            using MemoryStream ms = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>1</v></c></row>""");
            XlsxWorkbook workbook = Excel.FromXlsx(ms);
            workbook.Dispose();

            Assert.Throws<ObjectDisposedException>(() => workbook.Tables);
            Assert.Throws<ObjectDisposedException>(() => workbook.TryGetTable("Sales", out _));
        }
    }
}
