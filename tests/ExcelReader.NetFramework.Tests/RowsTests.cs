using ExcelReader.NetFramework.Interop;

namespace ExcelReader.NetFramework.Tests
{
    public sealed class RowsTests
    {
        private const int RealRowCount = 101;
        private static readonly string[] SalesHeader = { "Product", "Qty", "Price" };
        private static readonly int[] SalesColumns = { 2, 3, 4 };

        [Theory]
        [InlineData("RealExcel.xlsx")]
        [InlineData("RealExcel.xlsb")]
        public void ReadsHeaderAndDataRows(string file)
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path(file));
            List<ExcelRow> rows = workbook.Sheet(0).ReadRows().ToList();
            Assert.Equal(RealRowCount, rows.Count);
            Assert.Contains(rows[0].Cells, c => string.Equals(c.Text, "Coluna1", StringComparison.Ordinal) && c.Type == ExcelCellType.String);
            Assert.Contains(rows[1].Cells, c => string.Equals(c.Text, "Valor1", StringComparison.Ordinal));
        }

        [Theory]
        [InlineData("tables.xlsx")]
        [InlineData("tables.xlsb")]
        public void TableRowsStartAtHeader(string file)
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("tables", file));
            ExcelSheet table = workbook.Table("Sales");
            ExcelRow header = table.ReadRows().First();
            Assert.Equal(SalesHeader, header.Cells.Select(c => c.Text).ToArray());
            Assert.Equal(SalesColumns, header.Cells.Select(c => c.Column).ToArray());
            Assert.Equal(4, table.ReadRows().Count());
        }

        [Fact]
        public void HeaderReaderTrimsAndIgnoresCase()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            Dictionary<string, int> header = HeaderReader.Read(workbook.Sheet(0), 1);
            Assert.True(header.ContainsKey("coluna1"));
            Assert.Equal(header["Coluna1"], header["coluna1"]);
        }

        [Fact]
        public void HeaderReaderThrowsWhenSheetIsShorterThanHeaderRow()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            ExcelReaderException error = Assert.Throws<ExcelReaderException>(() => HeaderReader.Read(workbook.Sheet(0), 1000));
            Assert.Equal(ExcelStatus.InvalidArgument, error.Status);
        }

        [Fact]
        public void BreakingOutReleasesTheCursor()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            ExcelSheet sheet = workbook.Sheet(0);
            for (int i = 0; i < 1000; i++)
            {
                foreach (ExcelRow row in sheet.ReadRows())
                {
                    break;
                }
            }

            Assert.Equal(RealRowCount, sheet.ReadRows().Count());
        }

        [Fact]
        public void EnumerationSurvivesWorkbookDispose()
        {
            ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            using IEnumerator<ExcelRow> rows = workbook.Sheet(0).ReadRows().GetEnumerator();
            Assert.True(rows.MoveNext());
            workbook.Dispose();
            int total = 1;
            while (rows.MoveNext())
            {
                total++;
            }

            Assert.Equal(RealRowCount, total);
        }

        [Fact]
        public void ConcurrentEnumerationsOnOneWorkbook()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            ExcelSheet sheet = workbook.Sheet(0);
            Parallel.For(0, 8, _ => Assert.Equal(RealRowCount, sheet.ReadRows().Count()));
        }
    }
}
