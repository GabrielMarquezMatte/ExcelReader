using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests.Reader.Tables
{
    public class TableBindingTests
    {
        private sealed class SalesRow
        {
            public string? Product { get; set; }

            public int Qty { get; set; }

            public double Price { get; set; }
        }

        private sealed class StockRow
        {
            public string? Product { get; set; }

            public int Qty { get; set; }
        }

        [Theory]
        [MemberData(nameof(TableMetadataTests.Fixtures), MemberType = typeof(TableMetadataTests))]
        public void SalesBindsItsDataRowsWithoutTheTotalsRow(string file)
        {
            using IExcelWorkbook workbook = Excel.Open(TableMetadataTests.FixturePath(file));

            List<SalesRow> rows = ExcelParser.FromAttributes<SalesRow>().Parse(Table(workbook, "Sales").AsSheet()).ToList();

            Assert.Equal(
                [("Pen", 10, 2.5), ("Notebook", 4, 18.9), ("Pencil", 25, 1.2)],
                rows.Select(static r => (r.Product, r.Qty, r.Price)));
        }

        [Theory]
        [MemberData(nameof(TableMetadataTests.Fixtures), MemberType = typeof(TableMetadataTests))]
        public void StockBindsItsOwnColumnsDespiteSharedHeaderNames(string file)
        {
            using IExcelWorkbook workbook = Excel.Open(TableMetadataTests.FixturePath(file));

            List<StockRow> rows = ExcelParser.FromAttributes<StockRow>().Parse(Table(workbook, "Stock").AsSheet()).ToList();

            Assert.Equal([("Pen", 100), ("Notebook", 40), ("Pencil", 250)], rows.Select(static r => (r.Product, r.Qty)));
        }

        [Theory]
        [MemberData(nameof(TableMetadataTests.Fixtures), MemberType = typeof(TableMetadataTests))]
        public async Task StockBindsAsynchronouslyOverAStream(string file)
        {
            await using IExcelWorkbook workbook = await Excel.OpenAsync(
                File.OpenRead(TableMetadataTests.FixturePath(file)), leaveOpen: false, ct: TestContext.Current.CancellationToken);
            List<StockRow> rows = [];

            await foreach (StockRow row in ExcelParser.FromAttributes<StockRow>().Parse(Table(workbook, "Stock").AsSheet())
                .WithCancellation(TestContext.Current.CancellationToken))
            {
                rows.Add(row);
            }

            Assert.Equal([("Pen", 100), ("Notebook", 40), ("Pencil", 250)], rows.Select(static r => (r.Product, r.Qty)));
        }

        [Theory]
        [MemberData(nameof(TableMetadataTests.Fixtures), MemberType = typeof(TableMetadataTests))]
        public void HeaderlessTableRejectsBindingByName(string file)
        {
            using IExcelWorkbook workbook = Excel.Open(TableMetadataTests.FixturePath(file));
            IExcelSheet view = Table(workbook, "NoHeader").AsSheet();

            Assert.Throws<InvalidOperationException>(() => ExcelParser.FromAttributes<StockRow>().Parse(view));
            Assert.Throws<InvalidOperationException>(() => new ExcelDataReader(view));
            Assert.Throws<InvalidOperationException>(() => Excel.InferSchema(view));
        }

        [Theory]
        [MemberData(nameof(TableMetadataTests.Fixtures), MemberType = typeof(TableMetadataTests))]
        public void HeaderlessTableReadsWithoutAHeader(string file)
        {
            using IExcelWorkbook workbook = Excel.Open(TableMetadataTests.FixturePath(file));
            IExcelSheet view = Table(workbook, "NoHeader").AsSheet();

            using ExcelDataReader reader = new(view, headerRow: 0);
            List<string> letters = [];
            while (reader.Read())
            {
                letters.Add(reader.GetString(0));
            }

            Assert.Equal(["a", "b", "c"], letters);
            Assert.Equal(2, Excel.InferSchema(view, headerRow: 0).Length);
        }

        private static ExcelTable Table(IExcelWorkbook workbook, string name)
        {
            Assert.True(workbook.TryGetTable(name, out ExcelTable? table));
            return table;
        }
    }
}
