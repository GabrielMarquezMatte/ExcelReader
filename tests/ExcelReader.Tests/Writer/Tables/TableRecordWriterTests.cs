using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Writer.Tables
{
    public class TableRecordWriterTests
    {
        public sealed class Sale
        {
            public string? Product { get; set; }

            public int Qty { get; set; }

            public double Price { get; set; }
        }

        private static readonly Sale[] Sales =
        [
            new() { Product = "Pen", Qty = 10, Price = 2.5 },
            new() { Product = "Notebook", Qty = 4, Price = 18.9 },
        ];

        [Fact]
        public async Task XlsxRecordsRoundTripAsATable()
        {
            using MemoryStream stream = new();
            await using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(stream, leaveOpen: true))
            {
                XlsxSheetWriter sheet = workbook.AddSheet("Report");
                await sheet.WriteTableAsync(Sales, "Sales", ExcelRecordLayout.FromAttributes<Sale>(),
                    ExcelTableOptions.Default with { FirstColumn = 1 }, TestContext.Current.CancellationToken);
            }

            AssertSalesTable(stream.ToArray(), "B1:D3");
        }

        [Fact]
        public async Task XlsbAsyncRecordsRoundTripAsATable()
        {
            using MemoryStream stream = new();
            await using (XlsbWorkbookWriter workbook = XlsbWorkbookWriter.Create(stream, leaveOpen: true))
            {
                XlsbSheetWriter sheet = workbook.AddSheet("Report");
                await sheet.WriteTableAsync(Stream(), "Sales", ExcelRecordLayout.FromAttributes<Sale>(),
                    ct: TestContext.Current.CancellationToken);
            }

            AssertSalesTable(stream.ToArray(), "A1:C3");
        }

        [Fact]
        public async Task TableIsClosedSoAnotherCanFollow()
        {
            using MemoryStream stream = new();
            await using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(stream, leaveOpen: true))
            {
                XlsxSheetWriter sheet = workbook.AddSheet("Report");
                await sheet.WriteTableAsync(Sales, "First", ExcelRecordLayout.FromAttributes<Sale>(), ct: TestContext.Current.CancellationToken);
                await sheet.WriteTableAsync(Sales, "Second", ExcelRecordLayout.FromAttributes<Sale>(), ct: TestContext.Current.CancellationToken);
            }

            using IExcelWorkbook read = Excel.Open(new ReadOnlyMemory<byte>(stream.ToArray()));
            Assert.Equal([("First", "A1:C3"), ("Second", "A4:C6")], read.Tables.Select(static t => (t.Name, t.Ref)));
        }

        private static async IAsyncEnumerable<Sale> Stream()
        {
            foreach (Sale sale in Sales)
            {
                await Task.Yield();
                yield return sale;
            }
        }

        private static void AssertSalesTable(byte[] package, string reference)
        {
            using IExcelWorkbook read = Excel.Open(new ReadOnlyMemory<byte>(package));
            ExcelTable table = Assert.Single(read.Tables);
            Assert.Equal(("Sales", reference), (table.Name, table.Ref));
            Assert.Equal(["Product", "Qty", "Price"], table.ColumnNames);
            List<(string?, int, double)> rows = [];
            foreach (Sale sale in ExcelParser.FromAttributes<Sale>().Parse(table.AsSheet()))
            {
                rows.Add((sale.Product, sale.Qty, sale.Price));
            }
            Assert.Equal([("Pen", 10, 2.5), ("Notebook", 4, 18.9)], rows);
        }
    }
}
