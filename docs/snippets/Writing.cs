namespace Snippets.Writing.Xlsx
{
    #region writing-xlsx
    using ExcelReader.Core.Writer;
    using ExcelReader.Core.Writer.Xlsx;
    #endregion

    internal static class Snippet
    {
        internal static async Task Run()
        {
            #region writing-xlsx
            await using var stream = File.Create("out.xlsx");
            await using var workbook = XlsxWorkbookWriter.Create(stream);

            await using (var sheet = workbook.AddSheet("Summary"))
            {
                await using (var row = await sheet.StartRowAsync())
                {
                    row.Write("Name");
                    row.Write("Total");
                    row.Write("Created");
                }

                await using (var row = await sheet.StartRowAsync())
                {
                    row.Write("Q1");
                    row.Write(42);
                    row.Write(DateTime.UtcNow);
                }
            }

            await workbook.EndAsync();
            #endregion
        }
    }
}

namespace Snippets.Writing.XlsxSharedStrings
{
    internal static class Snippet
    {
        internal static async Task Run(Stream stream)
        {
            #region writing-xlsx-shared-strings
            await using var workbook = XlsxWorkbookWriter.Create(stream, options: new XlsxWriterOptions { UseSharedStrings = true });
            #endregion
        }
    }
}

namespace Snippets.Writing.HiddenSheets
{
    #region writing-hidden-sheets
    using ExcelReader.Core; // ExcelSheetVisibility, shared by the readers and writers
    #endregion

    internal static class Snippet
    {
        internal static async Task Run(XlsxWorkbookWriter workbook)
        {
            #region writing-hidden-sheets
            await using (var data = workbook.AddSheet("Data")) { /* rows */ } // visible
            await using (var lookups = workbook.AddSheet("Lookups", ExcelSheetVisibility.Hidden)) { /* rows */ }
            await using (var audit = workbook.AddSheet("Audit", ExcelSheetVisibility.VeryHidden)) { /* rows */ } // not in Excel's unhide dialog
            #endregion
        }
    }
}

namespace Snippets.Writing.Styles
{
    #region writing-styles
    using ExcelReader.Core.Writer;
    using ExcelReader.Core.Writer.Xlsx;
    #endregion

    internal static class Snippet
    {
        internal static async Task Run(Stream stream)
        {
            #region writing-styles
            await using var workbook = XlsxWorkbookWriter.Create(stream);

            int currency = workbook.AddStyle(new CellStyle { NumberFormat = "R$ #,##0.00" });
            int header = workbook.AddStyle(new CellStyle { Bold = true });

            await using var sheet = workbook.AddSheet("Summary");
            sheet.SetColumnStyle(columnIndex: 1, currency); // before the first row
            sheet.SetColumnWidth(columnIndex: 0, width: 20);

            await using (var row = await sheet.StartRowAsync(header))
            {
                row.Write("Product");
                row.Write("Total");
            }
            await using (var row = await sheet.StartRowAsync())
            {
                row.Write("Widget");
                row.Write(1234.5);
            }

            await workbook.EndAsync();
            #endregion
        }
    }
}

namespace Snippets.Writing.Xlsb
{
    #region writing-xlsb
    using ExcelReader.Core.Writer;
    using ExcelReader.Core.Writer.Xlsb;
    #endregion

    internal static class Snippet
    {
        internal static async Task Run()
        {
            #region writing-xlsb
            await using var stream = File.Create("out.xlsb");
            await using var workbook = XlsbWorkbookWriter.Create(stream);

            await using (XlsbSheetWriter sheet = workbook.AddSheet("Summary"))
            {
                await using (XlsbRowWriter row = await sheet.StartRowAsync())
                {
                    row.Write("Name");
                    row.Write("Total");
                    row.Write("Created");
                }
            }

            await workbook.EndAsync();
            #endregion
        }
    }
}

namespace Snippets.Writing.Xls
{
    #region writing-xls
    using ExcelReader.Core.Writer;
    using ExcelReader.Core.Writer.Xls;
    #endregion

    internal static class Snippet
    {
        internal static async Task Run()
        {
            #region writing-xls
            await using var stream = File.Create("out.xls");
            await using var workbook = XlsWorkbookWriter.Create(stream);

            using (var sheet = workbook.AddSheet("Summary"))
            {
                using (var row = sheet.StartRow())
                {
                    row.Write("Name");
                    row.Write("Total");
                    row.Write("Created");
                }

                using (var row = sheet.StartRow())
                {
                    row.Write("Q1");
                    row.Write(42);
                    row.Write(DateTime.UtcNow);
                }
            }

            await workbook.EndAsync();
            #endregion
        }
    }
}

namespace Snippets.Writing.Records
{
    #region writing-records
    using ExcelReader.Core.Writer;
    using ExcelReader.Core.Writer.Xlsx;
    #endregion

    #region writing-records
    public sealed class Sale
    {
        public string? Region { get; set; }
        public int Units { get; set; }
        public decimal Revenue { get; set; }
        public DateOnly Date { get; set; }
    }
    #endregion

    internal static class Snippet
    {
        internal static async Task Run()
        {
            #region writing-records
            var sales = new[]
            {
                new Sale { Region = "North", Units = 42, Revenue = 1234.50m, Date = new DateOnly(2026, 1, 2) },
                new Sale { Region = "South", Units = 17, Revenue = 512.00m,  Date = new DateOnly(2026, 1, 3) },
            };

            await using var stream = File.Create("sales.xlsx");
            await using var workbook = XlsxWorkbookWriter.Create(stream);   // or XlsbWorkbookWriter / XlsWorkbookWriter / CsvWorkbookWriter
            await using (var sheet = workbook.AddSheet("Sales"))
            {
                await sheet.WriteRecordsAsync(sales, ExcelRecordLayout.FromAttributes<Sale>());
            }
            #endregion
        }
    }
}

namespace Snippets.Writing.Table
{
    using Snippets.Writing.Records;

    internal static class Snippet
    {
        internal static void Run(Sale[] sales)
        {
            #region writing-table
            using XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(File.Create("report.xlsx"));
            XlsxSheetWriter sheet = workbook.AddSheet("Report");
            sheet.BeginTable("Sales", ["Region", "Units", "Revenue"]);
            foreach (Sale s in sales)
            {
                using XlsxRowWriter row = sheet.StartRow();
                row.Write(s.Region);
                row.Write(s.Units);
                row.Write(s.Revenue);
            }
            sheet.EndTable();
            #endregion
        }
    }
}

namespace Snippets.Writing.TableRecords
{
    using Snippets.Writing.Records;

    internal static class Snippet
    {
        internal static async Task Run(XlsxSheetWriter sheet, Sale[] sales)
        {
            #region writing-table-records
            await sheet.WriteTableAsync(sales, "Sales", ExcelRecordLayout.FromAttributes<Sale>());
            #endregion
        }
    }
}

namespace Snippets.Writing.Prefetch
{
    internal sealed record Metric(string Name, double Value);

    internal static class Snippet
    {
        internal static async Task Run(Stream stream, Metric[] records)
        {
            #region writing-prefetch
            await using var wb = XlsxWorkbookWriter.Create(stream, leaveOpen: true, options: new XlsxWriterOptions { PrefetchWrite = true });
            XlsxSheetWriter sheet = wb.AddSheet("S1");

            foreach (var record in records)
            {
                using XlsxRowWriter row = sheet.StartRow();
                row.Write(record.Name);
                row.Write(record.Value);
            }

            await sheet.EndAsync();
            #endregion
        }
    }
}
