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
            var stream = File.Create("out.xlsx");
            #region writing-xlsx
            await using (stream.ConfigureAwait(false))
            {
                await using var workbook = XlsxWorkbookWriter.Create(stream).ConfigureAwait(false);

                await using (var sheet = workbook.AddSheet("Summary").ConfigureAwait(false))
                {
                    await using (var row = (await sheet.StartRowAsync()).ConfigureAwait(false))
                    {
                        row.Write("Name");
                        row.Write("Total");
                        row.Write("Created");
                    }

                    await using (var row = (await sheet.StartRowAsync()).ConfigureAwait(false))
                    {
                        row.Write("Q1");
                        row.Write(42);
                        row.Write(DateTime.UtcNow);
                    }
                }

                await workbook.EndAsync().ConfigureAwait(false);
            }
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
            var workbook = XlsxWorkbookWriter.Create(stream, options: new XlsxWriterOptions { UseSharedStrings = true });
            #region writing-xlsx-shared-strings
            await using (workbook.ConfigureAwait(false))
            {
            }
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
            var data = workbook.AddSheet("Data");
            #region writing-hidden-sheets
            await using (data.ConfigureAwait(false)) { /* rows */ } // visible
            var lookups = workbook.AddSheet("Lookups", ExcelSheetVisibility.Hidden);
            await using (lookups.ConfigureAwait(false)) { /* rows */ }
            var audit = workbook.AddSheet("Audit", ExcelSheetVisibility.VeryHidden);
            await using (audit.ConfigureAwait(false)) { /* rows */ } // not in Excel's unhide dialog
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
            var workbook = XlsxWorkbookWriter.Create(stream);
            #region writing-styles
            await using (workbook.ConfigureAwait(false))
            {
                int currency = workbook.AddStyle(new CellStyle { NumberFormat = "R$ #,##0.00" });
                int header = workbook.AddStyle(new CellStyle { Bold = true });

                await using var sheet = workbook.AddSheet("Summary").ConfigureAwait(false);
                sheet.SetColumnStyle(columnIndex: 1, currency); // before the first row
                sheet.SetColumnWidth(columnIndex: 0, width: 20);

                await using (var row = (await sheet.StartRowAsync(header).ConfigureAwait(false)).ConfigureAwait(false))
                {
                    row.Write("Product");
                    row.Write("Total");
                }
                {
                    var row = await sheet.StartRowAsync().ConfigureAwait(false);
                    await using (row.ConfigureAwait(false))
                    {
                        row.Write("Widget");
                        row.Write(1234.5);
                    }
                }

                await workbook.EndAsync().ConfigureAwait(false);
            }
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
            var stream = File.Create("out.xlsb");
            #region writing-xlsb
            await using (stream.ConfigureAwait(false))
            {
                await using var workbook = XlsbWorkbookWriter.Create(stream).ConfigureAwait(false);

                await using (XlsbSheetWriter sheet = workbook.AddSheet("Summary").ConfigureAwait(false).ConfigureAwait(false))
                {
                    await using (XlsbRowWriter row = (await sheet.StartRowAsync().ConfigureAwait(false)).ConfigureAwait(false))
                    {
                        row.Write("Name");
                        row.Write("Total");
                        row.Write("Created");
                    }
                }

                await workbook.EndAsync().ConfigureAwait(false);
            }
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
            var stream = File.Create("out.xls");
            #region writing-xls
            await using (stream.ConfigureAwait(false))
            {
                await using var workbook = XlsWorkbookWriter.Create(stream).ConfigureAwait(false);

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

                await workbook.EndAsync().ConfigureAwait(false);
            }
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

            var stream = File.Create("sales.xlsx");
            await using (stream.ConfigureAwait(false))
            {
                await using var workbook = XlsxWorkbookWriter.Create(stream).ConfigureAwait(false);   // or XlsbWorkbookWriter / XlsWorkbookWriter / CsvWorkbookWriter
                await using (var sheet = workbook.AddSheet("Sales").ConfigureAwait(false))
                {
                    await sheet.WriteRecordsAsync(sales, ExcelRecordLayout.FromAttributes<Sale>()).ConfigureAwait(false);
                }
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
            await sheet.WriteTableAsync(sales, "Sales", ExcelRecordLayout.FromAttributes<Sale>()).ConfigureAwait(false);
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
            var wb = XlsxWorkbookWriter.Create(stream, leaveOpen: true, options: new XlsxWriterOptions { PrefetchWrite = true });
            #region writing-prefetch
            await using (wb.ConfigureAwait(false))
            {
                XlsxSheetWriter sheet = wb.AddSheet("S1");

                foreach (var record in records)
                {
                    using XlsxRowWriter row = sheet.StartRow();
                    row.Write(record.Name);
                    row.Write(record.Value);
                }

                await sheet.EndAsync().ConfigureAwait(false);
            }
            #endregion
        }
    }
}

namespace Snippets.Writing.Parallel
{
    public sealed record Metric(string Name, double Value);

    internal static class Snippet
    {
        internal static async Task Run(Stream stream, IEnumerable<Metric> records)
        {
            var wb = XlsxWorkbookWriter.Create(stream, leaveOpen: true);
            #region writing-parallel
            await using (wb.ConfigureAwait(false))
            {
                XlsxSheetWriter sheet = wb.AddSheet("S1");

                // Header row from the layout, then one row per record.
                await sheet.WriteRecordsParallelAsync(records, ExcelRecordLayout.FromAttributes<Metric>()).ConfigureAwait(false);

                // Or format each row yourself; this delegate runs on several threads at once.
                await sheet.WriteRecordsParallelAsync(records, static (XlsxRowWriter row, Metric record) =>
                {
                    row.Write(record.Name);
                    row.Write(record.Value);
                }, new ParallelWriteOptions { DegreeOfParallelism = 4 }).ConfigureAwait(false);

                await sheet.EndAsync().ConfigureAwait(false);
            }
            #endregion
        }
    }
}
