namespace Snippets.Reading.ReadRows
{
    #region reading-read-rows
    using ExcelReader.Core.Reader;
    #endregion

    internal static class Snippet
    {
        internal static void Run()
        {
            #region reading-read-rows
            using var workbook = Excel.FromXlsxFile("report.xlsx");

            foreach (var row in workbook.FirstSheet)
            {
                string name = row[0].GetString();

                if (row[1].TryParse(null, out int quantity))
                {
                    Console.WriteLine($"{name}: {quantity}");
                }

                if (row[2].Type == CellType.Date && row[2].TryGetDateTime(workbook.IsDate1904, out var date))
                {
                    Console.WriteLine(date);
                }
            }
            #endregion
        }
    }
}

namespace Snippets.Reading.ParallelParse
{
    internal sealed class SalesRow
    {
        public decimal Revenue { get; set; }
    }

    internal static class Snippet
    {
        private static decimal Total;

        internal static async Task Run()
        {
            #region reading-parallel-parse
            using XlsxWorkbook workbook = Excel.FromXlsxFile("big.xlsx");

            await foreach (var row in SheetParallel.ParseAsync(workbook.FirstSheet, ExcelParser.FromAttributes<SalesRow>(), new SheetParallelOptions { DegreeOfParallelism = 8 }))
            {
                Total += row.Revenue;
            }
            #endregion
        }
    }
}

namespace Snippets.Reading.ParallelAggregate
{
    #region reading-parallel-aggregate
    public ref struct SaleRow
    {
        [ExcelColumn("Region")] public ReadOnlySpan<byte> Region { get; set; }
        [ExcelColumn("Units")] public int Units { get; set; }
    }

    public sealed class UnitTotal : IRowAccumulator<UnitTotal, SaleRow>
    {
        public long Units;

        public void Add(SaleRow model) => Units += model.Units;

        public void Merge(UnitTotal following) => Units += following.Units;
    }
    #endregion

    internal static class Snippet
    {
        internal static async Task Run()
        {
            #region reading-parallel-aggregate
            using XlsxWorkbook workbook = Excel.FromXlsxFile("big.xlsx");

            UnitTotal total = await SheetParallel.AggregateAsync<UnitTotal, SaleRow>(
                workbook.FirstSheet,
                ExcelParser.FromAttributes<SaleRow>(),
                new SheetParallelOptions { HeaderRow = 1 });
            #endregion
        }
    }
}

namespace Snippets.Reading.ParallelForeach
{
    using Snippets.Reading.ParallelAggregate;

    internal static class Snippet
    {
        internal static async Task Run()
        {
            #region reading-parallel-foreach
            using XlsxWorkbook workbook = Excel.FromXlsxFile("big.xlsx");

            long total = 0;
            await SheetParallel.ForEachAsync(
                workbook.FirstSheet,
                ExcelParser.FromAttributes<SaleRow>(),
                row => Interlocked.Add(ref total, row.Units),
                new SheetParallelOptions { DegreeOfParallelism = 8, HeaderRow = 1 });
            #endregion
        }
    }
}

namespace Snippets.Reading.ArrowRead
{
    #region reading-arrow-read
    using ExcelReader.Arrow;
    using ExcelReader.Core.Reader;
    #endregion

    internal static class Snippet
    {
        internal static void Run()
        {
            #region reading-arrow-read
            using var workbook = Excel.FromXlsxFile("report.xlsx");
            Apache.Arrow.RecordBatch batch = workbook.FirstSheet.ToArrowRecordBatch();
            #endregion
        }
    }
}

namespace Snippets.Reading.ArrowWrite
{
    #region reading-arrow-write
    using ExcelReader.Arrow;
    using ExcelReader.Core.Writer;
    using ExcelReader.Core.Writer.Xlsx;
    #endregion

    internal static class Snippet
    {
        internal static void Run(Apache.Arrow.RecordBatch batch)
        {
            #region reading-arrow-write
            using XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(File.Create("report.xlsx"));
            workbook.WriteRecordBatch(batch); // adds the sheet, writes the header + every row, ends the workbook
            #endregion
        }
    }
}

namespace Snippets.Reading.Open
{
    #region reading-open
    using ExcelReader.Core.Reader;
    #endregion

    internal static class Snippet
    {
        internal static void Run()
        {
            #region reading-open
            using IExcelWorkbook workbook = Excel.Open("report.xlsx"); // or report.xlsb / report.xls

            foreach (var row in workbook.FirstSheet)
            {
                Console.WriteLine(row[0].GetString());
            }
            #endregion
        }
    }
}

namespace Snippets.Reading.OpenEverySheet
{
    internal static class Snippet
    {
        internal static void Run()
        {
            #region reading-open-every-sheet
            using IExcelWorkbook workbook = Excel.Open("report.xlsx");

            for (int i = 0; i < workbook.SheetCount; i++)
            {
                IExcelSheet sheet = workbook.SheetAt(i);
                Console.WriteLine(sheet.Name);
                foreach (var row in sheet)
                {
                    Console.WriteLine(row[0].GetString());
                }
            }
            #endregion
        }
    }
}

namespace Snippets.Reading.OpenTryGetSheet
{
    internal static class Snippet
    {
        internal static void Run(IExcelWorkbook workbook)
        {
            #region reading-open-try-get-sheet
            if (workbook.TryGetSheet("Totals", out var totals))
            {
                foreach (var row in totals)
                {
                    Console.WriteLine(row[0].GetString());
                }
            }
            #endregion
        }
    }
}

namespace Snippets.Reading.OpenVisibility
{
    #region reading-open-visibility
    using ExcelReader.Core; // ExcelSheetVisibility, shared by the readers and writers
    #endregion

    internal static class Snippet
    {
        internal static void Run(IExcelWorkbook workbook)
        {
            #region reading-open-visibility
            for (int i = 0; i < workbook.SheetCount; i++)
            {
                IExcelSheet sheet = workbook.SheetAt(i);
                if (sheet.Visibility != ExcelSheetVisibility.Visible)
                {
                    continue; // skip hidden and veryHidden sheets
                }
                Console.WriteLine(sheet.Name);
            }
            #endregion
        }
    }
}

namespace Snippets.Reading.OpenCsv
{
    internal static class Snippet
    {
        internal static void Run()
        {
            #region reading-open-csv
            using IExcelWorkbook workbook = Excel.Open("report.csv", ExcelFileFormat.Csv);
            #endregion
        }
    }
}

namespace Snippets.Reading.Table
{
    internal sealed class ProductRow
    {
        public string Product { get; set; } = "";
        public int Qty { get; set; }
    }

    internal static class Snippet
    {
        internal static void Run()
        {
            #region reading-table
            using IExcelWorkbook workbook = Excel.Open("sales.xlsx");
            if (workbook.TryGetTable("Sales", out ExcelTable? sales))
            {
                foreach (ProductRow row in ExcelParser.FromAttributes<ProductRow>().Parse(sales.AsSheet()))
                {
                    Console.WriteLine($"{row.Product}: {row.Qty}");
                }
            }
            #endregion
        }
    }
}

namespace Snippets.Reading.SheetsInParallel
{
    internal static class Snippet
    {
        internal static void Run(string path)
        {
            #region reading-sheets-in-parallel
            using XlsxWorkbook workbook = Excel.FromXlsxFile(path);

            Parallel.For(0, workbook.Sheets.Count, i =>
            {
                foreach (Row row in workbook.Sheets[i])
                {
                    // consume
                }
            });
            #endregion
        }
    }
}

namespace Snippets.Reading.Async
{
    #region reading-async
    using ExcelReader.Core.Reader;
    #endregion

    internal static class Snippet
    {
        internal static async Task Run(CancellationToken cancellationToken)
        {
            #region reading-async
            await using var workbook = await Excel.FromXlsxFileAsync("report.xlsx", ct: cancellationToken);

            await foreach (var row in workbook.FirstSheet)
            {
                Console.WriteLine(row[0].GetString());
            }
            #endregion
        }
    }
}

namespace Snippets.Reading.AsyncEnumerator
{
    internal static class Snippet
    {
        internal static async Task Run(CancellationToken cancellationToken)
        {
            #region reading-async-enumerator
            await using var workbook = await Excel.FromXlsxFileAsync("report.xlsx", ct: cancellationToken);
            await using var rows = workbook.FirstSheet.GetAsyncEnumerator(cancellationToken);

            while (await rows.MoveNextAsync())
            {
                var row = rows.Current;
                Console.WriteLine(row[0].GetString());
            }
            #endregion
        }
    }
}

namespace Snippets.Reading.Prefetch
{
    internal static class Snippet
    {
        internal static void Run()
        {
            #region reading-prefetch
            var options = new ExcelReaderOptions { PrefetchDecompression = true };
            using var workbook = Excel.FromXlsxFile("report.xlsx", options);

            foreach (var row in workbook.FirstSheet)
            {
                Console.WriteLine(row[0].GetString());
            }
            #endregion
        }
    }
}

namespace Snippets.Reading.DataReader
{
    #region reading-data-reader
    using System.Data;
    using ExcelReader.Core.Reader;
    #endregion

    internal static class Snippet
    {
        internal static void Run()
        {
            #region reading-data-reader
            using IExcelWorkbook workbook = Excel.Open("report.xlsx");
            using IDataReader data = new SheetDataReader(workbook.FirstSheet); // header on row 1 by default

            var table = new DataTable();
            table.Load(data);
            #endregion
        }
    }
}
