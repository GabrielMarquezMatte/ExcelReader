using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Benchmarks
{
    [MemoryDiagnoser]
    public class LargeSheetParallelBenchmark
    {
        private const int Rows = 1_000_000;
        private const int CsvRows = 4_000_000;

        private sealed class Sum : IRowAccumulator<Sum, Record>
        {
            public long Total;

            public void Add(Record model)
            {
                Total += model.Id + (long)model.Value + model.Name!.Length;
            }

            public void Merge(Sum following)
            {
                Total += following.Total;
            }
        }

        private sealed class CsvSum : IRowAccumulator<CsvSum, WideRowRef>
        {
            public long Total;

            public void Add(WideRowRef model)
            {
                Total += model.Units + (long)model.TotalRevenue + model.Region.Length;
            }

            public void Merge(CsvSum following)
            {
                Total += following.Total;
            }
        }

        private readonly ExcelParser<Record> _parser = ExcelParser.FromAttributes<Record>();
        private readonly ExcelParser<WideRow> _wideParser = ExcelParser.FromAttributes<WideRow>();
        private readonly ExcelParser<WideRowRef> _wideRefParser = ExcelParser.FromAttributes<WideRowRef>();
        private byte[] _xlsx = [];
        private byte[] _xlsb = [];
        private byte[] _csv = [];

        [Params(1, 2, 4, 8, 16)]
        public int Dop { get; set; }

        [GlobalSetup]
        public async Task SetupAsync()
        {
            _csv = CsvGenerator.BuildConversionHeavy(CsvRows);
            List<Record> records = WorkbookGenerator.Records(Rows);
            ExcelRecordLayout<Record> layout = ExcelRecordLayout.FromAttributes<Record>();
            await using (var ms = new MemoryStream(32 * 1024 * 1024))
            {
                await using (var writer = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
                {
                    await using (var sheet = writer.AddSheet("S1"))
                    {
                        await sheet.WriteRecordsParallelAsync(records, layout);
                    }
                }
                _xlsx = ms.ToArray();
            }
            await using (var ms = new MemoryStream(32 * 1024 * 1024))
            {
                await using (var writer = XlsbWorkbookWriter.Create(ms, leaveOpen: true))
                {
                    await using (var sheet = writer.AddSheet("S1"))
                    {
                        await sheet.WriteRecordsParallelAsync(records, layout);
                    }
                }
                _xlsb = ms.ToArray();
            }
        }

        [Benchmark]
        public async Task<long> Xlsx_Aggregate()
        {
            using XlsxWorkbook workbook = Excel.FromXlsx(_xlsx);
            Sum sum = await SheetParallel.AggregateAsync<Sum, Record>(workbook.FirstSheet, _parser, new SheetParallelOptions { DegreeOfParallelism = Dop, HeaderRow = 1 });
            return sum.Total;
        }

        [Benchmark]
        public async Task<long> Xlsb_Aggregate()
        {
            using XlsbWorkbook workbook = Excel.FromXlsb(_xlsb);
            Sum sum = await SheetParallel.AggregateAsync<Sum, Record>(workbook.FirstSheet, _parser, new SheetParallelOptions { DegreeOfParallelism = Dop, HeaderRow = 1 });
            return sum.Total;
        }

        [Benchmark]
        public async Task<long> Xlsx_Parse()
        {
            long total = 0;
            using XlsxWorkbook workbook = Excel.FromXlsx(_xlsx);
            await foreach (Record row in SheetParallel.ParseAsync(workbook.FirstSheet, _parser, new SheetParallelOptions { DegreeOfParallelism = Dop }))
            {
                total += row.Id;
            }
            return total;
        }

        [Benchmark]
        public async Task<long> Xlsb_Parse()
        {
            long total = 0;
            using XlsbWorkbook workbook = Excel.FromXlsb(_xlsb);
            await foreach (Record row in SheetParallel.ParseAsync(workbook.FirstSheet, _parser, new SheetParallelOptions { DegreeOfParallelism = Dop }))
            {
                total += row.Id;
            }
            return total;
        }

        [Benchmark]
        public async Task<long> Csv_Aggregate()
        {
            using CsvReader csv = Excel.FromCsv(_csv);
            CsvSum sum = await SheetParallel.AggregateAsync<CsvSum, WideRowRef>(csv.FirstSheet, _wideRefParser, new SheetParallelOptions { DegreeOfParallelism = Dop, HeaderRow = 1 });
            if (sum.Total == 0)
            {
                throw new InvalidOperationException("WideRowRef bound no columns; the aggregate benchmark is measuring nothing.");
            }
            return sum.Total;
        }

        [Benchmark]
        public async Task<long> Csv_Parse()
        {
            long total = 0;
            using CsvReader csv = Excel.FromCsv(_csv);
            await foreach (WideRow row in SheetParallel.ParseAsync(csv.FirstSheet, _wideParser, new SheetParallelOptions { DegreeOfParallelism = Dop }))
            {
                total += row.Units;
            }
            return total;
        }
    }
}
