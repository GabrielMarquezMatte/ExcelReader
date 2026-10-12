using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Reader.Xlsx;

namespace ExcelReader.Benchmarks
{
    [MemoryDiagnoser]
    public class SheetParallelBenchmark
    {
        private sealed class Sum : IRowAccumulator<Sum, FullRow>
        {
            public long Total;

            public void Add(FullRow model)
            {
                Total += model.OrderId + model.UnitsSold + (long)model.TotalRevenue + model.Region.Length;
            }

            public void Merge(Sum following)
            {
                Total += following.Total;
            }
        }

        private readonly ExcelParser<FullRow> _parser = ExcelParser.FromAttributes<FullRow>();
        private byte[] _xlsx = [];
        private byte[] _xlsb = [];

        [Params(1, 2, 4, 8)]
        public int Dop { get; set; }

        [GlobalSetup]
        public void Setup()
        {
            string dir = Path.Combine(AppContext.BaseDirectory, "Data");
            _xlsx = File.ReadAllBytes(Path.Combine(dir, "65K_Records_Data.xlsx"));
            _xlsb = File.ReadAllBytes(Path.Combine(dir, "65K_Records_Data.xlsb"));
        }

        [Benchmark]
        public async Task<long> Xlsx_Aggregate()
        {
            using XlsxWorkbook workbook = Excel.FromXlsx(_xlsx);
            Sum sum = await SheetParallel.AggregateAsync<Sum, FullRow>(workbook.FirstSheet, _parser, new SheetParallelOptions { DegreeOfParallelism = Dop, HeaderRow = 1 });
            return sum.Total;
        }

        [Benchmark]
        public async Task<long> Xlsb_Aggregate()
        {
            using XlsbWorkbook workbook = Excel.FromXlsb(_xlsb);
            Sum sum = await SheetParallel.AggregateAsync<Sum, FullRow>(workbook.FirstSheet, _parser, new SheetParallelOptions { DegreeOfParallelism = Dop, HeaderRow = 1 });
            return sum.Total;
        }
    }
}
