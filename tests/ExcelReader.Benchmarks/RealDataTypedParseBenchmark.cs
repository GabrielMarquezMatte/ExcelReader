using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;

namespace ExcelReader.Benchmarks
{
    // Full-schema typed parse of the real 65K_Records_Data workbook (all 14 columns, 65,535 rows) -
    // the JIT/CoreCLR counterpart to the C++/Rust native-binding benchmarks' `parse_sheet<FullRow>`
    // (cpp/benchmarks/benchmark_compare.cpp, rust/excelreader/benches/compare_bench.rs), which run
    // the same ExcelReader.Core parsing logic ahead-of-time compiled via NativeAOT and called
    // through the C ABI instead of JIT-compiled in-process. RealDataReadBenchmark's Xlsx_ExcelReader
    // is not a fair comparison against those: it reads raw cells zero-copy with no column-name
    // resolution, while parse_sheet<FullRow> resolves 14 named columns into a typed struct - this
    // class does the same schema-resolution work ExcelParser<T> does, so the two are matched work.
    public sealed class FullRow
    {
        [ExcelColumn("Region")]
        public string Region { get; set; } = "";

        [ExcelColumn("Country")]
        public string Country { get; set; } = "";

        [ExcelColumn("Item Type")]
        public string ItemType { get; set; } = "";

        [ExcelColumn("Sales Channel")]
        public string SalesChannel { get; set; } = "";

        [ExcelColumn("Order Priority")]
        public string OrderPriority { get; set; } = "";

        [ExcelColumn("Order Date")]
        public DateTime OrderDate { get; set; }

        [ExcelColumn("Order ID")]
        public long OrderId { get; set; }

        [ExcelColumn("Ship Date")]
        public DateTime ShipDate { get; set; }

        [ExcelColumn("Units Sold")]
        public long UnitsSold { get; set; }

        [ExcelColumn("Unit Price")]
        public double UnitPrice { get; set; }

        [ExcelColumn("Unit Cost")]
        public double UnitCost { get; set; }

        [ExcelColumn("Total Revenue")]
        public double TotalRevenue { get; set; }

        [ExcelColumn("Total Cost")]
        public double TotalCost { get; set; }

        [ExcelColumn("Total Profit")]
        public double TotalProfit { get; set; }
    }

    [MemoryDiagnoser]
    public class RealDataTypedParseBenchmark
    {
        private byte[] _xlsx = [];
        private byte[] _xlsb = [];

        [GlobalSetup]
        public void Setup()
        {
            string dir = Path.Combine(AppContext.BaseDirectory, "Data");
            _xlsx = File.ReadAllBytes(Path.Combine(dir, "65K_Records_Data.xlsx"));
            _xlsb = File.ReadAllBytes(Path.Combine(dir, "65K_Records_Data.xlsb"));
        }

        private static long Accumulate(FullRow row)
        {
            return row.Region.Length + row.Country.Length + row.ItemType.Length
                + row.SalesChannel.Length + row.OrderPriority.Length
                + row.OrderDate.Ticks + row.OrderId + row.ShipDate.Ticks + row.UnitsSold
                + (long)row.UnitPrice + (long)row.UnitCost + (long)row.TotalRevenue
                + (long)row.TotalCost + (long)row.TotalProfit;
        }

        [Benchmark(Baseline = true)]
        public long Xlsx_ExcelParser()
        {
            using MemoryStream ms = new(_xlsx, writable: false);
            using XlsxReader reader = Excel.From(ms);
            long acc = 0;
            foreach (FullRow row in new ExcelParser<FullRow>().Parse(reader))
            {
                acc += Accumulate(row);
            }
            return acc;
        }

        [Benchmark]
        public long Xlsb_ExcelParser()
        {
            using MemoryStream ms = new(_xlsb, writable: false);
            using XlsbReader reader = Excel.FromXlsb(ms);
            long acc = 0;
            foreach (FullRow row in new ExcelParser<FullRow>().Parse(reader))
            {
                acc += Accumulate(row);
            }
            return acc;
        }
    }
}
