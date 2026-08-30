using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;

namespace ExcelReader.Benchmarks
{
    public sealed class WideRow
    {
        public string? Region { get; set; }
        public string? Country { get; set; }
        public DateTime OrderDate { get; set; }
        public decimal UnitPrice { get; set; }
        public decimal TotalRevenue { get; set; }
        public int Units { get; set; }
    }

    public sealed class NarrowRow
    {
        public int A { get; set; }
        public int B { get; set; }
        public int C { get; set; }
    }

    // Measures the parallel CSV path against the sequential parser on a corpus large enough for
    // partitioning to mean anything. The dop=1 leg is the baseline; it runs the sequential fallback,
    // so the speedup column reads directly as "what parallelism bought".
    [MemoryDiagnoser]
    public class CsvParallelParseBenchmark
    {
        [Params(1, 2, 4, 8, 16)]
        public int Dop { get; set; }

        private string _wide = "";
        private string _narrow = "";

        [GlobalSetup]
        public void Setup()
        {
            // 2,000,000 rows measured ~92 MB on this machine (the brief's row count assumed a
            // shorter average row); 4,300,000 lands the conversion-heavy corpus at ~200 MB too.
            _wide = CsvGenerator.WriteConversionHeavyFile(4_300_000);
            _narrow = CsvGenerator.WriteNarrowIntFile(8_000_000);
        }

        [GlobalCleanup]
        public void Cleanup()
        {
            File.Delete(_wide);
            File.Delete(_narrow);
        }

        [Benchmark(Baseline = true)]
        public async Task<long> ConversionHeavy()
        {
            long n = 0;
            await foreach (WideRow row in Excel.ParseCsvParallelAsync<WideRow>(_wide, Dop))
            {
                n += row.Units;
            }
            return n;
        }

        [Benchmark]
        public async Task<long> NarrowInt()
        {
            long n = 0;
            await foreach (NarrowRow row in Excel.ParseCsvParallelAsync<NarrowRow>(_narrow, Dop))
            {
                n += row.A;
            }
            return n;
        }
    }
}
