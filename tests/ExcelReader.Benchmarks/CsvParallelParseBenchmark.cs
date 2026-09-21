using System.Runtime.InteropServices;
using BenchmarkDotNet.Attributes;
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
    public ref struct WideRowRef
    {
        public ReadOnlySpan<byte> Region { get; set; }
        public ReadOnlySpan<byte> Country { get; set; }
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
    [StructLayout(LayoutKind.Auto)]
    public struct NarrowRowStruct
    {
        public int A { get; set; }
        public int B { get; set; }
        public int C { get; set; }
    }

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
            await foreach (WideRow row in Excel.ParseCsvParallelAsync<WideRow>(_wide, new CsvParallelOptions { DegreeOfParallelism = Dop }))
            {
                n += row.Units;
            }
            return n;
        }

        [Benchmark]
        public async Task<long> NarrowInt()
        {
            long n = 0;
            await foreach (NarrowRow row in Excel.ParseCsvParallelAsync<NarrowRow>(_narrow, new CsvParallelOptions { DegreeOfParallelism = Dop }))
            {
                n += row.A;
            }
            return n;
        }

        private sealed class Aggregation : ICsvAccumulator<Aggregation, WideRowRef>
        {
            public long Units { get; private set; }
            public void Add(WideRowRef model)
            {
                Units += model.Units;
            }

            public void Merge(Aggregation following)
            {
                Units += following.Units;
            }
        }

        private sealed class AggregationNarrow : ICsvAccumulator<AggregationNarrow, NarrowRowStruct>
        {
            public long Units { get; private set; }
            public void Add(NarrowRowStruct model)
            {
                Units += model.A;
            }

            public void Merge(AggregationNarrow following)
            {
                Units += following.Units;
            }
        }

        [Benchmark]
        public async Task<long> ConversionHeavyAggregate()
        {
            CsvParallelOptions options = new() { DegreeOfParallelism = Dop, HeaderRow = 1 };
            CsvModelMap<WideRowRef> map = CsvModelMap.FromAttributes<WideRowRef>();
            var aggregation = await Excel.AggregateCsvParallelAsync<Aggregation, WideRowRef>(_wide, map, options);
            // Every generated row has Units >= 1, so a zero total means the map bound no columns
            // and this benchmark is timing record splitting rather than conversion.
            if (aggregation.Units == 0)
            {
                throw new InvalidOperationException("WideRowRef bound no columns; the aggregate benchmark is measuring nothing.");
            }
            return aggregation.Units;
        }

        [Benchmark]
        public async Task<long> NarrowIntAggregate()
        {
            CsvParallelOptions options = new() { DegreeOfParallelism = Dop, HeaderRow = 1 };
            CsvModelMap<NarrowRowStruct> map = CsvModelMap.FromAttributes<NarrowRowStruct>();
            var aggregation = await Excel.AggregateCsvParallelAsync<AggregationNarrow, NarrowRowStruct>(_narrow, map, options);
            if (aggregation.Units == 0)
            {
                throw new InvalidOperationException("NarrowRowStruct bound no columns; the aggregate benchmark is measuring nothing.");
            }
            return aggregation.Units;
        }
    }
}
