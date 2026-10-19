using System.Runtime.InteropServices;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;

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
    [StructLayout(LayoutKind.Auto)]
    public struct WideRowStruct
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
            using CsvReader csv = Excel.FromCsvFile(_wide);
            await foreach (WideRow row in SheetParallel.ParseAsync(csv.FirstSheet, ExcelParser.FromAttributes<WideRow>(), new SheetParallelOptions { DegreeOfParallelism = Dop }))
            {
                n += row.Units;
            }
            return n;
        }

        /// <summary>
        /// The same parse into a struct with <see cref="ExcelReaderOptions.InternStrings"/>: no object per row, and the
        /// corpus's few distinct region and country names are shared, so the GC has almost nothing to collect.
        /// </summary>
        [Benchmark]
        public async Task<long> ConversionHeavyStructInterned()
        {
            long n = 0;
            using CsvReader csv = Excel.FromCsvFile(_wide, new ExcelReaderOptions { InternStrings = true });
            await foreach (WideRowStruct row in SheetParallel.ParseAsync(csv.FirstSheet, ExcelParser.FromAttributes<WideRowStruct>(), new SheetParallelOptions { DegreeOfParallelism = Dop }))
            {
                n += row.Units;
            }
            return n;
        }

        [Benchmark]
        public async Task<long> NarrowInt()
        {
            long n = 0;
            using CsvReader csv = Excel.FromCsvFile(_narrow);
            await foreach (NarrowRow row in SheetParallel.ParseAsync(csv.FirstSheet, ExcelParser.FromAttributes<NarrowRow>(), new SheetParallelOptions { DegreeOfParallelism = Dop }))
            {
                n += row.A;
            }
            return n;
        }

        private sealed class Aggregation : IRowAccumulator<Aggregation, WideRowRef>
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

        private sealed class AggregationNarrow : IRowAccumulator<AggregationNarrow, NarrowRowStruct>
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
            SheetParallelOptions options = new() { DegreeOfParallelism = Dop, HeaderRow = 1 };
            ExcelParser<WideRowRef> map = ExcelParser.FromAttributes<WideRowRef>();
            using CsvReader csv = Excel.FromCsvFile(_wide);
            var aggregation = await SheetParallel.AggregateAsync<Aggregation, WideRowRef>(csv.FirstSheet, map, options);
            if (aggregation.Units == 0)
            {
                throw new InvalidOperationException("WideRowRef bound no columns; the aggregate benchmark is measuring nothing.");
            }
            return aggregation.Units;
        }

        [Benchmark]
        public async Task<long> NarrowIntAggregate()
        {
            SheetParallelOptions options = new() { DegreeOfParallelism = Dop, HeaderRow = 1 };
            ExcelParser<NarrowRowStruct> map = ExcelParser.FromAttributes<NarrowRowStruct>();
            using CsvReader csv = Excel.FromCsvFile(_narrow);
            var aggregation = await SheetParallel.AggregateAsync<AggregationNarrow, NarrowRowStruct>(csv.FirstSheet, map, options);
            if (aggregation.Units == 0)
            {
                throw new InvalidOperationException("NarrowRowStruct bound no columns; the aggregate benchmark is measuring nothing.");
            }
            return aggregation.Units;
        }
    }
}
