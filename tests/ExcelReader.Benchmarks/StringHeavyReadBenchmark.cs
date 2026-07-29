using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using Sylvan.Data.Excel;
using static ExcelReader.Benchmarks.BenchmarkAccumulators;

namespace ExcelReader.Benchmarks
{
    // Reads StringHeavyWorkbookGenerator's fixture cell-by-cell: 8 text columns against 3
    // numeric/date columns, tens of thousands of distinct shared strings. Companion to
    // RealDataReadBenchmark, which covers the 65K_Records_Data.* corpus — that corpus has only
    // 5 KB of shared strings across ~910K cells, so it never exercises the shared-string cache,
    // dictionary lookups or string materialization that prefetch and parsing changes touch. This
    // class isolates that axis: same [MemoryDiagnoser]/AccumulateRow shape, prefetch on and off,
    // for both ZIP-based formats (xlsx, xlsb) prefetch actually affects.
    [MemoryDiagnoser]
    public class StringHeavyReadBenchmark
    {
        [Params(65_536)]
        public int Rows { get; set; }

        private static readonly ExcelReaderOptions _prefetchOptions = new() { PrefetchDecompression = true };

        private byte[] _xlsx = [];
        private byte[] _xlsb = [];

        // BenchmarkDotNet runs each [Benchmark] in its own process, so an unsplit setup would rebuild
        // both 65K-row fixtures for every method — doubling suite wall time to build one it never reads.
        [GlobalSetup(Targets = [nameof(Xlsx_ExcelReader), nameof(Xlsx_ExcelReader_Prefetch), nameof(Xlsx_Sylvan)])]
        public async Task SetupXlsxAsync()
        {
            _xlsx = await StringHeavyWorkbookGenerator.BuildXlsxAsync(Rows);
        }

        [GlobalSetup(Targets = [nameof(Xlsb_ExcelReader), nameof(Xlsb_ExcelReader_Prefetch), nameof(Xlsb_Sylvan)])]
        public async Task SetupXlsbAsync()
        {
            _xlsb = await StringHeavyWorkbookGenerator.BuildXlsbAsync(Rows);
        }

        // --- XLSX ---

        [Benchmark(Baseline = true)]
        public long Xlsx_ExcelReader()
        {
            using MemoryStream ms = new(_xlsx, writable: false);
            using XlsxReader reader = Excel.From(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsx_ExcelReader_Prefetch()
        {
            using MemoryStream ms = new(_xlsx, writable: false);
            using XlsxReader reader = Excel.From(ms, options: _prefetchOptions);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsx_Sylvan()
        {
            using MemoryStream ms = new(_xlsx, writable: false);
            using ExcelDataReader reader = ExcelDataReader.Create(ms, ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions());
            return AccumulateSylvanExcel(reader);
        }

        // Matched-work counterpart to Xlsx_ExcelReader: materializes a string per cell like
        // Xlsx_Sylvan is forced to, instead of reading the zero-copy span.
        [Benchmark]
        public long Xlsx_ExcelReader_Materialized()
        {
            using MemoryStream ms = new(_xlsx, writable: false);
            using XlsxReader reader = Excel.From(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRowMaterialized(row); }
            return acc;
        }

        // --- XLSB ---

        [Benchmark]
        public long Xlsb_ExcelReader()
        {
            using MemoryStream ms = new(_xlsb, writable: false);
            using XlsbReader reader = Excel.FromXlsb(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsb_ExcelReader_Prefetch()
        {
            using MemoryStream ms = new(_xlsb, writable: false);
            using XlsbReader reader = Excel.FromXlsb(ms, options: _prefetchOptions);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsb_Sylvan()
        {
            using MemoryStream ms = new(_xlsb, writable: false);
            using ExcelDataReader reader = ExcelDataReader.Create(ms, ExcelWorkbookType.ExcelBinary, new ExcelDataReaderOptions());
            return AccumulateSylvanExcel(reader);
        }

        // Matched-work counterpart to Xlsb_ExcelReader — see Xlsx_ExcelReader_Materialized.
        [Benchmark]
        public long Xlsb_ExcelReader_Materialized()
        {
            using MemoryStream ms = new(_xlsb, writable: false);
            using XlsbReader reader = Excel.FromXlsb(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRowMaterialized(row); }
            return acc;
        }
    }
}
