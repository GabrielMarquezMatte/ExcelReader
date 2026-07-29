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
        // Every new [Benchmark] MUST be added to the matching Targets list: only the setup whose Targets
        // name the running method executes, so an unlisted method reads a never-built empty fixture.
        // Open() below is what turns that mistake into a readable failure instead of a silent one.
        [GlobalSetup(Targets = [nameof(Xlsx_ExcelReader), nameof(Xlsx_ExcelReader_Prefetch), nameof(Xlsx_Sylvan), nameof(Xlsx_ExcelReader_Materialized)])]
        public async Task SetupXlsxAsync()
        {
            _xlsx = await StringHeavyWorkbookGenerator.BuildXlsxAsync(Rows);
        }

        [GlobalSetup(Targets = [nameof(Xlsb_ExcelReader), nameof(Xlsb_ExcelReader_Prefetch), nameof(Xlsb_Sylvan), nameof(Xlsb_ExcelReader_Materialized)])]
        public async Task SetupXlsbAsync()
        {
            _xlsb = await StringHeavyWorkbookGenerator.BuildXlsbAsync(Rows);
        }

        // An empty fixture means the benchmark was never registered in a Targets list above. ZipArchive
        // happens to surface that as "Central Directory corrupt", but a format that tolerates a
        // zero-length input would instead measure no work at all and publish a meaningless number —
        // so fail loudly here rather than trusting each reader to reject it.
        private static MemoryStream Open(byte[] fixture, string benchmark)
        {
            if (fixture.Length == 0)
            {
                throw new InvalidOperationException(
                    $"{benchmark} is absent from every [GlobalSetup(Targets = ...)] list, so its fixture was never built.");
            }
            return new MemoryStream(fixture, writable: false);
        }

        // --- XLSX ---

        [Benchmark(Baseline = true)]
        public long Xlsx_ExcelReader()
        {
            using MemoryStream ms = Open(_xlsx, nameof(Xlsx_ExcelReader));
            using XlsxReader reader = Excel.From(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsx_ExcelReader_Prefetch()
        {
            using MemoryStream ms = Open(_xlsx, nameof(Xlsx_ExcelReader_Prefetch));
            using XlsxReader reader = Excel.From(ms, options: _prefetchOptions);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsx_Sylvan()
        {
            using MemoryStream ms = Open(_xlsx, nameof(Xlsx_Sylvan));
            using ExcelDataReader reader = ExcelDataReader.Create(ms, ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions());
            return AccumulateSylvanExcel(reader);
        }

        // Matched-work counterpart to Xlsx_ExcelReader: materializes a string per cell like
        // Xlsx_Sylvan is forced to, instead of reading the zero-copy span.
        [Benchmark]
        public long Xlsx_ExcelReader_Materialized()
        {
            using MemoryStream ms = Open(_xlsx, nameof(Xlsx_ExcelReader_Materialized));
            using XlsxReader reader = Excel.From(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRowMaterialized(row); }
            return acc;
        }

        // --- XLSB ---

        [Benchmark]
        public long Xlsb_ExcelReader()
        {
            using MemoryStream ms = Open(_xlsb, nameof(Xlsb_ExcelReader));
            using XlsbReader reader = Excel.FromXlsb(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsb_ExcelReader_Prefetch()
        {
            using MemoryStream ms = Open(_xlsb, nameof(Xlsb_ExcelReader_Prefetch));
            using XlsbReader reader = Excel.FromXlsb(ms, options: _prefetchOptions);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsb_Sylvan()
        {
            using MemoryStream ms = Open(_xlsb, nameof(Xlsb_Sylvan));
            using ExcelDataReader reader = ExcelDataReader.Create(ms, ExcelWorkbookType.ExcelBinary, new ExcelDataReaderOptions());
            return AccumulateSylvanExcel(reader);
        }

        // Matched-work counterpart to Xlsb_ExcelReader — see Xlsx_ExcelReader_Materialized.
        [Benchmark]
        public long Xlsb_ExcelReader_Materialized()
        {
            using MemoryStream ms = Open(_xlsb, nameof(Xlsb_ExcelReader_Materialized));
            using XlsbReader reader = Excel.FromXlsb(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRowMaterialized(row); }
            return acc;
        }
    }
}
