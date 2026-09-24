using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Reader.Xlsx;
using Sylvan.Data.Excel;
using static ExcelReader.Benchmarks.BenchmarkAccumulators;

namespace ExcelReader.Benchmarks
{
    [MemoryDiagnoser]
    public class StringHeavyReadBenchmark
    {
        [Params(65_536)]
        public int Rows { get; set; }

        private static readonly ExcelReaderOptions _prefetchOptions = new() { PrefetchDecompression = true };
        private static readonly ExcelReaderOptions _internOptions = new() { InternStrings = true };

        private byte[] _xlsx = [];
        private byte[] _xlsb = [];

        [GlobalSetup(Targets = [nameof(Xlsx_ExcelReader), nameof(Xlsx_ExcelReader_Prefetch), nameof(Xlsx_Sylvan), nameof(Xlsx_ExcelReader_Materialized), nameof(Xlsx_ExcelReader_Materialized_Interned)])]
        public async Task SetupXlsxAsync()
        {
            _xlsx = await StringHeavyWorkbookGenerator.BuildXlsxAsync(Rows);
        }

        [GlobalSetup(Targets = [nameof(Xlsb_ExcelReader), nameof(Xlsb_ExcelReader_Prefetch), nameof(Xlsb_Sylvan), nameof(Xlsb_ExcelReader_Materialized), nameof(Xlsb_ExcelReader_Materialized_Interned), nameof(Xlsb_ExcelReader_Memory)])]
        public async Task SetupXlsbAsync()
        {
            _xlsb = await StringHeavyWorkbookGenerator.BuildXlsbAsync(Rows);
        }

        private static MemoryStream Open(byte[] fixture, string benchmark)
        {
            if (fixture.Length == 0)
            {
                throw new InvalidOperationException(
                    $"{benchmark} is absent from every [GlobalSetup(Targets = ...)] list, so its fixture was never built.");
            }
            return new MemoryStream(fixture, writable: false);
        }


        [Benchmark(Baseline = true)]
        public long Xlsx_ExcelReader()
        {
            using MemoryStream ms = Open(_xlsx, nameof(Xlsx_ExcelReader));
            using XlsxReader reader = Excel.FromXlsx(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsx_ExcelReader_Prefetch()
        {
            using MemoryStream ms = Open(_xlsx, nameof(Xlsx_ExcelReader_Prefetch));
            using XlsxReader reader = Excel.FromXlsx(ms, options: _prefetchOptions);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsx_Sylvan()
        {
            using MemoryStream ms = Open(_xlsx, nameof(Xlsx_Sylvan));
            using Sylvan.Data.Excel.ExcelDataReader reader = Sylvan.Data.Excel.ExcelDataReader.Create(ms, ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions());
            return AccumulateSylvanExcel(reader);
        }

        [Benchmark]
        public long Xlsx_ExcelReader_Materialized()
        {
            using MemoryStream ms = Open(_xlsx, nameof(Xlsx_ExcelReader_Materialized));
            using XlsxReader reader = Excel.FromXlsx(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRowMaterialized(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsx_ExcelReader_Materialized_Interned()
        {
            using MemoryStream ms = Open(_xlsx, nameof(Xlsx_ExcelReader_Materialized_Interned));
            using XlsxReader reader = Excel.FromXlsx(ms, options: _internOptions);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRowMaterialized(row); }
            return acc;
        }


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
            using Sylvan.Data.Excel.ExcelDataReader reader = Sylvan.Data.Excel.ExcelDataReader.Create(ms, ExcelWorkbookType.ExcelBinary, new ExcelDataReaderOptions());
            return AccumulateSylvanExcel(reader);
        }

        [Benchmark]
        public long Xlsb_ExcelReader_Materialized()
        {
            using MemoryStream ms = Open(_xlsb, nameof(Xlsb_ExcelReader_Materialized));
            using XlsbReader reader = Excel.FromXlsb(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRowMaterialized(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsb_ExcelReader_Materialized_Interned()
        {
            using MemoryStream ms = Open(_xlsb, nameof(Xlsb_ExcelReader_Materialized_Interned));
            using XlsbReader reader = Excel.FromXlsb(ms, options: _internOptions);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRowMaterialized(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsb_ExcelReader_Memory()
        {
            if (_xlsb.Length == 0)
            {
                throw new InvalidOperationException($"{nameof(Xlsb_ExcelReader_Memory)} is absent from every [GlobalSetup(Targets = ...)] list, so its fixture was never built.");
            }
            using XlsbReader reader = Excel.FromXlsb(new ReadOnlyMemory<byte>(_xlsb));
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }
    }
}
