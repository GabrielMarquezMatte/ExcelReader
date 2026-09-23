using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using OfficeIMO.Excel;
using Sylvan.Data.Excel;
using static ExcelReader.Benchmarks.BenchmarkAccumulators;

namespace ExcelReader.Benchmarks
{
    [MemoryDiagnoser]
    public class ReadBenchmark
    {
        [Params(50_000)]
        public int Rows { get; set; }

        private byte[] _workbook = [];
        private byte[] _xlsbWorkbook = [];

        [GlobalSetup]
        public async Task SetupAsync()
        {
            _workbook = await WorkbookGenerator.BuildAsync(Rows);
            _xlsbWorkbook = await WorkbookGenerator.BuildXlsbAsync(Rows);
        }

        [Benchmark(Baseline = true)]
        public long ExcelReader()
        {
            using var ms = new MemoryStream(_workbook, writable: false);
            using var reader = Excel.FromXlsx(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public async Task<long> ExcelReaderAsync()
        {
            await using var ms = new MemoryStream(_workbook, writable: false);
            await using var reader = await Excel.FromXlsxAsync(ms);
            await using var e = reader.GetAsyncEnumerator();
            long acc = 0;
            while (await e.MoveNextAsync()) { acc += AccumulateRow(e.Current); }
            return acc;
        }

        [Benchmark]
        public long ExcelReaderXlsb()
        {
            using var ms = new MemoryStream(_xlsbWorkbook, writable: false);
            using var reader = Excel.FromXlsb(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public async Task<long> ExcelReaderXlsbAsync()
        {
            await using var ms = new MemoryStream(_xlsbWorkbook, writable: false);
            await using var reader = await Excel.FromXlsbAsync(ms);
            await using var e = reader.GetAsyncEnumerator();
            long acc = 0;
            while (await e.MoveNextAsync()) { acc += AccumulateRow(e.Current); }
            return acc;
        }

        [Benchmark]
        public long ExcelReaderMaterialized()
        {
            using var ms = new MemoryStream(_workbook, writable: false);
            using var reader = Excel.FromXlsx(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRowMaterialized(row); }
            return acc;
        }

        [Benchmark]
        public long Sylvan()
        {
            using var ms = new MemoryStream(_workbook, writable: false);
            using var reader = global::Sylvan.Data.Excel.ExcelDataReader.Create(ms, ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions());
            return AccumulateSylvanExcel(reader);
        }

        [Benchmark]
        public long OfficeIMO()
        {
            using var reader = ExcelDocument.OpenDataReader(_workbook, new ExcelReadOptions { HasHeaderRow = false });
            return AccumulateDataReader(reader);
        }
    }
}
