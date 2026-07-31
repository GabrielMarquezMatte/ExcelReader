using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Writer;

namespace ExcelReader.Benchmarks
{
    // Writes `Rows` records via the high-level RecordWriter (WorkbookRecordWriter.WriteSheetAsync), the
    // POCO-dump API an application uses. Unlike WriteBenchmark (which drives the low-level cell writers
    // directly), this routes every numeric property through the generic Write<T> overload — the path
    // that resolves through ToDouble<T> for XLSB/XLS. It is the only benchmark that exercises that path,
    // so the MemoryDiagnoser allocation numbers here reflect any per-numeric-cell boxing in it.
    [MemoryDiagnoser]
    public class RecordWriteBenchmark
    {
        [Params(50_000)]
        public int Rows { get; set; }

        private List<Record> _records = [];

        [GlobalSetup]
        public void Setup()
        {
            _records = WorkbookGenerator.Records(Rows);
        }

        [Benchmark(Baseline = true)]
        public async Task<long> Xlsx()
        {
            await using var ms = new MemoryStream(4 * 1024 * 1024);
            await using (var writer = await RecordWriter.CreateXlsxAsync(ms, leaveOpen: true))
            {
                await writer.WriteSheetAsync("S1", _records);
            }
            return ms.Length;
        }

        [Benchmark]
        public async Task<long> Xlsb()
        {
            await using var ms = new MemoryStream(4 * 1024 * 1024);
            await using (var writer = await RecordWriter.CreateXlsbAsync(ms, leaveOpen: true))
            {
                await writer.WriteSheetAsync("S1", _records);
            }
            return ms.Length;
        }

        [Benchmark]
        public async Task<long> Xls()
        {
            await using var ms = new MemoryStream(4 * 1024 * 1024);
            await using (var writer = await RecordWriter.CreateXlsAsync(ms, leaveOpen: true))
            {
                await writer.WriteSheetAsync("S1", _records);
            }
            return ms.Length;
        }

        [Benchmark]
        public async Task<long> Csv()
        {
            await using var ms = new MemoryStream(4 * 1024 * 1024);
            await using (var writer = await RecordWriter.CreateCsvAsync(ms, leaveOpen: true))
            {
                await writer.WriteSheetAsync("S1", _records);
            }
            return ms.Length;
        }
    }
}
