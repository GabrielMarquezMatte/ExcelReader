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
        private List<MappedRecord> _mapped = [];

        [GlobalSetup]
        public void Setup()
        {
            _records = WorkbookGenerator.Records(Rows);
            _mapped = [.. _records.Select(static r => new MappedRecord
            {
                Name = r.Name,
                Id = r.Id,
                Date = r.Date,
                Value = r.Value,
            })];
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

        // The AOT-clean twins. Their column plan writes through Action<IRowWriter, T>, so every cell is
        // an interface call, where the reflection path above compiles Action<TRow, T> against the
        // concrete row writer. These pairs are what measures whether that difference costs anything.
        [Benchmark]
        public async Task<long> XlsxMapped()
        {
            await using var ms = new MemoryStream(4 * 1024 * 1024);
            await using (var writer = await MappedRecordWriter.CreateMappedXlsxAsync(ms, leaveOpen: true))
            {
                await writer.WriteSheetAsync("S1", _mapped);
            }
            return ms.Length;
        }

        [Benchmark]
        public async Task<long> XlsbMapped()
        {
            await using var ms = new MemoryStream(4 * 1024 * 1024);
            await using (var writer = await MappedRecordWriter.CreateMappedXlsbAsync(ms, leaveOpen: true))
            {
                await writer.WriteSheetAsync("S1", _mapped);
            }
            return ms.Length;
        }

        [Benchmark]
        public async Task<long> CsvMapped()
        {
            await using var ms = new MemoryStream(4 * 1024 * 1024);
            await using (var writer = await MappedRecordWriter.CreateMappedCsvAsync(ms, leaveOpen: true))
            {
                await writer.WriteSheetAsync("S1", _mapped);
            }
            return ms.Length;
        }
    }

    // Record's twin with a hand-written map, so the mapped benchmarks write the same four columns in
    // the same order as the reflection-driven ones.
    public sealed class MappedRecord : IExcelRecordMap<MappedRecord>
    {
        public string? Name { get; set; }
        public int Id { get; set; }
        public DateTime Date { get; set; }
        public double Value { get; set; }

        public static void ConfigureExcelRecordMap(ExcelRecordMapBuilder<MappedRecord> builder)
        {
            builder.Column("Name", static (row, r) => row.Write(r.Name))
                   .Column("Id", static (row, r) => row.Write(r.Id))
                   .Column("Date", static (row, r) => row.Write(r.Date))
                   .Column("Value", static (row, r) => row.Write(r.Value));
        }
    }
}
