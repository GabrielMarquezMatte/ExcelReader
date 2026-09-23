using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Writer;

namespace ExcelReader.Benchmarks
{
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
            await using (var writer = RecordWriter.CreateXlsx(ms, leaveOpen: true))
            {
                await writer.WriteSheetAsync("S1", _records);
            }
            return ms.Length;
        }

        [Benchmark]
        public async Task<long> Xlsb()
        {
            await using var ms = new MemoryStream(4 * 1024 * 1024);
            await using (var writer = RecordWriter.CreateXlsb(ms, leaveOpen: true))
            {
                await writer.WriteSheetAsync("S1", _records);
            }
            return ms.Length;
        }

        [Benchmark]
        public async Task<long> Xls()
        {
            await using var ms = new MemoryStream(4 * 1024 * 1024);
            await using (var writer = RecordWriter.CreateXls(ms, leaveOpen: true))
            {
                await writer.WriteSheetAsync("S1", _records);
            }
            return ms.Length;
        }

        [Benchmark]
        public async Task<long> Csv()
        {
            await using var ms = new MemoryStream(4 * 1024 * 1024);
            await using (var writer = RecordWriter.CreateCsv(ms, leaveOpen: true))
            {
                await writer.WriteSheetAsync("S1", _records);
            }
            return ms.Length;
        }

        [Benchmark]
        public async Task<long> XlsxMapped()
        {
            await using var ms = new MemoryStream(4 * 1024 * 1024);
            await using (var writer = MappedRecordWriter.CreateXlsx(ms, leaveOpen: true))
            {
                await writer.WriteSheetAsync("S1", _mapped);
            }
            return ms.Length;
        }

        [Benchmark]
        public async Task<long> XlsbMapped()
        {
            await using var ms = new MemoryStream(4 * 1024 * 1024);
            await using (var writer = MappedRecordWriter.CreateXlsb(ms, leaveOpen: true))
            {
                await writer.WriteSheetAsync("S1", _mapped);
            }
            return ms.Length;
        }

        [Benchmark]
        public async Task<long> CsvMapped()
        {
            await using var ms = new MemoryStream(4 * 1024 * 1024);
            await using (var writer = MappedRecordWriter.CreateCsv(ms, leaveOpen: true))
            {
                await writer.WriteSheetAsync("S1", _mapped);
            }
            return ms.Length;
        }
    }

    public sealed class MappedRecord : IExcelRecordMap<MappedRecord>
    {
        public string? Name { get; set; }
        public int Id { get; set; }
        public DateTime Date { get; set; }
        public double Value { get; set; }

        public static void ConfigureExcelRecordMap<TRow>(ExcelRecordMapBuilder<MappedRecord, TRow> builder)
            where TRow : IRowWriter
        {
            builder.Column("Name", static (row, r) => row.Write(r.Name))
                   .Column("Id", static (row, r) => row.Write(r.Id))
                   .Column("Date", static (row, r) => row.Write(r.Date))
                   .Column("Value", static (row, r) => row.Write(r.Value));
        }
    }
}
