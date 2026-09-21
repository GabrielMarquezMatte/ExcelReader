using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

namespace ExcelReader.Benchmarks
{
    [MemoryDiagnoser]
    [SimpleJob(RunStrategy.ColdStart, launchCount: 16, warmupCount: 0, iterationCount: 1, invocationCount: 1)]
    public class ColdStartBenchmark
    {
        [Params(200)]
        public int Rows { get; set; }

        private List<Record> _records = [];
        private byte[] _workbook = [];

        [GlobalSetup]
        public async Task SetupAsync()
        {
            _records = WorkbookGenerator.Records(Rows);
            _workbook = await BuildTypedLowLevelAsync(_records);
        }

        private static async Task<byte[]> BuildTypedLowLevelAsync(List<Record> records)
        {
            await using var ms = new MemoryStream();
            await using (XlsxWorkbookWriter wb = await XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true))
            {
                await wb.StartAsync();
                XlsxSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync();
                await using (XlsxRowWriter header = await sheet.StartRowAsync())
                {
                    header.Write("Name");
                    header.Write("Id");
                    header.Write("Date");
                    header.Write("Value");
                }
                foreach (Record rec in records)
                {
                    await using XlsxRowWriter row = await sheet.StartRowAsync();
                    row.Write(rec.Name);
                    row.Write(rec.Id);
                    row.Write(rec.Date);
                    row.Write(rec.Value);
                }
                await sheet.EndAsync();
                await wb.EndAsync();
            }
            return ms.ToArray();
        }

        [Benchmark]
        public long TypedParseFirstUse()
        {
            using var ms = new MemoryStream(_workbook, writable: false);
            using var reader = Excel.FromXlsx(ms);
            long acc = 0;
            foreach (Record rec in new ExcelParser<Record>().Parse(reader))
            {
                acc += rec.Id;
            }
            return acc;
        }

        [Benchmark]
        public async Task<long> RecordWriteFirstUse()
        {
            await using var ms = new MemoryStream(64 * 1024);
            await using (var writer = await RecordWriter.CreateXlsxAsync(ms, leaveOpen: true))
            {
                await writer.WriteSheetAsync("S1", _records);
            }
            return ms.Length;
        }

        [Benchmark]
        public long FluentParseFirstUse()
        {
            using var ms = new MemoryStream(_workbook, writable: false);
            using var reader = Excel.FromXlsx(ms);
            var parser = new ExcelFluentParser<Record>(static builder => builder
                .Factory(static () => new Record())
                .Property(["Name"], ExcelCellReaders.String, static (ref r, v) => r.Name = v)
                .Property(["Id"], ExcelCellReaders.Parsable, static (ref Record r, int v) => r.Id = v)
                .Property(["Date"], ExcelCellReaders.DateTimeSerial, static (ref r, v) => r.Date = v)
                .Property(["Value"], ExcelCellReaders.Parsable, static (ref Record r, double v) => r.Value = v));
            long acc = 0;
            foreach (Record rec in parser.Parse(reader))
            {
                acc += rec.Id;
            }
            return acc;
        }

        [Benchmark]
        public long FluentParseWithAttributeFallbackFirstUse()
        {
            using var ms = new MemoryStream(_workbook, writable: false);
            using var reader = Excel.FromXlsx(ms);
            ExcelFluentParser<Record> parser = ExcelFluentParser<Record>.WithAttributeFallback(static builder => builder
                .Property(["Id"], ExcelCellReaders.Parsable, static (ref Record r, int v) => r.Id = v));
            long acc = 0;
            foreach (Record rec in parser.Parse(reader))
            {
                acc += rec.Id;
            }
            return acc;
        }
    }
}
