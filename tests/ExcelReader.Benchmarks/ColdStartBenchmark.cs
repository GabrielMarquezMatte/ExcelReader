using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

namespace ExcelReader.Benchmarks
{
    // First-use (cold) cost of the reflection + Expression.Compile machinery behind the typed parser and
    // the record writer. RunStrategy.ColdStart with launchCount>1 runs each benchmark in a fresh process
    // with no warmup, so every measured invocation pays JIT compilation plus the one-time per-type setter/
    // column-writer compilation — the startup cost that steady-state throughput benchmarks warm away.
    // Rows is deliberately small so the fixed compile cost dominates the per-row work.
    //
    // GlobalSetup builds the parse workbook through the LOW-LEVEL writer (XlsxRowWriter), never through
    // ExcelParser<T> or the record writer, so neither TypeMapper<Record> (parser) nor RecordColumns<Record>
    // (writer) is warmed before the measured invocation.
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

        // Header + one row per record via the low-level cell writer — the same bytes RecordWriter would
        // produce, but built without touching the record-writer's compiled column plan.
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

        // Cold cost of the first typed parse: JIT + TypeMapper<Record> reflection and setter compilation.
        [Benchmark]
        public long TypedParseFirstUse()
        {
            using var ms = new MemoryStream(_workbook, writable: false);
            using var reader = Excel.From(ms);
            long acc = 0;
            foreach (Record rec in new ExcelParser<Record>().Parse(reader))
            {
                acc += rec.Id;
            }
            return acc;
        }

        // Cold cost of the first record write: JIT + RecordColumns<Record> column-writer compilation.
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

        // Cold cost of the plain ExcelFluentParser<T> constructor: no reflection at all — `configure`
        // only allocates delegates and a PropertyMap<T>[] via ExcelRowMapBuilder<T>. Contrast with
        // TypedParseFirstUse to check the "no reflection => faster cold start" claim empirically instead
        // of by heuristic.
        [Benchmark]
        public long FluentParseFirstUse()
        {
            using var ms = new MemoryStream(_workbook, writable: false);
            using var reader = Excel.From(ms);
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

        // Cold cost of ExcelFluentParser<T>.WithAttributeFallback: it reflects via TypeMapper<T>.GetInfo()
        // for the fallback half, so it pays the same reflection/setter-compilation cost as
        // TypedParseFirstUse, plus the fluent build and merge on top.
        [Benchmark]
        public long FluentParseWithAttributeFallbackFirstUse()
        {
            using var ms = new MemoryStream(_workbook, writable: false);
            using var reader = Excel.From(ms);
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
