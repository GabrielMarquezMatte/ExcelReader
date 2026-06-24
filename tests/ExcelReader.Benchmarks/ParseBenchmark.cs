using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using MiniExcelLibs;
using Sylvan.Data;
using Sylvan.Data.Excel;

namespace ExcelReader.Benchmarks
{
    // Maps a header + `Rows` data rows into strongly-typed Record objects,
    // comparing ExcelReader's ExcelParser<T> (sync + async) against MiniExcel.
    [MemoryDiagnoser]
    public class ParseBenchmark
    {
        [Params(50_000)]
        public int Rows { get; set; }

        private byte[] _workbook = [];

        [GlobalSetup]
        public async Task SetupAsync()
        {
            _workbook = await WorkbookGenerator.BuildTypedAsync(Rows);
        }

        private static long Accumulate(Record rec)
        {
            return rec.Id + (long)rec.Value + (rec.Name?.Length ?? 0) + rec.Date.Ticks;
        }


        [Benchmark(Baseline = true)]
        public long ExcelParserSync()
        {
            using var ms = new MemoryStream(_workbook, writable: false);
            using var reader = Excel.From(ms);
            long acc = 0;
            foreach (Record rec in new ExcelParser<Record>().Parse(reader))
            {
                acc += Accumulate(rec);
            }
            return acc;
        }

        [Benchmark]
        public async Task<long> ExcelParserAsync()
        {
            await using var ms = new MemoryStream(_workbook, writable: false);
            await using var reader = await Excel.FromAsync(ms);
            long acc = 0;
            await foreach (Record rec in new ExcelParser<Record>().ParseAsync(reader))
            {
                acc += Accumulate(rec);
            }
            return acc;
        }

        [Benchmark]
        public long MiniExcel()
        {
            using var ms = new MemoryStream(_workbook, writable: false);
            long acc = 0;
            foreach (Record rec in ms.Query<Record>(excelType: ExcelType.XLSX))
            {
                acc += Accumulate(rec);
            }
            return acc;
        }

        [Benchmark]
        public long Sylvan()
        {
            using var ms = new MemoryStream(_workbook, writable: false);
            using var reader = ExcelDataReader.Create(ms, ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions());
            long acc = 0;
            foreach (Record rec in reader.GetRecords<Record>())
            {
                acc += Accumulate(rec);
            }
            return acc;
        }

        [Benchmark]
        public async Task<long> SylvanAsync()
        {
            await using var ms = new MemoryStream(_workbook, writable: false);
            await using var reader = await ExcelDataReader.CreateAsync(ms, ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions()).ConfigureAwait(false);
            long acc = 0;
            await foreach (Record rec in reader.GetRecordsAsync<Record>())
            {
                acc += Accumulate(rec);
            }
            return acc;
        }
    }
}
