using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using OfficeIMO.Excel;
using Sylvan.Data;
using Sylvan.Data.Excel;

namespace ExcelReader.Benchmarks
{
    [MemoryDiagnoser]
    public class ParseBenchmark
    {
        [Params(50_000)]
        public int Rows { get; set; }

        private byte[] _workbook = [];
        private byte[] _xlsbWorkbook = [];
        private byte[] _workbookSharedStrings = [];

        [GlobalSetup]
        public async Task SetupAsync()
        {
            _workbook = await WorkbookGenerator.BuildTypedAsync(Rows);
            _xlsbWorkbook = await WorkbookGenerator.BuildTypedXlsbAsync(Rows);
            _workbookSharedStrings = await WorkbookGenerator.BuildTypedSharedStringsAsync(Rows);
        }

        private static long Accumulate(Record rec)
        {
            return rec.Id + (long)rec.Value + (rec.Name?.Length ?? 0) + rec.Date.Ticks;
        }

        private static long Accumulate(RecordStruct rec)
        {
            return rec.Id + (long)rec.Value + (rec.Name?.Length ?? 0) + rec.Date.Ticks;
        }

        private static long Accumulate(RecordNamedRef rec)
        {
            return rec.Id + (long)rec.Value + rec.Name.Length + rec.Date.Ticks;
        }


        [Benchmark(Baseline = true)]
        public long ExcelParserSync()
        {
            using var ms = new MemoryStream(_workbook, writable: false);
            using var reader = Excel.FromXlsx(ms);
            long acc = 0;
            foreach (Record rec in ExcelParser.FromAttributes<Record>().Parse(reader))
            {
                acc += Accumulate(rec);
            }
            return acc;
        }

        [Benchmark]
        public long ExcelParserSyncSharedStrings()
        {
            using var ms = new MemoryStream(_workbookSharedStrings, writable: false);
            using var reader = Excel.FromXlsx(ms);
            long acc = 0;
            foreach (Record rec in ExcelParser.FromAttributes<Record>().Parse(reader))
            {
                acc += Accumulate(rec);
            }
            return acc;
        }

        [Benchmark]
        public long ExcelParserStructSync()
        {
            using var ms = new MemoryStream(_workbook, writable: false);
            using var reader = Excel.FromXlsx(ms);
            long acc = 0;
            foreach (RecordStruct rec in ExcelParser.FromAttributes<RecordStruct>().Parse(reader))
            {
                acc += Accumulate(rec);
            }
            return acc;
        }

        [Benchmark]
        public long RefStructModelSync()
        {
            using var ms = new MemoryStream(_workbook, writable: false);
            using var reader = Excel.FromXlsx(ms);
            long acc = 0;
            foreach (RecordNamedRef rec in ExcelParser.FromAttributes<RecordNamedRef>().Parse(reader))
            {
                acc += Accumulate(rec);
            }
            return acc;
        }

        [Benchmark]
        public async Task<long> ExcelParserAsync()
        {
            await using var ms = new MemoryStream(_workbook, writable: false);
            await using var reader = await Excel.FromXlsxAsync(ms);
            long acc = 0;
            await foreach (Record rec in ExcelParser.FromAttributes<Record>().Parse(reader))
            {
                acc += Accumulate(rec);
            }
            return acc;
        }


        [Benchmark]
        public long ExcelParserXlsbSync()
        {
            using var ms = new MemoryStream(_xlsbWorkbook, writable: false);
            using var reader = Excel.FromXlsb(ms);
            long acc = 0;
            foreach (Record rec in ExcelParser.FromAttributes<Record>().Parse(reader))
            {
                acc += Accumulate(rec);
            }
            return acc;
        }

        [Benchmark]
        public async Task<long> ExcelParserXlsbAsync()
        {
            await using var ms = new MemoryStream(_xlsbWorkbook, writable: false);
            await using var reader = await Excel.FromXlsbAsync(ms);
            long acc = 0;
            await foreach (Record rec in ExcelParser.FromAttributes<Record>().Parse(reader))
            {
                acc += Accumulate(rec);
            }
            return acc;
        }

        [Benchmark]
        public long OfficeIMO()
        {
            // Hand-mapped: OfficeIMO's RowsAs<T> needs an r attribute on every <row> and <c>, which the spec makes optional and our writer omits.
            using var reader = ExcelDocument.OpenDataReader(_workbook, new ExcelReadOptions { HasHeaderRow = true });
            long acc = 0;
            while (reader.Read())
            {
                acc += Accumulate(new Record
                {
                    Name = reader.GetString(0),
                    Id = reader.GetInt32(1),
                    Date = reader.GetDateTime(2),
                    Value = reader.GetDouble(3),
                });
            }
            return acc;
        }

        [Benchmark]
        public long Sylvan()
        {
            using var ms = new MemoryStream(_workbook, writable: false);
            using var reader = global::Sylvan.Data.Excel.ExcelDataReader.Create(ms, ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions());
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
            await using var reader = await global::Sylvan.Data.Excel.ExcelDataReader.CreateAsync(ms, ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions()).ConfigureAwait(false);
            long acc = 0;
            await foreach (Record rec in reader.GetRecordsAsync<Record>())
            {
                acc += Accumulate(rec);
            }
            return acc;
        }
    }
}
