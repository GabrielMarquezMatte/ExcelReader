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
            using var reader = Excel.From(ms);
            long acc = 0;
            foreach (Record rec in new ExcelParser<Record>().Parse(reader))
            {
                acc += Accumulate(rec);
            }
            return acc;
        }

        // Same as ExcelParserSync, but Name comes from a shared-strings workbook instead of inline
        // strings. Name is drawn from an 8-value Pool, so this is where the reader's shared-string dedup
        // cache actually engages: cell.GetString() (called by BuildStringParser, ColumnParserFactory.cs)
        // resolves 8 distinct string instances instead of allocating one per row — ExcelParserSync's
        // inline-string workbook can never hit that cache (CellValueSource.RowValues/RowBuffer, not
        // Shared), so its 1.59 MB Name-string allocation is a property of the benchmark's input shape,
        // not an inherent parser cost.
        [Benchmark]
        public long ExcelParserSyncSharedStrings()
        {
            using var ms = new MemoryStream(_workbookSharedStrings, writable: false);
            using var reader = Excel.From(ms);
            long acc = 0;
            foreach (Record rec in new ExcelParser<Record>().Parse(reader))
            {
                acc += Accumulate(rec);
            }
            return acc;
        }

        // Same workbook/columns as ExcelParserSync, but the target is a struct — proves out the
        // zero-per-row-allocation path: ExcelParser<T> binds columns via `ref TModel`
        // (ColumnParser<T>/RefAction<T,TProperty>, see Delegates.cs) all the way down, and Row/RowCell
        // are ref structs, so a struct T and a direct foreach never box or allocate a model per row.
        [Benchmark]
        public long ExcelParserStructSync()
        {
            using var ms = new MemoryStream(_workbook, writable: false);
            using var reader = Excel.From(ms);
            long acc = 0;
            foreach (RecordStruct rec in new ExcelParser<RecordStruct>().Parse(reader))
            {
                acc += Accumulate(rec);
            }
            return acc;
        }

        // Reflection/attribute-driven ref-struct parse (ExcelReader.Core.Parser.RefParser.ParseNamed) —
        // same workbook/columns, matched by header name instead of ExcelParser<T>'s reflection-built
        // property setters. Name binds via ColumnParserFactory's span parser (zero-copy Cell.Value),
        // not GetString(), so this measures the fully-zero-alloc path, not just the container saving
        // ExcelParserStructSync/RecordStruct already showed.
        [Benchmark]
        public long RefParserParseNamedSync()
        {
            using var ms = new MemoryStream(_workbook, writable: false);
            using var reader = Excel.From(ms);
            long acc = 0;
            foreach (RecordNamedRef rec in RefParser.ParseNamed<RecordNamedRef>(reader))
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
        public long ExcelParserXlsbSync()
        {
            using var ms = new MemoryStream(_xlsbWorkbook, writable: false);
            using var reader = Excel.FromXlsb(ms);
            long acc = 0;
            foreach (Record rec in new ExcelParser<Record>().Parse(reader))
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
