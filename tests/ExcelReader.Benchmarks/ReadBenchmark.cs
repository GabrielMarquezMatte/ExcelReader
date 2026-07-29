using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;
using MiniExcelLibs;
using Sylvan.Data.Excel;
using static ExcelReader.Benchmarks.BenchmarkAccumulators;

namespace ExcelReader.Benchmarks
{
    // Reads a headerless workbook cell-by-cell, accumulating a checksum, across
    // ExcelReader (sync + async), MiniExcel, and Sylvan.
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
            using var reader = Excel.From(ms);
            long acc = 0;
            foreach (var row in reader)
            {
                foreach (var rowCell in row.Cells)
                {
                    var cell = rowCell.Value;
                    switch (cell.Type)
                    {
                        case CellType.ExcelString:
                            acc += cell.Value.Length;
                            break;
                        case CellType.Number:
                            if (cell.TryParse(null, out double n)) { acc += (long)n; }
                            break;
                        case CellType.Date:
                            if (cell.TryGetDateTime(out var d)) { acc += d.Ticks; }
                            break;
                        default:
                            break;
                    }
                }
            }
            return acc;
        }

        [Benchmark]
        public async Task<long> ExcelReaderAsync()
        {
            await using var ms = new MemoryStream(_workbook, writable: false);
            await using var reader = await Excel.FromAsync(ms);
            await using var e = await reader.GetAsyncEnumeratorAsync();
            long acc = 0;
            while (await e.MoveNextAsync())
            {
                var row = e.Current;
                foreach (var rowCell in row.Cells)
                {
                    var cell = rowCell.Value;
                    switch (cell.Type)
                    {
                        case CellType.ExcelString:
                            acc += cell.Value.Length;
                            break;
                        case CellType.Number:
                            if (cell.TryParse(null, out double n)) { acc += (long)n; }
                            break;
                        case CellType.Date:
                            if (cell.TryGetDateTime(out var d)) { acc += d.Ticks; }
                            break;
                        default:
                            break;
                    }
                }
            }
            return acc;
        }


        [Benchmark]
        public long ExcelReaderXlsb()
        {
            using var ms = new MemoryStream(_xlsbWorkbook, writable: false);
            using var reader = Excel.FromXlsb(ms);
            long acc = 0;
            foreach (var row in reader)
            {
                foreach (var rowCell in row.Cells)
                {
                    var cell = rowCell.Value;
                    switch (cell.Type)
                    {
                        case CellType.ExcelString:
                            acc += cell.Value.Length;
                            break;
                        case CellType.Number:
                            if (cell.TryParse(null, out double n)) { acc += (long)n; }
                            break;
                        case CellType.Date:
                            if (cell.TryGetDateTime(out var d)) { acc += d.Ticks; }
                            break;
                        default:
                            break;
                    }
                }
            }
            return acc;
        }

        [Benchmark]
        public async Task<long> ExcelReaderXlsbAsync()
        {
            await using var ms = new MemoryStream(_xlsbWorkbook, writable: false);
            await using var reader = await Excel.FromXlsbAsync(ms);
            await using var e = await reader.GetAsyncEnumeratorAsync();
            long acc = 0;
            while (await e.MoveNextAsync())
            {
                var row = e.Current;
                foreach (var rowCell in row.Cells)
                {
                    var cell = rowCell.Value;
                    switch (cell.Type)
                    {
                        case CellType.ExcelString:
                            acc += cell.Value.Length;
                            break;
                        case CellType.Number:
                            if (cell.TryParse(null, out double n)) { acc += (long)n; }
                            break;
                        case CellType.Date:
                            if (cell.TryGetDateTime(out var d)) { acc += d.Ticks; }
                            break;
                        default:
                            break;
                    }
                }
            }
            return acc;
        }

        // Matched-work counterpart to ExcelReader: materializes a string per cell like Sylvan's
        // GetString below is forced to, instead of reading the zero-copy span (see the README's
        // "Benchmark methodology" note).
        [Benchmark]
        public long ExcelReaderMaterialized()
        {
            using var ms = new MemoryStream(_workbook, writable: false);
            using var reader = Excel.From(ms);
            long acc = 0;
            foreach (var row in reader) { acc += AccumulateRowMaterialized(row); }
            return acc;
        }

        [Benchmark]
        public long MiniExcel()
        {
            using var ms = new MemoryStream(_workbook, writable: false);
            long acc = 0;
            foreach (var row in ms.Query(useHeaderRow: false, excelType: ExcelType.XLSX))
            {
                var r = (IDictionary<string, object?>)row;
                foreach (var val in r.Values)
                {
                    switch (val)
                    {
                        case string s: acc += s.Length; break;
                        case double d: acc += (long)d; break;
                        case DateTime dt: acc += dt.Ticks; break;
                    }
                }
            }
            return acc;
        }

        [Benchmark]
        public long Sylvan()
        {
            using var ms = new MemoryStream(_workbook, writable: false);
            using var reader = ExcelDataReader.Create(ms, ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions());
            long acc = 0;
            do
            {
                while (reader.Read())
                {
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        if (reader.IsDBNull(i)) { continue; }
                        switch (reader.GetExcelDataType(i))
                        {
                            case ExcelDataType.String:
                                acc += reader.GetString(i).Length;
                                break;
                            case ExcelDataType.Numeric:
                                acc += (long)reader.GetDouble(i);
                                break;
                            case ExcelDataType.DateTime:
                                acc += reader.GetDateTime(i).Ticks;
                                break;
                            case ExcelDataType.Boolean:
                            case ExcelDataType.Error:
                            case ExcelDataType.Null:
                                break;
                        }
                    }
                }
            }
            while (reader.NextResult());
            return acc;
        }
    }
}
