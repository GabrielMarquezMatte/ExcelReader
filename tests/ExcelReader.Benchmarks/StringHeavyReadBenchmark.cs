using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using Sylvan.Data.Excel;

namespace ExcelReader.Benchmarks
{
    // Reads StringHeavyWorkbookGenerator's fixture cell-by-cell: 8 text columns against 3
    // numeric/date columns, tens of thousands of distinct shared strings. Companion to
    // RealDataReadBenchmark, which covers the 65K_Records_Data.* corpus — that corpus has only
    // 5 KB of shared strings across ~910K cells, so it never exercises the shared-string cache,
    // dictionary lookups or string materialization that prefetch and parsing changes touch. This
    // class isolates that axis: same [MemoryDiagnoser]/AccumulateRow shape, prefetch on and off,
    // for both ZIP-based formats (xlsx, xlsb) prefetch actually affects.
    [MemoryDiagnoser]
    public class StringHeavyReadBenchmark
    {
        [Params(65_536)]
        public int Rows { get; set; }

        private static readonly ExcelReaderOptions _prefetchOptions = new() { PrefetchDecompression = true };

        private byte[] _xlsx = [];
        private byte[] _xlsb = [];

        [GlobalSetup]
        public async Task SetupAsync()
        {
            _xlsx = await StringHeavyWorkbookGenerator.BuildXlsxAsync(Rows);
            _xlsb = await StringHeavyWorkbookGenerator.BuildXlsbAsync(Rows);
        }

        private static long AccumulateRow(Row row)
        {
            long acc = 0;
            foreach (RowCell rowCell in row.Cells)
            {
                Cell cell = rowCell.Value;
                switch (cell.Type)
                {
                    case CellType.ExcelString:
                        acc += cell.Value.Length;
                        break;
                    case CellType.Number:
                        if (cell.TryParse(null, out double n)) { acc += (long)n; }
                        break;
                    case CellType.Date:
                        if (cell.TryGetDateTime(out DateTime d)) { acc += d.Ticks; }
                        break;
                }
            }
            return acc;
        }

        // --- XLSX ---

        [Benchmark(Baseline = true)]
        public long Xlsx_ExcelReader()
        {
            using MemoryStream ms = new(_xlsx, writable: false);
            using XlsxReader reader = Excel.From(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsx_ExcelReader_Prefetch()
        {
            using MemoryStream ms = new(_xlsx, writable: false);
            using XlsxReader reader = Excel.From(ms, options: _prefetchOptions);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsx_Sylvan()
        {
            using MemoryStream ms = new(_xlsx, writable: false);
            using ExcelDataReader reader = ExcelDataReader.Create(ms, ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions());
            return AccumulateSylvanExcel(reader);
        }

        // --- XLSB ---

        [Benchmark]
        public long Xlsb_ExcelReader()
        {
            using MemoryStream ms = new(_xlsb, writable: false);
            using XlsbReader reader = Excel.FromXlsb(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsb_ExcelReader_Prefetch()
        {
            using MemoryStream ms = new(_xlsb, writable: false);
            using XlsbReader reader = Excel.FromXlsb(ms, options: _prefetchOptions);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsb_Sylvan()
        {
            using MemoryStream ms = new(_xlsb, writable: false);
            using ExcelDataReader reader = ExcelDataReader.Create(ms, ExcelWorkbookType.ExcelBinary, new ExcelDataReaderOptions());
            return AccumulateSylvanExcel(reader);
        }

        private static long AccumulateSylvanExcel(ExcelDataReader reader)
        {
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
                        }
                    }
                }
            }
            while (reader.NextResult());
            return acc;
        }
    }
}
