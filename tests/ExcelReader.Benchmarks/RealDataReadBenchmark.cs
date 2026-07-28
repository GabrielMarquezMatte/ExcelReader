using System.Globalization;
using System.Text;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using Sylvan.Data.Csv;
using Sylvan.Data.Excel;
using static ExcelReader.Benchmarks.BenchmarkAccumulators;

namespace ExcelReader.Benchmarks
{
    // Reads the real-world 65K-row dataset (Data/65K_Records_Data.*) cell-by-cell across every
    // format this library supports, against Sylvan — the only other library referenced here that
    // also reads xls/xlsb/xlsx/xlsm/csv. Unlike the synthetic *ReadBenchmark classes (headerless,
    // 4 generated columns), this exercises a real file: 14 columns, a header row, real compression
    // ratios, and real shared-string/date-style density.
    [MemoryDiagnoser]
    public class RealDataReadBenchmark
    {
        private static readonly ExcelReaderOptions _prefetchOptions = new() { PrefetchDecompression = true };

        private byte[] _xlsx = [];
        private byte[] _xlsm = [];
        private byte[] _xlsb = [];
        private byte[] _xls = [];
        private byte[] _csv = [];

        [GlobalSetup]
        public void Setup()
        {
            // Sylvan decodes legacy .xls text as CP1252, which .NET only exposes via this provider.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            string dir = Path.Combine(AppContext.BaseDirectory, "Data");
            _xlsx = File.ReadAllBytes(Path.Combine(dir, "65K_Records_Data.xlsx"));
            _xlsm = File.ReadAllBytes(Path.Combine(dir, "65K_Records_Data.xlsm"));
            _xlsb = File.ReadAllBytes(Path.Combine(dir, "65K_Records_Data.xlsb"));
            _xls = File.ReadAllBytes(Path.Combine(dir, "65K_Records_Data.xls"));
            _csv = File.ReadAllBytes(Path.Combine(dir, "65K_Records_Data.csv"));
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
        public long Xlsx_Sylvan()
        {
            using MemoryStream ms = new(_xlsx, writable: false);
            using ExcelDataReader reader = ExcelDataReader.Create(ms, ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions());
            return AccumulateSylvanExcel(reader);
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

        // --- XLSM (same OOXML container as XLSX; ExcelReader parses it identically) ---

        [Benchmark]
        public long Xlsm_ExcelReader()
        {
            using MemoryStream ms = new(_xlsm, writable: false);
            using XlsxReader reader = Excel.From(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsm_Sylvan()
        {
            using MemoryStream ms = new(_xlsm, writable: false);
            using ExcelDataReader reader = ExcelDataReader.Create(ms, ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions());
            return AccumulateSylvanExcel(reader);
        }

        [Benchmark]
        public long Xlsm_ExcelReader_Prefetch()
        {
            using MemoryStream ms = new(_xlsm, writable: false);
            using XlsxReader reader = Excel.From(ms, options: _prefetchOptions);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
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
        public long Xlsb_Sylvan()
        {
            using MemoryStream ms = new(_xlsb, writable: false);
            using ExcelDataReader reader = ExcelDataReader.Create(ms, ExcelWorkbookType.ExcelBinary, new ExcelDataReaderOptions());
            return AccumulateSylvanExcel(reader);
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

        // --- XLS ---

        [Benchmark]
        public long Xls_ExcelReader()
        {
            using MemoryStream ms = new(_xls, writable: false);
            using XlsReader reader = Excel.FromXls(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public long Xls_Sylvan()
        {
            using MemoryStream ms = new(_xls, writable: false);
            using ExcelDataReader reader = ExcelDataReader.Create(ms, ExcelWorkbookType.Excel, new ExcelDataReaderOptions());
            return AccumulateSylvanExcel(reader);
        }

        // --- CSV --- (CSV cells are always plain text on both sides — no style-driven date/number
        // typing is possible — so both benchmarks just sum text length for a like-for-like comparison.)

        [Benchmark]
        public long Csv_ExcelReader()
        {
            using MemoryStream ms = new(_csv, writable: false);
            using CsvReader reader = Excel.FromCsv(ms);
            long acc = 0;
            foreach (Row row in reader)
            {
                foreach (RowCell rowCell in row.Cells)
                {
                    acc += rowCell.Value.Value.Length;
                }
            }
            return acc;
        }

        [Benchmark]
        public long Csv_Sylvan()
        {
            using MemoryStream ms = new(_csv, writable: false);
            using StreamReader tr = new(ms);
            CsvDataReaderOptions options = new() { HasHeaders = false, Culture = CultureInfo.InvariantCulture };
            using CsvDataReader reader = CsvDataReader.Create(tr, options);
            long acc = 0;
            while (reader.Read())
            {
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    acc += reader.GetString(i).Length;
                }
            }
            return acc;
        }
    }
}
