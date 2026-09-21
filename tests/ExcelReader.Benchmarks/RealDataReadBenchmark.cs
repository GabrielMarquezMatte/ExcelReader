using System.Globalization;
using System.Text;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using Sylvan.Data.Csv;
using Sylvan.Data.Excel;
using static ExcelReader.Benchmarks.BenchmarkAccumulators;

namespace ExcelReader.Benchmarks
{
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
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            string dir = Path.Combine(AppContext.BaseDirectory, "Data");
            _xlsx = File.ReadAllBytes(Path.Combine(dir, "65K_Records_Data.xlsx"));
            _xlsm = File.ReadAllBytes(Path.Combine(dir, "65K_Records_Data.xlsm"));
            _xlsb = File.ReadAllBytes(Path.Combine(dir, "65K_Records_Data.xlsb"));
            _xls = File.ReadAllBytes(Path.Combine(dir, "65K_Records_Data.xls"));
            _csv = File.ReadAllBytes(Path.Combine(dir, "65K_Records_Data.csv"));
        }


        [Benchmark(Baseline = true)]
        public long Xlsx_ExcelReader()
        {
            using MemoryStream ms = new(_xlsx, writable: false);
            using XlsxReader reader = Excel.FromXlsx(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsx_Sylvan()
        {
            using MemoryStream ms = new(_xlsx, writable: false);
            using Sylvan.Data.Excel.ExcelDataReader reader = Sylvan.Data.Excel.ExcelDataReader.Create(ms, ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions());
            return AccumulateSylvanExcel(reader);
        }

        [Benchmark]
        public long Xlsx_ExcelReader_Materialized()
        {
            using MemoryStream ms = new(_xlsx, writable: false);
            using XlsxReader reader = Excel.FromXlsx(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRowMaterialized(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsx_ExcelReader_Prefetch()
        {
            using MemoryStream ms = new(_xlsx, writable: false);
            using XlsxReader reader = Excel.FromXlsx(ms, options: _prefetchOptions);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsx_ExcelReader_Memory()
        {
            using XlsxReader reader = Excel.FromXlsx(_xlsx.AsMemory());
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsx_ExcelReader_Memory_Prefetch()
        {
            using XlsxReader reader = Excel.FromXlsx(_xlsx.AsMemory(), options: _prefetchOptions);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }


        [Benchmark]
        public long Xlsm_ExcelReader()
        {
            using MemoryStream ms = new(_xlsm, writable: false);
            using XlsxReader reader = Excel.FromXlsx(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsm_Sylvan()
        {
            using MemoryStream ms = new(_xlsm, writable: false);
            using Sylvan.Data.Excel.ExcelDataReader reader = Sylvan.Data.Excel.ExcelDataReader.Create(ms, ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions());
            return AccumulateSylvanExcel(reader);
        }

        [Benchmark]
        public long Xlsm_ExcelReader_Materialized()
        {
            using MemoryStream ms = new(_xlsm, writable: false);
            using XlsxReader reader = Excel.FromXlsx(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRowMaterialized(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsm_ExcelReader_Prefetch()
        {
            using MemoryStream ms = new(_xlsm, writable: false);
            using XlsxReader reader = Excel.FromXlsx(ms, options: _prefetchOptions);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsm_ExcelReader_Memory()
        {
            using XlsxReader reader = Excel.FromXlsx(_xlsm.AsMemory());
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsm_ExcelReader_Memory_Prefetch()
        {
            using XlsxReader reader = Excel.FromXlsx(_xlsm.AsMemory(), options: _prefetchOptions);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }


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
            using Sylvan.Data.Excel.ExcelDataReader reader = Sylvan.Data.Excel.ExcelDataReader.Create(ms, ExcelWorkbookType.ExcelBinary, new ExcelDataReaderOptions());
            return AccumulateSylvanExcel(reader);
        }

        [Benchmark]
        public long Xlsb_ExcelReader_Materialized()
        {
            using MemoryStream ms = new(_xlsb, writable: false);
            using XlsbReader reader = Excel.FromXlsb(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRowMaterialized(row); }
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
        public long Xlsb_ExcelReader_Memory()
        {
            using XlsbReader reader = Excel.FromXlsb(_xlsb.AsMemory());
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        [Benchmark]
        public long Xlsb_ExcelReader_Memory_Prefetch()
        {
            using XlsbReader reader = Excel.FromXlsb(_xlsb.AsMemory(), options: _prefetchOptions);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }


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
            using Sylvan.Data.Excel.ExcelDataReader reader = Sylvan.Data.Excel.ExcelDataReader.Create(ms, ExcelWorkbookType.Excel, new ExcelDataReaderOptions());
            return AccumulateSylvanExcel(reader);
        }

        [Benchmark]
        public long Xls_ExcelReader_Materialized()
        {
            using MemoryStream ms = new(_xls, writable: false);
            using XlsReader reader = Excel.FromXls(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRowMaterialized(row); }
            return acc;
        }

        [Benchmark]
        public long Xls_ExcelReader_Memory()
        {
            using XlsReader reader = Excel.FromXls(_xls.AsMemory());
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }


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

        [Benchmark]
        public long Csv_ExcelReader_Materialized()
        {
            using MemoryStream ms = new(_csv, writable: false);
            using CsvReader reader = Excel.FromCsv(ms);
            long acc = 0;
            foreach (Row row in reader)
            {
                foreach (RowCell rowCell in row.Cells)
                {
                    acc += rowCell.Value.GetString().Length;
                }
            }
            return acc;
        }

        [Benchmark]
        public long Csv_ExcelReader_Memory()
        {
            using CsvReader reader = Excel.FromCsv(_csv.AsMemory());
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
    }
}
