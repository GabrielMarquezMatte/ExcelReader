using System.Globalization;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Writer.Csv;
using nietras.SeparatedValues;

namespace ExcelReader.Benchmarks
{
    [MemoryDiagnoser]
    public class CsvWriteBenchmark
    {
        [Params(50_000)]
        public int Rows { get; set; }

        private List<Record> _records = [];

        [GlobalSetup]
        public void Setup()
        {
            _records = WorkbookGenerator.Records(Rows);
        }

        [Benchmark(Baseline = true)]
        public long ExcelReaderWriter()
        {
            using var ms = new MemoryStream(4 * 1024 * 1024);
            using (CsvWriter writer = CsvWriter.Create(ms, leaveOpen: true))
            {
                using (CsvRowWriter header = writer.StartRow())
                {
                    header.Write("Name");
                    header.Write("Id");
                    header.Write("Date");
                    header.Write("Value");
                }
                for (int i = 0; i < _records.Count; i++)
                {
                    Record rec = _records[i];
                    using CsvRowWriter row = writer.StartRow();
                    row.Write(rec.Name);
                    row.Write(rec.Id);
                    row.Write(rec.Date);
                    row.Write(rec.Value);
                }
            }
            return ms.Length;
        }

        [Benchmark]
        public long Sep()
        {
            using var ms = new MemoryStream(4 * 1024 * 1024);
            using (var writer = nietras.SeparatedValues.Sep.Writer(o => o with
            {
                Sep = nietras.SeparatedValues.Sep.New(','),
                CultureInfo = CultureInfo.InvariantCulture,
            }).To(ms, leaveOpen: true))
            {
                for (int i = 0; i < _records.Count; i++)
                {
                    Record rec = _records[i];
                    using var row = writer.NewRow();
                    row["Name"].Set(rec.Name ?? "");
                    row["Id"].Format(rec.Id);
                    row["Date"].Format(rec.Date);
                    row["Value"].Format(rec.Value);
                }
            }
            return ms.Length;
        }

        [Benchmark]
        public long SylvanWriter()
        {
            using var ms = new MemoryStream(4 * 1024 * 1024);
            using (var tw = new StreamWriter(ms, leaveOpen: true))
            using (var writer = Sylvan.Data.Csv.CsvDataWriter.Create(tw))
            {
                using var reader = Sylvan.Data.ObjectDataReader.Create(_records);
                writer.Write(reader);
            }
            return ms.Length;
        }
    }
}
