using System.Buffers.Text;
using System.Globalization;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using nietras.SeparatedValues;
using Sylvan.Data.Csv;

namespace ExcelReader.Benchmarks
{
    [MemoryDiagnoser]
    public class CsvReadBenchmark
    {
        [Params(50_000)]
        public int Rows { get; set; }

        private byte[] _csv = [];
        private byte[] _csvWide = [];

        [GlobalSetup]
        public void Setup()
        {
            _csv = CsvGenerator.Build(Rows);
            _csvWide = CsvGenerator.BuildWide(Rows, columns: 32);
            if (_csvWide.Length == 0)
            {
                throw new InvalidOperationException("Wide CSV fixture must not be empty.");
            }
        }

        [Benchmark(Baseline = true)]
        public long ExcelReader()
        {
            using var ms = new MemoryStream(_csv, writable: false);
            using var reader = Excel.FromCsv(ms);
            long acc = 0;
            foreach (var row in reader)
            {
                acc += AccumulateRow(row);
            }
            return acc;
        }

        [Benchmark]
        public long ExcelReaderWide()
        {
            using var ms = new MemoryStream(_csvWide, writable: false);
            using var reader = Excel.FromCsv(ms);
            long acc = 0;
            foreach (var row in reader)
            {
                for (int i = 0; i < row.ColumnCount; i++)
                {
                    acc += row[i].Value.Length;
                }
            }
            return acc;
        }

        [Benchmark]
        public async Task<long> ExcelReaderAsync()
        {
            await using var ms = new MemoryStream(_csv, writable: false);
            await using var reader = Excel.FromCsv(ms);
            await using var e = await reader.GetAsyncEnumeratorAsync();
            long acc = 0;
            while (await e.MoveNextAsync())
            {
                acc += AccumulateRow(e.Current);
            }
            return acc;
        }

        private static long AccumulateRow(Row row)
        {
            long acc = row[0].Value.Length;
            if (row[1].TryParse(CultureInfo.InvariantCulture, out int id))
            {
                acc += id;
            }
            if (row[2].TryGetDateTime(out DateTime date))
            {
                acc += date.Ticks;
            }
            if (row[3].TryParse(CultureInfo.InvariantCulture, out double value))
            {
                acc += (long)value;
            }
            return acc;
        }

        [Benchmark]
        public long Sep()
        {
            using var ms = new MemoryStream(_csv, writable: false);
            using var reader = nietras.SeparatedValues.Sep.Reader(o => o with
            {
                Sep = nietras.SeparatedValues.Sep.New(','),
                HasHeader = false,
                CultureInfo = CultureInfo.InvariantCulture,
            }).From(ms);
            long acc = 0;
            foreach (var row in reader)
            {
                acc += row[0].Span.Length;
                acc += row[1].Parse<int>();
                if (DateTime.TryParse(row[2].Span, CultureInfo.InvariantCulture, DateTimeStyles.None, out var d))
                {
                    acc += d.Ticks;
                }
                acc += (long)row[3].Parse<double>();
            }
            return acc;
        }

        [Benchmark]
        public long Sylvan()
        {
            using var ms = new MemoryStream(_csv, writable: false);
            using var tr = new StreamReader(ms);
            var options = new CsvDataReaderOptions { HasHeaders = false, Culture = CultureInfo.InvariantCulture };
            using var reader = CsvDataReader.Create(tr, options);
            long acc = 0;
            while (reader.Read())
            {
                acc += reader.GetFieldSpan(0).Length;
                acc += reader.GetInt32(1);
                acc += reader.GetDateTime(2).Ticks;
                acc += (long)reader.GetDouble(3);
            }
            return acc;
        }

        [Benchmark]
        public long CsvHelperLib()
        {
            using var ms = new MemoryStream(_csv, writable: false);
            using var tr = new StreamReader(ms);
            var config = new CsvHelper.Configuration.CsvConfiguration(CultureInfo.InvariantCulture) { HasHeaderRecord = false };
            using var csv = new CsvHelper.CsvReader(tr, config);
            long acc = 0;
            while (csv.Read())
            {
                acc += csv.GetField(0)!.Length;
                acc += csv.GetField<int>(1);
                acc += csv.GetField<DateTime>(2).Ticks;
                acc += (long)csv.GetField<double>(3);
            }
            return acc;
        }
    }
}
