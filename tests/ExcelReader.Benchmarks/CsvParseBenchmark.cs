using System.Globalization;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using nietras.SeparatedValues;
using Sylvan.Data;
using Sylvan.Data.Csv;

namespace ExcelReader.Benchmarks
{
    // The CSV parser parses DateTime from the cell text natively (no converter), matching how Sep,
    // Sylvan, and CsvHelper handle the Date column — a like-for-like comparison.
    public sealed class CsvRecord
    {
        public string? Name { get; set; }
        public int Id { get; set; }
        public DateTime Date { get; set; }
        public double Value { get; set; }
    }

    // Maps a header + `Rows` data rows into strongly-typed CsvRecord objects, comparing
    // ExcelReader's ExcelParser<T> (sync + async) against Sep, Sylvan.Data.Csv, and CsvHelper.
    [MemoryDiagnoser]
    public class CsvParseBenchmark
    {
        [Params(50_000)]
        public int Rows { get; set; }

        private byte[] _csv = [];

        [GlobalSetup]
        public void Setup()
        {
            _csv = CsvGenerator.BuildTyped(Rows);
        }

        private static long Accumulate(CsvRecord rec)
        {
            return rec.Id + (long)rec.Value + (rec.Name?.Length ?? 0) + rec.Date.Ticks;
        }

        [Benchmark(Baseline = true)]
        public long ExcelParserSync()
        {
            using var ms = new MemoryStream(_csv, writable: false);
            using var reader = Excel.FromCsv(ms);
            long acc = 0;
            foreach (CsvRecord rec in new ExcelParser<CsvRecord>().Parse(reader))
            {
                acc += Accumulate(rec);
            }
            return acc;
        }

        [Benchmark]
        public async Task<long> ExcelParserAsync()
        {
            await using var ms = new MemoryStream(_csv, writable: false);
            await using var reader = Excel.FromCsv(ms);
            long acc = 0;
            await foreach (CsvRecord rec in new ExcelParser<CsvRecord>().ParseAsync(reader))
            {
                acc += Accumulate(rec);
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
                CultureInfo = CultureInfo.InvariantCulture,
            }).From(ms);
            long acc = 0;
            foreach (var row in reader)
            {
                var rec = new CsvRecord
                {
                    Name = row["Name"].ToString(),
                    Id = row["Id"].Parse<int>(),
                    Date = DateTime.Parse(row["Date"].Span, CultureInfo.InvariantCulture, DateTimeStyles.None),
                    Value = row["Value"].Parse<double>(),
                };
                acc += Accumulate(rec);
            }
            return acc;
        }

        [Benchmark]
        public long Sylvan()
        {
            using var ms = new MemoryStream(_csv, writable: false);
            using var tr = new StreamReader(ms);
            var options = new CsvDataReaderOptions { Culture = CultureInfo.InvariantCulture };
            using var reader = CsvDataReader.Create(tr, options);
            long acc = 0;
            foreach (CsvRecord rec in reader.GetRecords<CsvRecord>())
            {
                acc += Accumulate(rec);
            }
            return acc;
        }

        [Benchmark]
        public long CsvHelperLib()
        {
            using var ms = new MemoryStream(_csv, writable: false);
            using var tr = new StreamReader(ms);
            using var csv = new CsvHelper.CsvReader(tr, CultureInfo.InvariantCulture);
            long acc = 0;
            foreach (CsvRecord rec in csv.GetRecords<CsvRecord>())
            {
                acc += Accumulate(rec);
            }
            return acc;
        }
    }
}
