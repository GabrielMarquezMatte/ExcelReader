using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using nietras.SeparatedValues;

namespace ExcelReader.Benchmarks
{
    /// <summary>Which corpus shape the parallel comparison runs against.</summary>
    public enum ParallelCorpus
    {
        /// <summary>Three integer columns: almost nothing to convert, so record scanning and the ordered merge dominate.</summary>
        NarrowInt,

        /// <summary>Two strings, a date, two decimals and an integer: per-row conversion dominates, which is the shape parallelism helps most.</summary>
        ConversionHeavy,
    }

    [MemoryDiagnoser]
    public class CsvParallelVsSepBenchmark
    {
        private const int NarrowRows = 8_000_000;
        private const int WideRows = 3_000_000;

        [Params(ParallelCorpus.NarrowInt, ParallelCorpus.ConversionHeavy)]
        public ParallelCorpus Corpus { get; set; }

        private string _narrow = "";
        private string _wide = "";

        private string Path => Corpus == ParallelCorpus.NarrowInt ? _narrow : _wide;

        [GlobalSetup]
        public void Setup()
        {
            _narrow = CsvGenerator.WriteNarrowIntFile(NarrowRows);
            _wide = CsvGenerator.WriteConversionHeavyFile(WideRows);
        }

        [GlobalCleanup]
        public void Cleanup()
        {
            File.Delete(_narrow);
            File.Delete(_wide);
        }

        private static NarrowRow ParseNarrow(SepReader.Row row)
        {
            return new NarrowRow
            {
                A = row["A"].Parse<int>(),
                B = row["B"].Parse<int>(),
                C = row["C"].Parse<int>(),
            };
        }

        private static WideRow ParseWide(SepReader.Row row)
        {
            return new WideRow
            {
                Region = row["Region"].ToString(),
                Country = row["Country"].ToString(),
                OrderDate = row["OrderDate"].Parse<DateTime>(),
                UnitPrice = row["UnitPrice"].Parse<decimal>(),
                TotalRevenue = row["TotalRevenue"].Parse<decimal>(),
                Units = row["Units"].Parse<int>(),
            };
        }

        [Benchmark(Baseline = true)]
        public Task<long> ExcelReader_Sequential()
        {
            return ReadExcelReaderAsync(degreeOfParallelism: 1);
        }

        [Benchmark]
        public Task<long> ExcelReader_Parallel()
        {
            return ReadExcelReaderAsync(degreeOfParallelism: 0);
        }

        [Benchmark]
        public long Sep_Sequential()
        {
            long acc = 0;
            using SepReader reader = Sep.Reader().FromFile(Path);
            if (Corpus == ParallelCorpus.NarrowInt)
            {
                foreach (SepReader.Row row in reader)
                {
                    acc += ParseNarrow(row).A;
                }
                return acc;
            }
            foreach (SepReader.Row row in reader)
            {
                acc += ParseWide(row).Units;
            }
            return acc;
        }

        [Benchmark]
        public long Sep_Parallel()
        {
            long acc = 0;
            using SepReader reader = Sep.Reader().FromFile(Path);
            if (Corpus == ParallelCorpus.NarrowInt)
            {
                foreach (NarrowRow row in reader.ParallelEnumerate(ParseNarrow))
                {
                    acc += row.A;
                }
                return acc;
            }
            foreach (WideRow row in reader.ParallelEnumerate(ParseWide))
            {
                acc += row.Units;
            }
            return acc;
        }

        private async Task<long> ReadExcelReaderAsync(int degreeOfParallelism)
        {
            if (degreeOfParallelism == 1)
            {
                await using var reader = Excel.FromCsvFile(Path);
                long seqAcc = 0;
                if (Corpus == ParallelCorpus.NarrowInt)
                {
                    ExcelParser<NarrowRow> parser = new();
                    foreach (NarrowRow row in parser.Parse(reader))
                    {
                        seqAcc += row.A;
                    }
                    return seqAcc;
                }
                ExcelParser<WideRow> wideParser = new();
                foreach (WideRow row in wideParser.Parse(reader))
                {
                    seqAcc += row.Units;
                }
                return seqAcc;
            }
            long acc = 0;
            if (Corpus == ParallelCorpus.NarrowInt)
            {
                await foreach (NarrowRow row in Excel.ParseCsvParallelAsync<NarrowRow>(Path, degreeOfParallelism))
                {
                    acc += row.A;
                }
                return acc;
            }
            await foreach (WideRow row in Excel.ParseCsvParallelAsync<WideRow>(Path, degreeOfParallelism))
            {
                acc += row.Units;
            }
            return acc;
        }
    }
}
