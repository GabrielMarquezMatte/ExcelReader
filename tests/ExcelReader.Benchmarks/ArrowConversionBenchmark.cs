using Apache.Arrow;
using BenchmarkDotNet.Attributes;
using ExcelReader.Arrow;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;

namespace ExcelReader.Benchmarks
{
    // CsvAllString is the shape every ToArrowRecordBatch() call without an explicit schema actually
    // takes (CSV has no type tags, so inference can only return StringColumn). XlsbTyped uses the
    // reader's shared-string cache, which favors the string path least.
    [MemoryDiagnoser]
    public class ArrowConversionBenchmark
    {
        [Params(100_000)]
        public int Rows { get; set; }

        private byte[] _csvAllString = [];
        private byte[] _csvTyped = [];
        private byte[] _xlsbTyped = [];

        private static readonly ExcelColumnSchema[] AllStringSchema = BuildAllStringSchema(8);

        private static readonly ExcelColumnSchema[] TypedSchema =
        [
            new() { Index = 0, Name = "Name", Type = ExcelColumnType.StringColumn },
            new() { Index = 1, Name = "Id", Type = ExcelColumnType.Int64Column },
            new() { Index = 2, Name = "Date", Type = ExcelColumnType.TimestampColumn },
            new() { Index = 3, Name = "Value", Type = ExcelColumnType.Float64Column },
        ];

        private static ExcelColumnSchema[] BuildAllStringSchema(int columns)
        {
            var schema = new ExcelColumnSchema[columns];
            for (int i = 0; i < columns; i++)
            {
                schema[i] = new ExcelColumnSchema { Index = i, Name = $"c{i}", Type = ExcelColumnType.StringColumn };
            }
            return schema;
        }

        [GlobalSetup]
        public async Task SetupAsync()
        {
            _csvAllString = CsvGenerator.BuildWide(Rows, AllStringSchema.Length);
            _csvTyped = CsvGenerator.BuildTyped(Rows);
            _xlsbTyped = await WorkbookGenerator.BuildTypedXlsbAsync(Rows);
        }

        [Benchmark(Baseline = true)]
        public long CsvAllString()
        {
            using var reader = Excel.FromCsv(_csvAllString);
            using RecordBatch batch = reader.ToArrowRecordBatch(AllStringSchema, headerRow: 0);
            return batch.Length;
        }

        [Benchmark]
        public long CsvTyped()
        {
            using var reader = Excel.FromCsv(_csvTyped);
            using RecordBatch batch = reader.ToArrowRecordBatch(TypedSchema);
            return batch.Length;
        }

        [Benchmark]
        public long XlsbTyped()
        {
            using var ms = new MemoryStream(_xlsbTyped, writable: false);
            using var reader = Excel.FromXlsb(ms);
            using RecordBatch batch = reader.ToArrowRecordBatch(TypedSchema);
            return batch.Length;
        }

        [Benchmark]
        public long CsvTypedWithInference()
        {
            using var reader = Excel.FromCsv(_csvTyped);
            using RecordBatch batch = reader.ToArrowRecordBatch();
            return batch.Length;
        }
    }
}
