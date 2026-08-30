using Apache.Arrow;
using BenchmarkDotNet.Attributes;
using ExcelReader.Arrow;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;

namespace ExcelReader.Benchmarks
{
    // Measures ExcelReader.Arrow's whole-sheet conversion across the three shapes that stress it
    // differently:
    //
    //   CsvAllString  - every column a string. Not a corner case: CSV cells carry no type tag, so
    //                   SchemaInference can only ever return StringColumn for them (see its "no text
    //                   sniffing" remark), making this the shape every ToArrowRecordBatch() call
    //                   without an explicit schema actually takes.
    //   CsvTyped      - string/int64/date/float64, so per-cell conversion competes with the string path.
    //   XlsbTyped     - the same typed shape from a binary workbook, where repeated strings come back
    //                   from the reader's shared-string cache already interned. This is the shape that
    //                   flatters the string path LEAST, and it is measured for exactly that reason.
    //
    // InferSchema is kept out of the measured region (explicit schemas below) so these numbers move
    // only when the conversion itself does.
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

        // BuildWide is headerless, so headerRow: 0 - otherwise the first data row is eaten as a header.
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

        // The default path an application takes when it does not hand in a schema: inference samples
        // the sheet first, then the conversion runs. Separated from CsvTyped so a regression in either
        // half is attributable.
        [Benchmark]
        public long CsvTypedWithInference()
        {
            using var reader = Excel.FromCsv(_csvTyped);
            using RecordBatch batch = reader.ToArrowRecordBatch();
            return batch.Length;
        }
    }
}
