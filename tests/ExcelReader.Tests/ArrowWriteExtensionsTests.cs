using Apache.Arrow;
using Apache.Arrow.Types;
using ExcelReader.Arrow;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    public sealed class ArrowWriteExtensionsTests
    {
        private static readonly TimestampType MicrosecondTimestamp = new(TimeUnit.Microsecond, timezone: (string?)null);
        private static readonly Time64Type MicrosecondTime = new(TimeUnit.Microsecond);

        private static RecordBatch BuildBatch()
        {
            var name = new StringArray.Builder();
            name.Append("alice");
            name.AppendNull();

            // qty stays non-null in both rows: a row with every other column blank still needs one
            // real cell, or a sparse-format writer (XLSB) has nothing to record for that row at all
            // and it disappears entirely on read-back instead of coming back as an empty row.
            var qty = new Int64Array.Builder();
            qty.Append(3);
            qty.Append(99);

            var price = new DoubleArray.Builder();
            price.Append(1.5);
            price.AppendNull();

            var flag = new BooleanArray.Builder();
            flag.Append(true);
            flag.AppendNull();

            var when = new Date32Array.Builder();
            when.Append(new DateOnly(2024, 1, 15));
            when.AppendNull();

            var atTime = new Time64Array.Builder(MicrosecondTime);
            atTime.Append(new TimeOnly(13, 30, 0));
            atTime.AppendNull();

            var stamp = new TimestampArray.Builder(MicrosecondTimestamp);
            stamp.Append(new DateTimeOffset(new DateTime(2024, 1, 15, 13, 30, 0), TimeSpan.Zero));
            stamp.AppendNull();

            Schema schema = new Schema.Builder()
                .Field(new Field("name", StringType.Default, nullable: true))
                .Field(new Field("qty", Int64Type.Default, nullable: true))
                .Field(new Field("price", DoubleType.Default, nullable: true))
                .Field(new Field("flag", BooleanType.Default, nullable: true))
                .Field(new Field("when", Date32Type.Default, nullable: true))
                .Field(new Field("at_time", MicrosecondTime, nullable: true))
                .Field(new Field("stamp", MicrosecondTimestamp, nullable: true))
                .Build();

            return new RecordBatch(schema, [name.Build(), qty.Build(), price.Build(), flag.Build(), when.Build(), atTime.Build(), stamp.Build()], 2);
        }

        private static ExcelColumnSchema[] ReadBackSchema()
        {
            return
            [
                new() { Index = 0, Name = "name", Type = ExcelColumnType.StringColumn, IsNullable = true },
                new() { Index = 1, Name = "qty", Type = ExcelColumnType.Int64Column, IsNullable = true },
                new() { Index = 2, Name = "price", Type = ExcelColumnType.Float64Column, IsNullable = true },
                new() { Index = 3, Name = "flag", Type = ExcelColumnType.BoolColumn, IsNullable = true },
                new() { Index = 4, Name = "when", Type = ExcelColumnType.DateColumn, IsNullable = true },
                new() { Index = 5, Name = "at_time", Type = ExcelColumnType.TimeColumn, IsNullable = true },
                new() { Index = 6, Name = "stamp", Type = ExcelColumnType.TimestampColumn, IsNullable = true },
            ];
        }

        private static void AssertRoundTrips(RecordBatch roundTripped)
        {
            Assert.Equal(2, roundTripped.Length);

            // ToArrowRecordBatch's StringColumnAppender never appends null (see ColumnAppender.cs) — a
            // blank string cell always round-trips as "", regardless of the column's IsNullable flag.
            var name = Assert.IsType<StringArray>(roundTripped.Column(0));
            Assert.Equal("alice", name.GetString(0));
            Assert.Equal("", name.GetString(1));

            var qty = Assert.IsType<Int64Array>(roundTripped.Column(1));
            Assert.Equal(3L, qty.GetValue(0));
            Assert.Equal(99L, qty.GetValue(1));

            var price = Assert.IsType<DoubleArray>(roundTripped.Column(2));
            Assert.Equal(1.5, price.GetValue(0));
            Assert.True(price.IsNull(1));

            var flag = Assert.IsType<BooleanArray>(roundTripped.Column(3));
            Assert.True(flag.GetValue(0));
            Assert.True(flag.IsNull(1));

            var when = Assert.IsType<Date32Array>(roundTripped.Column(4));
            Assert.Equal(new DateOnly(2024, 1, 15), when.GetDateOnly(0));
            Assert.True(when.IsNull(1));

            var atTime = Assert.IsType<Time64Array>(roundTripped.Column(5));
            Assert.Equal(new TimeOnly(13, 30, 0), atTime.GetTime(0));
            Assert.True(atTime.IsNull(1));

            var stamp = Assert.IsType<TimestampArray>(roundTripped.Column(6));
            Assert.Equal(new DateTime(2024, 1, 15, 13, 30, 0), stamp.GetTimestamp(0)!.Value.UtcDateTime);
            Assert.True(stamp.IsNull(1));
        }

        [Fact]
        public void WriteRecordBatch_Xlsx_RoundTrips_Every_Supported_Type()
        {
            using var ms = new MemoryStream();
            using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
            {
                workbook.WriteRecordBatch(BuildBatch());
            }
            ms.Position = 0;

            using XlsxReader reader = Excel.FromXlsx(ms.ToArray());
            RecordBatch roundTripped = reader.ToArrowRecordBatch(ReadBackSchema());

            AssertRoundTrips(roundTripped);
        }

        [Fact]
        public void WriteRecordBatch_Xlsb_RoundTrips_Every_Supported_Type()
        {
            using var ms = new MemoryStream();
            using (XlsbWorkbookWriter workbook = XlsbWorkbookWriter.Create(ms, leaveOpen: true))
            {
                workbook.WriteRecordBatch(BuildBatch());
            }
            ms.Position = 0;

            using var reader = Excel.FromXlsb(ms.ToArray());
            RecordBatch roundTripped = reader.ToArrowRecordBatch(ReadBackSchema());

            AssertRoundTrips(roundTripped);
        }

        [Fact]
        public void WriteRecordBatch_Xls_RoundTrips_String_And_Numeric_Columns()
        {
            var name = new StringArray.Builder();
            name.Append("alice");
            var qty = new Int64Array.Builder();
            qty.Append(3);
            Schema schema = new Schema.Builder()
                .Field(new Field("name", StringType.Default, nullable: false))
                .Field(new Field("qty", Int64Type.Default, nullable: false))
                .Build();
            var batch = new RecordBatch(schema, [name.Build(), qty.Build()], 1);

            using var ms = new MemoryStream();
            using (XlsWorkbookWriter workbook = XlsWorkbookWriter.Create(ms, leaveOpen: true))
            {
                workbook.WriteRecordBatch(batch);
            }
            ms.Position = 0;

            using var reader = Excel.FromXls(ms.ToArray());
            RecordBatch roundTripped = reader.ToArrowRecordBatch();

            var nameCol = Assert.IsType<StringArray>(roundTripped.Column(0));
            Assert.Equal("alice", nameCol.GetString(0));
            var qtyCol = Assert.IsType<Int64Array>(roundTripped.Column(1));
            Assert.Equal(3L, qtyCol.GetValue(0));
        }

        [Fact]
        public void WriteRecordBatch_Csv_WritesHeaderAndRows()
        {
            var name = new StringArray.Builder();
            name.Append("alice");
            name.Append("bob");
            Schema schema = new Schema.Builder()
                .Field(new Field("name", StringType.Default, nullable: false))
                .Build();
            var batch = new RecordBatch(schema, [name.Build()], 2);

            using var ms = new MemoryStream();
            using (CsvWorkbookWriter workbook = CsvWorkbookWriter.Create(ms, leaveOpen: true))
            {
                workbook.WriteRecordBatch(batch);
            }
            ms.Position = 0;

            string csv = System.Text.Encoding.UTF8.GetString(ms.ToArray());
            Assert.Equal("name\r\nalice\r\nbob\r\n", csv);
        }

        [Fact]
        public void WriteRecordBatch_WithHeaderFalse_OmitsTheHeaderRow()
        {
            var qty = new Int64Array.Builder();
            qty.Append(3);
            Schema schema = new Schema.Builder()
                .Field(new Field("qty", Int64Type.Default, nullable: false))
                .Build();
            var batch = new RecordBatch(schema, [qty.Build()], 1);

            using var ms = new MemoryStream();
            using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
            {
                workbook.WriteRecordBatch(batch, writeHeader: false);
            }
            ms.Position = 0;

            using XlsxReader reader = Excel.FromXlsx(ms.ToArray());
            using XlsxReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.True(e.Current[0].TryParse(null, out long value));
            Assert.Equal(3L, value);
            Assert.False(e.MoveNext());
        }

        [Fact]
        public void WriteRecordBatch_UnsupportedArrowType_Throws()
        {
            var col = new FloatArray.Builder();
            col.Append(1.5f);
            Schema schema = new Schema.Builder()
                .Field(new Field("f", FloatType.Default, nullable: false))
                .Build();
            var batch = new RecordBatch(schema, [col.Build()], 1);

            using var ms = new MemoryStream();
            using XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(ms, leaveOpen: true);

            Assert.Throws<NotSupportedException>(() => workbook.WriteRecordBatch(batch));
        }

        [Fact]
        public async Task WriteRecordBatchAsync_Xlsx_RoundTrips()
        {
            using var ms = new MemoryStream();
            await using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
            {
                await workbook.WriteRecordBatchAsync(BuildBatch(), ct: TestContext.Current.CancellationToken);
            }
            ms.Position = 0;

            using XlsxReader reader = Excel.FromXlsx(ms.ToArray());
            RecordBatch roundTripped = reader.ToArrowRecordBatch(ReadBackSchema());

            AssertRoundTrips(roundTripped);
        }

        [Fact]
        public void WriteRecordBatch_NullWorkbook_Throws()
        {
            XlsxWorkbookWriter workbook = null!;
            Assert.Throws<ArgumentNullException>(() => workbook.WriteRecordBatch(BuildBatch()));
        }

        [Fact]
        public void WriteRecordBatch_NullBatch_Throws()
        {
            using var ms = new MemoryStream();
            using XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(ms, leaveOpen: true);
            Assert.Throws<ArgumentNullException>(() => workbook.WriteRecordBatch(null!));
        }
    }
}
