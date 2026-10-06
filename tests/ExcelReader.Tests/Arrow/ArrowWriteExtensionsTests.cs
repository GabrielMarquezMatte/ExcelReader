using Apache.Arrow;
using Apache.Arrow.Types;
using ExcelReader.Arrow;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Schema;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Core.Writer.Csv;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Arrow
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

            var name = Assert.IsType<StringArray>(roundTripped.Column(0));
            Assert.Equal("alice", name.GetString(0));
            Assert.True(name.IsNull(1));

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

            using XlsxWorkbook reader = Excel.FromXlsx(ms.ToArray());
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

            using XlsxWorkbook reader = Excel.FromXlsx(ms.ToArray());
            using XlsxWorkbook.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.True(e.Current[0].TryParse(null, out long value));
            Assert.Equal(3L, value);
            Assert.False(e.MoveNext());
        }

        private static RecordBatch BuildWideBatch()
        {
            var decimalType = new Decimal128Type(10, 2);
            var secondTime = new Time32Type(TimeUnit.Second);

            var i8 = new Int8Array.Builder();
            i8.Append(-8);
            i8.Append(7);
            var i16 = new Int16Array.Builder();
            i16.Append(-16);
            i16.AppendNull();
            var i32 = new Int32Array.Builder();
            i32.Append(-32);
            i32.AppendNull();
            var u8 = new UInt8Array.Builder();
            u8.Append(8);
            u8.AppendNull();
            var u16 = new UInt16Array.Builder();
            u16.Append(16);
            u16.AppendNull();
            var u32 = new UInt32Array.Builder();
            u32.Append(32);
            u32.AppendNull();
            var u64 = new UInt64Array.Builder();
            u64.Append(64);
            u64.AppendNull();
            var half = new HalfFloatArray.Builder();
            half.Append((Half)0.5f);
            half.AppendNull();
            var single = new FloatArray.Builder();
            single.Append(0.1f);
            single.AppendNull();
            var dec = new Decimal128Array.Builder(decimalType);
            dec.Append(12.34m);
            dec.AppendNull();
            var view = new StringViewArray.Builder();
            view.Append("alice");
            view.AppendNull();
            var date = new Date64Array.Builder();
            date.Append(new DateOnly(2024, 1, 15));
            date.AppendNull();
            var time = new Time32Array.Builder(secondTime);
            time.Append(new TimeOnly(13, 30, 0));
            time.AppendNull();

            Schema schema = new Schema.Builder()
                .Field(new Field("i8", Int8Type.Default, nullable: true))
                .Field(new Field("i16", Int16Type.Default, nullable: true))
                .Field(new Field("i32", Int32Type.Default, nullable: true))
                .Field(new Field("u8", UInt8Type.Default, nullable: true))
                .Field(new Field("u16", UInt16Type.Default, nullable: true))
                .Field(new Field("u32", UInt32Type.Default, nullable: true))
                .Field(new Field("u64", UInt64Type.Default, nullable: true))
                .Field(new Field("half", HalfFloatType.Default, nullable: true))
                .Field(new Field("single", FloatType.Default, nullable: true))
                .Field(new Field("dec", decimalType, nullable: true))
                .Field(new Field("view", StringViewType.Default, nullable: true))
                .Field(new Field("date", Date64Type.Default, nullable: true))
                .Field(new Field("time", secondTime, nullable: true))
                .Field(new Field("nothing", NullType.Default, nullable: true))
                .Build();

            return new RecordBatch(
                schema,
                [
                    i8.Build(), i16.Build(), i32.Build(), u8.Build(), u16.Build(), u32.Build(), u64.Build(),
                    half.Build(), single.Build(), dec.Build(), view.Build(), date.Build(), time.Build(), new NullArray(2),
                ],
                2);
        }

        private static void AssertWideRoundTrips(IExcelRowReader reader)
        {
            ExcelColumnType[] types =
            [
                ExcelColumnType.Int64Column, ExcelColumnType.Int64Column, ExcelColumnType.Int64Column,
                ExcelColumnType.Int64Column, ExcelColumnType.Int64Column, ExcelColumnType.Int64Column,
                ExcelColumnType.Int64Column, ExcelColumnType.Float64Column, ExcelColumnType.Float64Column,
                ExcelColumnType.Float64Column, ExcelColumnType.StringColumn, ExcelColumnType.DateColumn,
                ExcelColumnType.TimeColumn, ExcelColumnType.StringColumn,
            ];
            var schema = new ExcelColumnSchema[types.Length];
            for (int i = 0; i < types.Length; i++)
            {
                schema[i] = new() { Index = i, Name = "c" + i, Type = types[i], IsNullable = true };
            }
            RecordBatch batch = reader.ToArrowRecordBatch(schema);

            Assert.Equal(2, batch.Length);
            long[] integers = [-8, -16, -32, 8, 16, 32, 64];
            for (int i = 0; i < integers.Length; i++)
            {
                Assert.Equal(integers[i], Assert.IsType<Int64Array>(batch.Column(i)).GetValue(0));
            }
            Assert.Equal(0.5, Assert.IsType<DoubleArray>(batch.Column(7)).GetValue(0));
            Assert.Equal(0.1, Assert.IsType<DoubleArray>(batch.Column(8)).GetValue(0));
            Assert.Equal(12.34, Assert.IsType<DoubleArray>(batch.Column(9)).GetValue(0));
            Assert.Equal("alice", Assert.IsType<StringArray>(batch.Column(10)).GetString(0));
            Assert.Equal(new DateOnly(2024, 1, 15), Assert.IsType<Date32Array>(batch.Column(11)).GetDateOnly(0));
            Assert.Equal(new TimeOnly(13, 30, 0), Assert.IsType<Time64Array>(batch.Column(12)).GetTime(0));
            Assert.Equal(7L, ((Int64Array)batch.Column(0)).GetValue(1));
            for (int i = 1; i < types.Length; i++)
            {
                Assert.True(batch.Column(i).IsNull(1), "column " + i);
            }
            Assert.True(batch.Column(13).IsNull(0));
        }

        [Fact]
        public void WriteRecordBatch_Xlsx_RoundTrips_Widened_Types()
        {
            using var ms = new MemoryStream();
            using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
            {
                workbook.WriteRecordBatch(BuildWideBatch());
            }

            using XlsxWorkbook reader = Excel.FromXlsx(ms.ToArray());
            AssertWideRoundTrips(reader);
        }

        [Fact]
        public void WriteRecordBatch_Xlsb_RoundTrips_Widened_Types()
        {
            using var ms = new MemoryStream();
            using (XlsbWorkbookWriter workbook = XlsbWorkbookWriter.Create(ms, leaveOpen: true))
            {
                workbook.WriteRecordBatch(BuildWideBatch());
            }

            using var reader = Excel.FromXlsb(ms.ToArray());
            AssertWideRoundTrips(reader);
        }

        [Fact]
        public void WriteRecordBatch_UnsupportedArrowType_Throws()
        {
            var col = new BinaryArray.Builder();
            col.Append(new byte[] { 1 }.AsSpan());
            Schema schema = new Schema.Builder()
                .Field(new Field("f", BinaryType.Default, nullable: false))
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

            using XlsxWorkbook reader = Excel.FromXlsx(ms.ToArray());
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
