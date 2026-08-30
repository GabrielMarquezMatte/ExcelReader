using System.Runtime.InteropServices;
using System.Text;
using Apache.Arrow;
using ExcelReader.Arrow;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;
using ExcelReader.Native;

namespace ExcelReader.Tests
{
    public sealed class ArrowConversionTests
    {
        [Fact]
        public void ToArrowRecordBatch_Should_Convert_String_And_Int64_Columns()
        {
            using CsvReader reader = Excel.FromCsv(Encoding.UTF8.GetBytes("name,qty\nwidget,3\ngadget,7\n"));
            ExcelColumnSchema[] schema =
            [
                new() { Index = 0, Name = "name", Type = ExcelColumnType.StringColumn },
                new() { Index = 1, Name = "qty", Type = ExcelColumnType.Int64Column },
            ];

            RecordBatch batch = reader.ToArrowRecordBatch(schema);

            Assert.Equal(2, batch.Length);
            Assert.Equal(2, batch.ColumnCount);
            Assert.Equal("name", batch.Schema.GetFieldByIndex(0).Name);
            var name = Assert.IsType<StringArray>(batch.Column(0));
            Assert.Equal("widget", name.GetString(0));
            Assert.Equal("gadget", name.GetString(1));
            var qty = Assert.IsType<Int64Array>(batch.Column(1));
            Assert.Equal(3L, qty.GetValue(0));
            Assert.Equal(7L, qty.GetValue(1));
        }

        [Fact]
        public void ToArrowRecordBatch_Should_Append_Null_For_Blank_Cells_On_Nullable_Columns()
        {
            using CsvReader reader = Excel.FromCsv(Encoding.UTF8.GetBytes("qty\n5\n\n"));
            ExcelColumnSchema[] schema = [new() { Index = 0, Name = "qty", Type = ExcelColumnType.Int64Column, IsNullable = true }];

            RecordBatch batch = reader.ToArrowRecordBatch(schema);

            var qty = Assert.IsType<Int64Array>(batch.Column(0));
            Assert.False(qty.IsNull(0));
            Assert.True(qty.IsNull(1));
            Assert.Equal(1, qty.NullCount);
        }

        [Fact]
        public void ToArrowRecordBatch_Should_Throw_When_A_NonNullable_Value_Fails_To_Convert()
        {
            using CsvReader reader = Excel.FromCsv(Encoding.UTF8.GetBytes("qty\nnotanumber\n"));
            ExcelColumnSchema[] schema = [new() { Index = 0, Name = "qty", Type = ExcelColumnType.Int64Column }];

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => reader.ToArrowRecordBatch(schema));
            Assert.Contains("qty", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void ToArrowRecordBatch_Should_Convert_Float64_Column()
        {
            using CsvReader reader = Excel.FromCsv(Encoding.UTF8.GetBytes("price\n1.5\n2.25\n"));
            ExcelColumnSchema[] schema = [new() { Index = 0, Name = "price", Type = ExcelColumnType.Float64Column }];

            RecordBatch batch = reader.ToArrowRecordBatch(schema);

            var price = Assert.IsType<DoubleArray>(batch.Column(0));
            Assert.Equal(1.5, price.GetValue(0));
            Assert.Equal(2.25, price.GetValue(1));
        }

        [Fact]
        public void ToArrowRecordBatch_Should_Convert_Bool_Column()
        {
            using CsvReader reader = Excel.FromCsv(Encoding.UTF8.GetBytes("flag\ntrue\nfalse\n"));
            ExcelColumnSchema[] schema = [new() { Index = 0, Name = "flag", Type = ExcelColumnType.BoolColumn }];

            RecordBatch batch = reader.ToArrowRecordBatch(schema);

            var flag = Assert.IsType<BooleanArray>(batch.Column(0));
            Assert.True(flag.GetValue(0));
            Assert.False(flag.GetValue(1));
        }

        [Fact]
        public void ToArrowRecordBatch_Should_Convert_Date_Column()
        {
            using var reader = Excel.FromCsv(Encoding.UTF8.GetBytes("day\n2024-01-15\n2024-03-01\n"));
            ExcelColumnSchema[] schema = [new() { Index = 0, Name = "day", Type = ExcelColumnType.DateColumn }];

            RecordBatch batch = reader.ToArrowRecordBatch(schema);

            var day = Assert.IsType<Date32Array>(batch.Column(0));
            int expectedFirst = (new DateTime(2024, 1, 15) - DateTime.UnixEpoch).Days;
            int expectedSecond = (new DateTime(2024, 3, 1) - DateTime.UnixEpoch).Days;
            Assert.Equal(expectedFirst, day.GetValue(0));
            Assert.Equal(expectedSecond, day.GetValue(1));
        }

        [Fact]
        public void ToArrowRecordBatch_Should_Convert_Time_Column()
        {
            using var reader = Excel.FromCsv(Encoding.UTF8.GetBytes("clock\n13:30:00\n\n"));
            ExcelColumnSchema[] schema = [new() { Index = 0, Name = "clock", Type = ExcelColumnType.TimeColumn, IsNullable = true }];

            RecordBatch batch = reader.ToArrowRecordBatch(schema);

            var clock = Assert.IsType<Time64Array>(batch.Column(0));
            Assert.Equal(new TimeSpan(13, 30, 0).Ticks / 10, clock.GetValue(0));
            Assert.True(clock.IsNull(1));
        }

        [Fact]
        public void ToArrowRecordBatch_Should_Convert_Timestamp_Column()
        {
            using var reader = Excel.FromCsv(Encoding.UTF8.GetBytes("stamp\n2024-01-15 13:30:00\n"));
            ExcelColumnSchema[] schema = [new() { Index = 0, Name = "stamp", Type = ExcelColumnType.TimestampColumn }];

            RecordBatch batch = reader.ToArrowRecordBatch(schema);

            var stamp = Assert.IsType<TimestampArray>(batch.Column(0));
            long expected = (new DateTime(2024, 1, 15, 13, 30, 0) - DateTime.UnixEpoch).Ticks / 10;
            Assert.Equal(expected, stamp.GetValue(0));
        }

        [Fact]
        public void ToArrowRecordBatch_Should_Convert_Timestamp_Column_With_Explicit_UtcOffset_Text()
        {
            using var reader = Excel.FromCsv(Encoding.UTF8.GetBytes("stamp\n2024-01-15T13:30:00Z\n"));
            ExcelColumnSchema[] schema = [new() { Index = 0, Name = "stamp", Type = ExcelColumnType.TimestampColumn }];

            RecordBatch batch = reader.ToArrowRecordBatch(schema);

            var stamp = Assert.IsType<TimestampArray>(batch.Column(0));
            long expected = (new DateTime(2024, 1, 15, 13, 30, 0, DateTimeKind.Utc) - DateTime.UnixEpoch).Ticks / 10;
            Assert.Equal(expected, stamp.GetValue(0));
        }

        [Fact]
        public void ToArrowRecordBatch_Should_Infer_Schema_When_None_Given()
        {
            using var reader = Excel.FromCsv(Encoding.UTF8.GetBytes("name,qty\nwidget,3\ngadget,7\n"));

            RecordBatch batch = reader.ToArrowRecordBatch();

            Assert.Equal(2, batch.ColumnCount);
            Assert.Equal("name", batch.Schema.GetFieldByIndex(0).Name);
            Assert.IsType<StringArray>(batch.Column(0));
            Assert.IsType<StringArray>(batch.Column(1));
        }

        [Fact]
        public void ToArrowRecordBatch_Should_Treat_HeaderRow_Zero_As_No_Header()
        {
            using var reader = Excel.FromCsv(Encoding.UTF8.GetBytes("widget,3\ngadget,7\n"));
            ExcelColumnSchema[] schema =
            [
                new() { Index = 0, Type = ExcelColumnType.StringColumn },
                new() { Index = 1, Type = ExcelColumnType.Int64Column },
            ];

            RecordBatch batch = reader.ToArrowRecordBatch(schema, headerRow: 0);

            Assert.Equal(2, batch.Length);
            var name = Assert.IsType<StringArray>(batch.Column(0));
            Assert.Equal("widget", name.GetString(0));
        }

        [Fact]
        public void ToArrowRecordBatch_Should_Match_The_Native_Arrow_Path_On_RealExcel()
        {
            string path = Path.Combine(AppContext.BaseDirectory, "data", "RealExcel.xlsb");

            using IExcelRowReader managedReader = Excel.Open(path);
            ExcelColumnSchema[] schema = Excel.InferSchema(managedReader);
            RecordBatch managedBatch = managedReader.ToArrowRecordBatch(schema, headerRow: 1);

            Assert.Equal(NativeStatus.Ok, NativeApiTests.OpenPath(path, NativeFormat.Xlsb, out NativeHandle? handle));
            try
            {
                NativeColumnSpec[] specs = new NativeColumnSpec[schema.Length];
                for (int i = 0; i < schema.Length; i++)
                {
                    specs[i] = new NativeColumnSpec
                    {
                        Names = schema[i].Name is null ? [] : [schema[i].Name!],
                        Index = schema[i].Index,
                        Type = (int)schema[i].Type,
                        Nullable = schema[i].IsNullable,
                    };
                }

                Assert.Equal(NativeStatus.Ok, NativeApi.ParseArrow(handle, specs, headerRow: 1, out ArrowArray nativeArray, out ArrowSchema nativeSchema));
                try
                {
                    Assert.Equal(managedBatch.Length, nativeArray.Length);
                    for (int i = 0; i < schema.Length; i++)
                    {
                        ArrowArray nativeColumn = NativeApiTests.ArrowChildArray(nativeArray, i);
                        Assert.Equal(nativeColumn.Length, managedBatch.Column(i).Length);
                        Assert.Equal(nativeColumn.NullCount, managedBatch.Column(i).NullCount);
                        AssertColumnValuesMatch(schema[i].Type, nativeColumn, managedBatch.Column(i));
                    }
                }
                finally
                {
                    NativeApiTests.ExercisedReleaseArrow(ref nativeArray, ref nativeSchema);
                }
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        private static void AssertColumnValuesMatch(ExcelColumnType type, ArrowArray nativeColumn, IArrowArray managedColumn)
        {
            for (int row = 0; row < nativeColumn.Length; row++)
            {
                if (managedColumn.IsNull(row))
                {
                    continue; // null positions already compared via NullCount equality above.
                }

                switch (type)
                {
                    case ExcelColumnType.StringColumn:
                        IntPtr offsets = NativeApiTests.ArrowBuffer(nativeColumn, 1);
                        IntPtr data = NativeApiTests.ArrowBuffer(nativeColumn, 2);
                        int start = Marshal.ReadInt32(offsets, row * sizeof(int));
                        int end = Marshal.ReadInt32(offsets, (row + 1) * sizeof(int));
                        byte[] bytes = new byte[end - start];
                        Marshal.Copy(IntPtr.Add(data, start), bytes, 0, bytes.Length);
                        Assert.Equal(Encoding.UTF8.GetString(bytes), ((StringArray)managedColumn).GetString(row));
                        break;
                    case ExcelColumnType.Int64Column:
                        Assert.Equal(Marshal.ReadInt64(NativeApiTests.ArrowBuffer(nativeColumn, 1), row * sizeof(long)), ((Int64Array)managedColumn).GetValue(row));
                        break;
                    case ExcelColumnType.Float64Column:
                        Assert.Equal(BitConverter.Int64BitsToDouble(Marshal.ReadInt64(NativeApiTests.ArrowBuffer(nativeColumn, 1), row * sizeof(double))), ((DoubleArray)managedColumn).GetValue(row));
                        break;
                    case ExcelColumnType.BoolColumn:
                        bool bit = (Marshal.ReadByte(NativeApiTests.ArrowBuffer(nativeColumn, 1), row / 8) & (1 << (row % 8))) != 0;
                        Assert.Equal(bit, ((BooleanArray)managedColumn).GetValue(row));
                        break;
                    case ExcelColumnType.DateColumn:
                        Assert.Equal(Marshal.ReadInt32(NativeApiTests.ArrowBuffer(nativeColumn, 1), row * sizeof(int)), ((Date32Array)managedColumn).GetValue(row));
                        break;
                    default: // TimeColumn, TimestampColumn — both 8-byte microsecond values on both paths.
                        Assert.Equal(Marshal.ReadInt64(NativeApiTests.ArrowBuffer(nativeColumn, 1), row * sizeof(long)), ((PrimitiveArray<long>)managedColumn).GetValue(row));
                        break;
                }
            }
        }
    }
}
