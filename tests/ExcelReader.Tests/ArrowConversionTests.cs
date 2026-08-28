using System.Text;
using Apache.Arrow;
using ExcelReader.Arrow;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;

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
    }
}
