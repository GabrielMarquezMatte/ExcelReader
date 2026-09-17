using ExcelReader.Core.Reader;
using ExcelReader.Native;

namespace ExcelReader.Tests
{
    public sealed unsafe class CsvAggregateNativeTests
    {
        [Fact]
        public void Translate_Should_Return_Defaults_When_Options_Are_Absent()
        {
            int status = NativeCsvAggregateOptions.Translate(null, out CsvParallelOptions options);

            Assert.Equal(NativeStatus.Ok, status);
            Assert.Equal(0, options.DegreeOfParallelism);
            Assert.Equal(0, options.HeaderRow);
            Assert.Equal((byte)',', options.Reader.Delimiter);
            Assert.Equal((byte)'"', options.Reader.Quote);
            Assert.True(options.Reader.DetectEncodingFromByteOrderMark);
            Assert.Equal(32 * 1024 * 1024, options.Reader.MaxCellBytes);
        }

        [Fact]
        public void Translate_Should_Apply_Dialect_When_Fields_Are_Set()
        {
            NativeCsvParallelOptionsRaw raw = new()
            {
                StructSize = sizeof(NativeCsvParallelOptionsRaw),
                DegreeOfParallelism = 4,
                HeaderRow = 1,
                Delimiter = ';',
                Quote = '\'',
                DetectBom = 1,
                MaxCellBytes = 4096,
            };

            int status = NativeCsvAggregateOptions.Translate(raw, out CsvParallelOptions options);

            Assert.Equal(NativeStatus.Ok, status);
            Assert.Equal(4, options.DegreeOfParallelism);
            Assert.Equal(1, options.HeaderRow);
            Assert.Equal((byte)';', options.Reader.Delimiter);
            Assert.Equal((byte)'\'', options.Reader.Quote);
            Assert.False(options.Reader.DetectEncodingFromByteOrderMark);
            Assert.Equal(4096, options.Reader.MaxCellBytes);
        }

        [Theory]
        [InlineData(-1, 0, 0, 0, 0)]
        [InlineData(0, -1, 0, 0, 0)]
        [InlineData(0, 0, 256, 0, 0)]
        [InlineData(0, 0, 0, 256, 0)]
        [InlineData(0, 0, 0, 0, -1)]
        public void Translate_Should_Reject_When_A_Field_Is_Out_Of_Range(
            int dop, int headerRow, int delimiter, int quote, int maxCellBytes)
        {
            NativeCsvParallelOptionsRaw raw = new()
            {
                StructSize = sizeof(NativeCsvParallelOptionsRaw),
                DegreeOfParallelism = dop,
                HeaderRow = headerRow,
                Delimiter = delimiter,
                Quote = quote,
                MaxCellBytes = maxCellBytes,
            };

            int status = NativeCsvAggregateOptions.Translate(raw, out _);

            Assert.Equal(NativeStatus.InvalidArgument, status);
        }

        [Fact]
        public void Translate_Should_Reject_When_StructSize_Is_Too_Small()
        {
            NativeCsvParallelOptionsRaw raw = new() { StructSize = 4 };

            int status = NativeCsvAggregateOptions.Translate(raw, out _);

            Assert.Equal(NativeStatus.InvalidArgument, status);
        }
    }
}
