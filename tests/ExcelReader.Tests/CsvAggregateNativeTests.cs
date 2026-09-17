using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
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
        [InlineData(0, 0, 0, 0, 0, 3)]
        public void Translate_Should_Reject_When_A_Field_Is_Out_Of_Range(
            int dop, int headerRow, int delimiter, int quote, int maxCellBytes, int detectBom = 0)
        {
            NativeCsvParallelOptionsRaw raw = new()
            {
                StructSize = sizeof(NativeCsvParallelOptionsRaw),
                DegreeOfParallelism = dop,
                HeaderRow = headerRow,
                Delimiter = delimiter,
                Quote = quote,
                DetectBom = detectBom,
                MaxCellBytes = maxCellBytes,
            };

            int status = NativeCsvAggregateOptions.Translate(raw, out _);

            Assert.Equal(NativeStatus.InvalidArgument, status);
        }

        [Theory]
        [InlineData(4)]
        [InlineData(1024)]
        public void Translate_Should_Reject_When_StructSize_Is_Invalid(int structSize)
        {
            NativeCsvParallelOptionsRaw raw = new() { StructSize = structSize };

            int status = NativeCsvAggregateOptions.Translate(raw, out _);

            Assert.Equal(NativeStatus.InvalidArgument, status);
        }

        private static Row FirstRowOf(string csv)
        {
            CsvReader reader = Excel.FromCsv(Encoding.UTF8.GetBytes(csv), CsvReaderOptions.Default);
            CsvReader.Enumerator rows = reader.GetEnumerator();
            Assert.True(rows.MoveNext());
            return rows.Current;
        }

        [Fact]
        public void WriteRow_Should_Expose_Every_Field_As_A_NulTerminated_Cell()
        {
            using CsvAggregateState state = new();

            NativeRow row = state.WriteRow(FirstRowOf("alpha,beta,gamma\n"));

            Assert.Equal(3, row.CellCount);
            NativeRowCell* cells = (NativeRowCell*)row.Cells;
            Assert.Equal("alpha", ReadCell(cells[0]));
            Assert.Equal("beta", ReadCell(cells[1]));
            Assert.Equal("gamma", ReadCell(cells[2]));
            Assert.Equal(0, cells[0].Column);
            Assert.Equal(2, cells[2].Column);
            Assert.Equal(1, cells[0].Type);
            Assert.Equal(0, ((byte*)cells[0].Value)[cells[0].ValueLength]);
        }

        [Fact]
        public void WriteRow_Should_Report_Empty_Type_When_A_Field_Is_Blank()
        {
            using CsvAggregateState state = new();

            NativeRow row = state.WriteRow(FirstRowOf("alpha,,gamma\n"));

            NativeRowCell* cells = (NativeRowCell*)row.Cells;
            Assert.Equal(3, row.CellCount);
            Assert.Equal(0, cells[1].Type);
            Assert.Equal(0, cells[1].ValueLength);
        }

        [Fact]
        public void WriteRow_Should_Reuse_Its_Buffer_When_A_Later_Row_Is_Smaller()
        {
            using CsvAggregateState state = new();

            NativeRow wide = state.WriteRow(FirstRowOf("aaaaaaaaaa,bbbbbbbbbb,cccccccccc\n"));
            IntPtr firstBuffer = wide.Cells;
            NativeRow narrow = state.WriteRow(FirstRowOf("x\n"));

            Assert.Equal(firstBuffer, narrow.Cells);
            Assert.Equal(1, narrow.CellCount);
            Assert.Equal("x", ReadCell(*(NativeRowCell*)narrow.Cells));
        }

        private static string ReadCell(NativeRowCell cell)
        {
            return Encoding.UTF8.GetString((byte*)cell.Value, cell.ValueLength);
        }
    }
}
