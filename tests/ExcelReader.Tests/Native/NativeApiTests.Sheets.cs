using System.Text;
using ExcelReader.Native;
using ExcelReader.Native.Arrow;
using ExcelReader.Native.Reading;
using ExcelReader.Native.Typed;

namespace ExcelReader.Tests.Native
{
    public sealed partial class NativeApiTests
    {
        private static NativeColumnSpec[] FirstColumnAsText()
        {
            return [new() { Index = 0, Type = NativeColumnType.String, Nullable = true }];
        }

        private static async Task<NativeHandle> OpenTwoSheetsAsync()
        {
            await using MemoryStream ms = await TypedWorkbook.BuildMultiSheetAsync(
                ("First", new object?[][] { ["a1"], ["a2"] }),
                ("Second", new object?[][] { ["b1"], ["b2"], ["b3"] }));
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            return handle!;
        }

        [Fact]
        public async Task ParseTyped_Should_Read_The_Sheet_At_The_Given_Index()
        {
            using NativeHandle handle = await OpenTwoSheetsAsync();

            Assert.Equal(NativeStatus.Ok, TypedApi.ParseTyped(handle, 1, FirstColumnAsText(), headerRow: 0, out NativeTable table));
            try
            {
                Assert.Equal(["b1", "b2", "b3"], DecodeStringColumn(ColumnAt(table, 0)));
            }
            finally
            {
                TypedApi.FreeTable(ref table);
            }
        }

        [Theory]
        [InlineData(-1, NativeStatus.InvalidArgument)]
        [InlineData(2, NativeStatus.Error)]
        public async Task EverySheetRead_Should_Reject_A_Sheet_Index_Out_Of_Range(int sheet, int expected)
        {
            using NativeHandle handle = await OpenTwoSheetsAsync();

            Assert.Equal(expected, TypedApi.ParseTyped(handle, sheet, FirstColumnAsText(), headerRow: 0, out _));
            Assert.NotEmpty(NativeApi.LastErrorText());
            Assert.Equal(expected, TypedApi.OpenTypedReader(handle, sheet, FirstColumnAsText(), headerRow: 0, maxRows: 0, out nint reader));
            Assert.Equal(0, reader);
            Assert.Equal(expected, ArrowApi.OpenArrowStream(handle, sheet, FirstColumnAsText(), headerRow: 0, maxRows: 0, out _));
            Assert.Equal(expected, ArrowApi.ParseArrow(handle, sheet, FirstColumnAsText(), headerRow: 0, out _, out _));
            Assert.Equal(expected, ReadApi.InferSchema(handle, sheet, headerRow: 0, sampleSize: 10, flags: 0, out _));
            Assert.NotEmpty(NativeApi.LastErrorText());
        }

        [Fact]
        public void ParseTyped_Should_Reject_Sheet_One_Of_A_Csv_Even_When_Asked_To_Run_In_Parallel()
        {
            StringBuilder csv = new("id\n");
            for (int i = 0; i < 200_000; i++)
            {
                csv.Append(i).Append('\n');
            }
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(Encoding.UTF8.GetBytes(csv.ToString()), NativeFormat.Csv, out NativeHandle? handle));
            using NativeHandle live = handle!;

            int status = TypedApi.ParseTypedTable(live, 1, FirstColumnAsText(), headerRow: 1, degreeOfParallelism: 0, out NativeTable table);

            Assert.Equal(NativeStatus.Error, status);
            Assert.Equal(IntPtr.Zero, table.Columns);
        }
    }
}
