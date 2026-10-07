using System.Text;
using ExcelReader.Native;
using ExcelReader.Native.Arrow;
using ExcelReader.Native.Reading;
using ExcelReader.Native.Typed;
using ExcelReader.Tests.Reader;

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

            NativeApi.ClearLastError();
            Assert.Equal(expected, TypedApi.ParseTyped(handle, sheet, FirstColumnAsText(), headerRow: 0, out _));
            Assert.NotEmpty(NativeApi.LastErrorText());
            NativeApi.ClearLastError();
            Assert.Equal(expected, TypedApi.OpenTypedReader(handle, sheet, FirstColumnAsText(), headerRow: 0, maxRows: 0, out nint reader));
            Assert.Equal(0, reader);
            Assert.NotEmpty(NativeApi.LastErrorText());
            NativeApi.ClearLastError();
            Assert.Equal(expected, ArrowApi.OpenArrowStream(handle, sheet, FirstColumnAsText(), headerRow: 0, maxRows: 0, out _));
            Assert.NotEmpty(NativeApi.LastErrorText());
            NativeApi.ClearLastError();
            Assert.Equal(expected, ArrowApi.ParseArrow(handle, sheet, FirstColumnAsText(), headerRow: 0, out _, out _));
            Assert.NotEmpty(NativeApi.LastErrorText());
            NativeApi.ClearLastError();
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

        private static List<string> DrainFirstColumn(NativeRowCursor cursor)
        {
            List<string> values = [];
            byte[] buffer = new byte[4096];
            while (true)
            {
                int status = ReadApi.NextRow(cursor, buffer, out int written);
                if (status == NativeStatus.Eof)
                {
                    return values;
                }
                Assert.Equal(NativeStatus.Ok, status);
                values.Add(DecodeRow(buffer.AsSpan(0, written))[0].Value);
            }
        }

        [Fact]
        public async Task OpenRows_Should_Read_The_Sheet_At_The_Given_Index()
        {
            using NativeHandle handle = await OpenTwoSheetsAsync();
            using NativeRowCursor second = OpenRows(handle, 1);
            using NativeRowCursor first = OpenRows(handle, 0);

            Assert.Equal(["b1", "b2", "b3"], DrainFirstColumn(second));
            Assert.Equal(["a1", "a2"], DrainFirstColumn(first));
        }

        [Fact]
        public async Task Two_Cursors_On_The_Same_Sheet_Should_Each_Return_Every_Row()
        {
            using NativeHandle handle = await OpenTwoSheetsAsync();
            using NativeRowCursor one = OpenRows(handle, 1);
            using NativeRowCursor two = OpenRows(handle, 1);
            byte[] buffer = new byte[4096];

            Assert.Equal(NativeStatus.Ok, ReadApi.NextRow(one, buffer, out _));
            Assert.Equal(NativeStatus.Ok, ReadApi.NextRowView(two, out _));

            Assert.Equal(["b2", "b3"], DrainFirstColumn(one));
            Assert.Equal(["b2", "b3"], DrainFirstColumn(two));
        }

        [Theory]
        [InlineData(-1, NativeStatus.InvalidArgument)]
        [InlineData(2, NativeStatus.Error)]
        public async Task OpenRows_And_SheetVisibilityAt_Should_Reject_An_Index_Out_Of_Range(int sheet, int expected)
        {
            using NativeHandle handle = await OpenTwoSheetsAsync();

            Assert.Equal(expected, ReadApi.OpenRows(handle, sheet, out NativeRowCursor? cursor));
            Assert.Null(cursor);
            Assert.Equal(expected, ReadApi.SheetVisibilityAt(handle, sheet, out _));
            Assert.NotEmpty(NativeApi.LastErrorText());
            NativeApi.ClearLastError();
            Assert.Equal(expected, ReadApi.SheetNameAt(handle, sheet, new byte[64], out _));
            Assert.NotEmpty(NativeApi.LastErrorText());
        }

        [Fact]
        public async Task A_Cursor_Opened_But_Not_Yet_Read_Should_Read_Every_Row_After_Its_Workbook_Is_Closed()
        {
            NativeHandle handle = await OpenTwoSheetsAsync();
            using NativeRowCursor cursor = OpenRows(handle, 1);

            Assert.Equal(NativeStatus.Ok, ReadApi.Close(handle));

            Assert.Equal(["b1", "b2", "b3"], DrainFirstColumn(cursor));
            Assert.Equal(NativeStatus.Error, ReadApi.OpenRows(handle, 0, out _));
        }

        [Theory]
        [InlineData("Second", 1)]
        [InlineData("sEcOnD", 1)]
        [InlineData("First", 0)]
        [InlineData("Missing", -1)]
        [InlineData("", -1)]
        public async Task SheetIndex_Should_Match_Names_Without_Regard_To_Case_And_Report_Minus_One_When_Absent(string name, int expected)
        {
            using NativeHandle handle = await OpenTwoSheetsAsync();

            Assert.Equal(NativeStatus.Ok, ReadApi.SheetIndex(handle, Encoding.UTF8.GetBytes(name), out int index));

            Assert.Equal(expected, index);
        }

        [Fact]
        public void SheetIndex_Should_Find_The_Unnamed_Sheet_Of_A_Csv()
        {
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory("a,b\n1,2\n"u8, NativeFormat.Csv, out NativeHandle? handle));
            using NativeHandle live = handle!;

            Assert.Equal(NativeStatus.Ok, ReadApi.SheetIndex(live, [], out int index));

            Assert.Equal(0, index);
        }

        [Fact]
        public void SheetVisibilityAt_Should_Report_Visible_Hidden_And_VeryHidden()
        {
            byte[] xlsx = SheetVisibilityTests.BuildXlsx(null, "hidden", "veryHidden");
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(xlsx, NativeFormat.Xlsx, out NativeHandle? handle));
            using NativeHandle live = handle!;

            for (int i = 0; i < 3; i++)
            {
                Assert.Equal(NativeStatus.Ok, ReadApi.SheetVisibilityAt(live, i, out int visibility));
                Assert.Equal(i, visibility);
            }
        }

        [Fact]
        public async Task A_Handle_Id_Should_Resolve_Only_As_Its_Own_Kind()
        {
            using NativeHandle handle = await OpenTwoSheetsAsync();
            NativeRowCursor cursor = OpenRows(handle);
            nint workbookId = NativeHandleTable.Register(handle);
            nint cursorId = NativeHandleTable.Register(cursor);

            Assert.Null(Exports.ResolveRows(workbookId));
            Assert.Null(Exports.Resolve(cursorId));
            Assert.False(Exports.TryFree(cursorId, out _));
            Assert.Same(cursor, Exports.ResolveRows(cursorId));

            Assert.True(NativeHandleTable.TryUnregister(workbookId, out NativeHandle? _));
            Assert.True(NativeHandleTable.TryUnregister(cursorId, out NativeRowCursor? _));
            cursor.Dispose();
        }

        [Fact]
        public async Task A_Cursor_Should_Close_Once_With_A_Read_All_Result_Still_Pending()
        {
            using NativeHandle handle = await OpenTwoSheetsAsync();
            NativeRowCursor cursor = OpenRows(handle, 1);
            nint id = NativeHandleTable.Register(cursor);
            Assert.Equal(NativeStatus.BufferTooSmall, ReadApi.ReadAllBlob(cursor, Span<byte>.Empty, out int needed));
            Assert.True(needed > 0);

            Assert.True(NativeHandleTable.TryUnregister(id, out NativeRowCursor? first));
            Assert.Equal(NativeStatus.Ok, ReadApi.CloseRows(first));

            Assert.False(NativeHandleTable.TryUnregister(id, out NativeRowCursor? second));
            Assert.Equal(NativeStatus.InvalidHandle, ReadApi.CloseRows(second));
        }
    }
}
