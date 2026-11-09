using System.Text;
using ExcelReader.Native;
using ExcelReader.Native.Reading;

namespace ExcelReader.Tests.Native
{
    public sealed partial class NativeApiTests
    {
        private static unsafe List<BorrowedCell> ReadCells(NativeCellView* cells, int count)
        {
            List<BorrowedCell> row = new(count);
            for (int index = 0; index < count; index++)
            {
                NativeCellView cell = cells[index];
                string text = cell.ValueLength == 0 ? "" : Encoding.UTF8.GetString(cell.Value, cell.ValueLength);
                row.Add(new BorrowedCell(cell.Column, cell.Type, text, cell.HasNumber != 0, cell.Number));
            }
            return row;
        }

        private static unsafe List<BorrowedCell>? NextCells(NativeRowCursor cursor)
        {
            int status = ReadApi.NextRowCells(cursor, out NativeCellView* cells, out int count);
            if (status == NativeStatus.Eof)
            {
                return null;
            }
            Assert.Equal(NativeStatus.Ok, status);
            return ReadCells(cells, count);
        }

        [Theory]
        [InlineData(NativeFormat.Xlsx)]
        [InlineData(NativeFormat.Xlsb)]
        [InlineData(NativeFormat.Xls)]
        public void NextRowCells_Should_Match_ForEachRow_On_Every_Row(int format)
        {
            string fixture = format switch
            {
                NativeFormat.Xlsb => XlsbFixture,
                NativeFormat.Xls => XlsFixture,
                _ => XlsxFixture,
            };
            Assert.Equal(NativeStatus.Ok, OpenPath(fixture, format, out NativeHandle? handle));
            try
            {
                ForEachSink sink = new();
                Assert.Equal(NativeStatus.Ok, ForEach(handle, sink));

                using NativeRowCursor cursor = OpenRows(handle);
                foreach (List<BorrowedCell> expected in sink.Rows)
                {
                    Assert.Equal(expected, NextCells(cursor));
                }
                Assert.True(sink.Rows.Count > 1);
                Assert.Null(NextCells(cursor));
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void NextRowCells_Should_Expose_A_Csv_Row_Exactly()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name,qty\nwidget,7\n");
            Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                using NativeRowCursor cursor = OpenRows(handle);
                Assert.Equal(["name", "qty"], NextCells(cursor)!.Select(cell => cell.Text), StringComparer.Ordinal);
                Assert.Equal(["widget", "7"], NextCells(cursor)!.Select(cell => cell.Text), StringComparer.Ordinal);
                Assert.Null(NextCells(cursor));
            }
            finally
            {
                ReadApi.Close(handle);
                File.Delete(path);
            }
        }

        [Fact]
        public void NextRowCells_Should_Pass_An_Xlsb_Number_As_Binary()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsbFixture, NativeFormat.Xlsb, out NativeHandle? handle));
            try
            {
                using NativeRowCursor cursor = OpenRows(handle);
                List<BorrowedCell> numbers = [];
                while (NextCells(cursor) is { } row)
                {
                    numbers.AddRange(row.Where(cell => cell.HasNumber));
                }
                Assert.NotEmpty(numbers);
                Assert.All(numbers, cell => Assert.Equal("", cell.Text));
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void NextRowCells_Should_Take_Over_A_Row_Pending_From_The_Blob_API()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name,qty\nwidget,7\n");
            Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                using NativeRowCursor cursor = OpenRows(handle);
                Assert.Equal(NativeStatus.BufferTooSmall, ReadApi.NextRow(cursor, Span<byte>.Empty, out _));

                Assert.Equal(
                    [new BorrowedCell(0, 1, "name", false, 0), new BorrowedCell(1, 1, "qty", false, 0)],
                    NextCells(cursor));
                Assert.Equal(["widget", "7"], NextCells(cursor)!.Select(cell => cell.Text), StringComparer.Ordinal);
            }
            finally
            {
                ReadApi.Close(handle);
                File.Delete(path);
            }
        }

        [Fact]
        public void NextRowCells_Should_Grow_For_A_Row_Wider_Than_Its_Buffers()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            string[] wide = Enumerable.Range(0, 300).Select(i => $"c{i}").ToArray();
            string big = new('x', 200_000);
            File.WriteAllText(path, $"a\n{string.Join(',', wide)}\n{big}\nb\n");
            Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                using NativeRowCursor cursor = OpenRows(handle);
                Assert.Equal("a", NextCells(cursor)![0].Text);
                Assert.Equal(wide, NextCells(cursor)!.Select(cell => cell.Text), StringComparer.Ordinal);
                Assert.Equal(big, NextCells(cursor)![0].Text);
                Assert.Equal("b", NextCells(cursor)![0].Text);
            }
            finally
            {
                ReadApi.Close(handle);
                File.Delete(path);
            }
        }

        [Fact]
        public unsafe void NextRowCells_Should_Return_A_Null_Cells_Pointer_For_An_Empty_Row()
        {
            using MemoryStream ms = WorkbookBuilder.Build("""<row r="1"/>""");
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                using NativeRowCursor cursor = OpenRows(handle);
                Assert.Equal(NativeStatus.Ok, ReadApi.NextRowCells(cursor, out NativeCellView* cells, out int count));
                Assert.Equal(0, count);
                Assert.True(cells is null);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public unsafe void NextRowCells_Should_Reject_A_Null_Handle()
        {
            Assert.Equal(NativeStatus.InvalidHandle, ReadApi.NextRowCells(null, out NativeCellView* cells, out int count));
            Assert.True(cells is null);
            Assert.Equal(0, count);
        }

        [Fact]
        public unsafe void NextRowCells_Export_Should_Reject_Null_Out_Pointers()
        {
            var export = (delegate* unmanaged<nint, NativeCellView**, int*, int>)&Exports.RowsNextCells;
            NativeCellView* cells;
            int count;
            Assert.Equal(NativeStatus.InvalidArgument, export(0, null, &count));
            Assert.Equal(NativeStatus.InvalidArgument, export(0, &cells, null));
        }
    }
}
