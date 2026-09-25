using System.Runtime.InteropServices;
using ExcelReader.Native;
using ExcelReader.Native.Reading;

namespace ExcelReader.Tests.Native
{
    public sealed partial class NativeApiTests
    {
        private static List<DecodedCell> ReadView(NativeRow row)
        {
            List<DecodedCell> cells = [];
            int cellSize = Marshal.SizeOf<NativeRowCell>();
            for (int index = 0; index < row.CellCount; index++)
            {
                NativeRowCell cell = Marshal.PtrToStructure<NativeRowCell>(IntPtr.Add(row.Cells, index * cellSize));
                Assert.Equal((byte)0, Marshal.ReadByte(cell.Value, cell.ValueLength));
                cells.Add(new DecodedCell(cell.Column, cell.Type, Marshal.PtrToStringUTF8(cell.Value, cell.ValueLength)));
            }
            return cells;
        }

        [Fact]
        public void NextRowView_Should_Expose_A_Csv_Row_Exactly()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name,qty\nwidget,7\n");
            Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, ReadApi.NextRowView(handle, out NativeRow header));
                Assert.Equal([new DecodedCell(0, 1, "name"), new DecodedCell(1, 1, "qty")], ReadView(header));

                Assert.Equal(NativeStatus.Ok, ReadApi.NextRowView(handle, out NativeRow data));
                Assert.Equal(["widget", "7"], ReadView(data).Select(cell => cell.Value), StringComparer.Ordinal);

                Assert.Equal(NativeStatus.Eof, ReadApi.NextRowView(handle, out NativeRow end));
                Assert.Equal(IntPtr.Zero, end.Cells);
            }
            finally
            {
                ReadApi.Close(handle);
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData(NativeFormat.Xlsx)]
        [InlineData(NativeFormat.Xlsb)]
        public void NextRowView_Should_Match_NextRow_On_Every_Row(int format)
        {
            string fixture = format == NativeFormat.Xlsb ? XlsbFixture : XlsxFixture;
            Assert.Equal(NativeStatus.Ok, OpenPath(fixture, format, out NativeHandle? blobHandle));
            Assert.Equal(NativeStatus.Ok, OpenPath(fixture, format, out NativeHandle? viewHandle));
            try
            {
                byte[] buffer = new byte[1 << 20];
                int rows = 0;
                while (ReadApi.NextRow(blobHandle, buffer, out int written) == NativeStatus.Ok)
                {
                    Assert.Equal(NativeStatus.Ok, ReadApi.NextRowView(viewHandle, out NativeRow row));
                    Assert.Equal(DecodeRow(buffer.AsSpan(0, written)), ReadView(row));
                    rows++;
                }
                Assert.True(rows > 1);
                Assert.Equal(NativeStatus.Eof, ReadApi.NextRowView(viewHandle, out _));
            }
            finally
            {
                ReadApi.Close(blobHandle);
                ReadApi.Close(viewHandle);
            }
        }

        [Fact]
        public void NextRowView_Should_Take_Over_A_Row_Pending_From_The_Blob_API()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name,qty\nwidget,7\n");
            Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.BufferTooSmall, ReadApi.NextRow(handle, Span<byte>.Empty, out _));

                Assert.Equal(NativeStatus.Ok, ReadApi.NextRowView(handle, out NativeRow header));
                Assert.Equal(["name", "qty"], ReadView(header).Select(cell => cell.Value), StringComparer.Ordinal);
                Assert.Equal(NativeStatus.Ok, ReadApi.NextRowView(handle, out NativeRow data));
                Assert.Equal(["widget", "7"], ReadView(data).Select(cell => cell.Value), StringComparer.Ordinal);
            }
            finally
            {
                ReadApi.Close(handle);
                File.Delete(path);
            }
        }

        [Fact]
        public void NextRowView_Should_Grow_For_A_Row_Wider_Than_Its_Buffers()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            string[] wide = Enumerable.Range(0, 300).Select(i => $"c{i}").ToArray();
            string big = new('x', 200_000);
            File.WriteAllText(path, $"a\n{string.Join(',', wide)}\n{big}\nb\n");
            Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, ReadApi.NextRowView(handle, out NativeRow first));
                Assert.Equal("a", ReadView(first)[0].Value);
                Assert.Equal(NativeStatus.Ok, ReadApi.NextRowView(handle, out NativeRow many));
                Assert.Equal(wide, ReadView(many).Select(cell => cell.Value), StringComparer.Ordinal);
                Assert.Equal(NativeStatus.Ok, ReadApi.NextRowView(handle, out NativeRow large));
                Assert.Equal(big, ReadView(large)[0].Value);
                Assert.Equal(NativeStatus.Ok, ReadApi.NextRowView(handle, out NativeRow last));
                Assert.Equal("b", ReadView(last)[0].Value);
            }
            finally
            {
                ReadApi.Close(handle);
                File.Delete(path);
            }
        }

        [Fact]
        public void NextRowView_Should_Return_A_Null_Cells_Pointer_For_An_Empty_Row()
        {
            using MemoryStream ms = WorkbookBuilder.Build("""<row r="1"/>""");
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, ReadApi.NextRowView(handle, out NativeRow row));
                Assert.Equal(0, row.CellCount);
                Assert.Equal(IntPtr.Zero, row.Cells);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void NextRowView_Should_Reject_A_Null_Handle()
        {
            Assert.Equal(NativeStatus.InvalidHandle, ReadApi.NextRowView(null, out NativeRow row));
            Assert.Equal(IntPtr.Zero, row.Cells);
        }

        [Fact]
        public unsafe void NextRowView_Export_Should_Reject_A_Null_Out_Row()
        {
            int status = ((delegate* unmanaged<nint, NativeRow*, int>)&Exports.NextRowView)(0, null);
            Assert.Equal(NativeStatus.InvalidArgument, status);
        }
    }
}
