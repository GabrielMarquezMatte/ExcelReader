using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Native;
using ExcelReader.Native.Reading;

namespace ExcelReader.Tests.Native
{
    public sealed partial class NativeApiTests
    {
        private static readonly string XlsFixture = Path.Combine(AppContext.BaseDirectory, "data", "excel-formula-strings.xls");

        private sealed record BorrowedCell(int Column, int Type, string Text, bool HasNumber, double Number);

        private sealed class ForEachSink
        {
            internal readonly List<List<BorrowedCell>> Rows = [];
            internal int StopAfter = int.MaxValue;
        }

        [UnmanagedCallersOnly]
        private static unsafe int Collect(NativeCellView* cells, int count, void* userData)
        {
            ForEachSink sink = (ForEachSink)GCHandle.FromIntPtr((nint)userData).Target!;
            List<BorrowedCell> row = new(count);
            for (int index = 0; index < count; index++)
            {
                NativeCellView cell = cells[index];
                string text = cell.ValueLength == 0 ? "" : Encoding.UTF8.GetString(cell.Value, cell.ValueLength);
                row.Add(new BorrowedCell(cell.Column, cell.Type, text, cell.HasNumber != 0, cell.Number));
            }
            sink.Rows.Add(row);
            return sink.Rows.Count >= sink.StopAfter ? 42 : NativeStatus.Ok;
        }

        private static unsafe int ForEach(NativeHandle? handle, ForEachSink sink)
        {
            GCHandle pinned = GCHandle.Alloc(sink);
            try
            {
                return ReadApi.ForEachRow(handle, 0, &Collect, (void*)GCHandle.ToIntPtr(pinned));
            }
            finally
            {
                pinned.Free();
            }
        }

        [Theory]
        [InlineData(NativeFormat.Xlsx)]
        [InlineData(NativeFormat.Xlsb)]
        [InlineData(NativeFormat.Xls)]
        public void ForEachRow_Should_Match_NextRowView_On_Every_Row(int format)
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
                foreach (List<BorrowedCell> borrowed in sink.Rows)
                {
                    Assert.Equal(NativeStatus.Ok, ReadApi.NextRowView(cursor, out NativeRow row));
                    List<DecodedCell> expected = ReadView(row);
                    Assert.Equal(expected.Count, borrowed.Count);
                    for (int index = 0; index < expected.Count; index++)
                    {
                        BorrowedCell cell = borrowed[index];
                        Assert.Equal(expected[index].Column, cell.Column);
                        Assert.Equal(expected[index].Type, cell.Type);
                        if (cell.Text.Length == 0 && cell.HasNumber)
                        {
                            Assert.Equal(double.Parse(expected[index].Value, CultureInfo.InvariantCulture), cell.Number);
                        }
                        else
                        {
                            Assert.Equal(expected[index].Value, cell.Text);
                        }
                    }
                }
                Assert.True(sink.Rows.Count > 1);
                Assert.Equal(NativeStatus.Eof, ReadApi.NextRowView(cursor, out _));
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void ForEachRow_Should_Stop_And_Return_The_Callback_Status()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                ForEachSink sink = new() { StopAfter = 2 };
                Assert.Equal(42, ForEach(handle, sink));
                Assert.Equal(2, sink.Rows.Count);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void ForEachRow_Should_Reject_A_Null_Handle()
        {
            Assert.Equal(NativeStatus.InvalidHandle, ForEach(null, new ForEachSink()));
        }

        [Fact]
        public unsafe void ForEachRow_Export_Should_Reject_A_Null_Callback()
        {
            int status = ((delegate* unmanaged<nint, int, delegate* unmanaged<NativeCellView*, int, void*, int>, void*, int>)&Exports.RowsForEach)(0, 0, null, null);
            Assert.Equal(NativeStatus.InvalidArgument, status);
        }
    }
}
