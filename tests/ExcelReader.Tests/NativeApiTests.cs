using System.Text;
using ExcelReader.Native;

namespace ExcelReader.Tests
{
    public sealed class NativeApiTests
    {
        private static readonly string XlsxFixture = Path.Combine(AppContext.BaseDirectory, "data", "sample.xlsx");
        private static readonly string XlsbFixture = Path.Combine(AppContext.BaseDirectory, "data", "RealExcel.xlsb");

        private static int OpenPath(string path, int format, out NativeHandle? handle)
        {
            return NativeApi.OpenFile(Encoding.UTF8.GetBytes(path), format, out handle);
        }

        [Fact]
        public void LastError_Should_Return_Stored_Message_As_Utf8()
        {
            NativeApi.SetLastError("boom");

            Span<byte> buffer = stackalloc byte[64];
            int status = NativeApi.LastError(buffer, out int length);

            Assert.Equal(NativeStatus.Ok, status);
            Assert.Equal(4, length);
            Assert.Equal("boom", Encoding.UTF8.GetString(buffer[..length]));
        }

        [Fact]
        public void LastError_Should_Report_Required_Size_When_Buffer_Too_Small()
        {
            NativeApi.SetLastError("boom");

            Span<byte> buffer = stackalloc byte[2];
            int status = NativeApi.LastError(buffer, out int length);

            Assert.Equal(NativeStatus.BufferTooSmall, status);
            Assert.Equal(4, length);
        }

        [Fact]
        public void LastError_Should_Return_Zero_Length_When_Cleared()
        {
            NativeApi.SetLastError("boom");
            NativeApi.ClearLastError();

            Span<byte> buffer = stackalloc byte[64];
            int status = NativeApi.LastError(buffer, out int length);

            Assert.Equal(NativeStatus.Ok, status);
            Assert.Equal(0, length);
        }

        [Theory]
        [InlineData(NativeFormat.Auto)]
        [InlineData(NativeFormat.Xlsx)]
        public void OpenFile_Should_Open_Xlsx(int format)
        {
            int status = OpenPath(XlsxFixture, format, out NativeHandle? handle);

            Assert.Equal(NativeStatus.Ok, status);
            Assert.NotNull(handle);
            Assert.Equal(NativeStatus.Ok, NativeApi.Close(handle));
        }

        [Theory]
        [InlineData(NativeFormat.Auto)]
        [InlineData(NativeFormat.Xlsb)]
        public void OpenFile_Should_Open_Xlsb(int format)
        {
            int status = OpenPath(XlsbFixture, format, out NativeHandle? handle);

            Assert.Equal(NativeStatus.Ok, status);
            Assert.NotNull(handle);
            Assert.Equal(NativeStatus.Ok, NativeApi.Close(handle));
        }

        [Fact]
        public void OpenFile_Should_Open_Csv_When_Format_Is_Explicit()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "a,b\n1,2\n");
            try
            {
                int status = OpenPath(path, NativeFormat.Csv, out NativeHandle? handle);

                Assert.Equal(NativeStatus.Ok, status);
                Assert.NotNull(handle);
                Assert.Equal(NativeStatus.Ok, NativeApi.Close(handle));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void OpenFile_Should_Fail_With_Error_When_File_Is_Missing()
        {
            string path = Path.Combine(Path.GetTempPath(), $"does-not-exist-{Guid.NewGuid():N}.xlsx");

            int status = OpenPath(path, NativeFormat.Auto, out NativeHandle? handle);

            Assert.Equal(NativeStatus.Error, status);
            Assert.Null(handle);

            Span<byte> buffer = stackalloc byte[512];
            Assert.Equal(NativeStatus.Ok, NativeApi.LastError(buffer, out int length));
            Assert.True(length > 0);
        }

        [Fact]
        public void OpenFile_Should_Reject_Unknown_Format_Code()
        {
            int status = OpenPath(XlsxFixture, 99, out NativeHandle? handle);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Null(handle);
        }

        [Fact]
        public void OpenMemory_Should_Open_Xlsx_From_A_Copy_Of_The_Bytes()
        {
            byte[] bytes = File.ReadAllBytes(XlsxFixture);

            int status = NativeApi.OpenMemory(bytes, NativeFormat.Auto, out NativeHandle? handle);

            Assert.Equal(NativeStatus.Ok, status);
            Assert.NotNull(handle);
            Assert.Equal(NativeStatus.Ok, NativeApi.Close(handle));
        }

        [Fact]
        public void Close_Should_Reject_A_Null_Handle()
        {
            Assert.Equal(NativeStatus.InvalidHandle, NativeApi.Close(null));
        }

        [Fact]
        public void SheetCount_Should_Report_At_Least_One_Sheet()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                int status = NativeApi.SheetCount(handle, out int count);

                Assert.Equal(NativeStatus.Ok, status);
                Assert.True(count >= 1);
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void SheetName_Should_Return_A_Non_Empty_Utf8_Name()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                Span<byte> buffer = stackalloc byte[256];
                int status = NativeApi.SheetName(handle, buffer, out int length);

                Assert.Equal(NativeStatus.Ok, status);
                Assert.True(length > 0);
                Assert.NotEmpty(Encoding.UTF8.GetString(buffer[..length]));
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void SheetName_Should_Report_Required_Size_When_Buffer_Too_Small()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                int status = NativeApi.SheetName(handle, Span<byte>.Empty, out int length);

                Assert.Equal(NativeStatus.BufferTooSmall, status);
                Assert.True(length > 0);
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void MoveToSheet_Should_Accept_The_First_Sheet()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, NativeApi.MoveToSheet(handle, 0));
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void MoveToSheet_Should_Fail_For_An_Out_Of_Range_Index()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Error, NativeApi.MoveToSheet(handle, 9999));
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void IsDate1904_Should_Report_Zero_For_A_1900_Based_Workbook()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                int status = NativeApi.IsDate1904(handle, out int flag);

                Assert.Equal(NativeStatus.Ok, status);
                Assert.Equal(0, flag);
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void Sheet_Functions_Should_Reject_A_Null_Handle()
        {
            Assert.Equal(NativeStatus.InvalidHandle, NativeApi.SheetCount(null, out _));
            Assert.Equal(NativeStatus.InvalidHandle, NativeApi.MoveToSheet(null, 0));
            Assert.Equal(NativeStatus.InvalidHandle, NativeApi.IsDate1904(null, out _));
        }

        private sealed record DecodedCell(int Column, int Type, string Value);

        private static List<DecodedCell> DecodeRow(ReadOnlySpan<byte> blob)
        {
            List<DecodedCell> cells = [];
            int count = BitConverter.ToInt32(blob[..4]);
            int offset = 4;
            for (int i = 0; i < count; i++)
            {
                int column = BitConverter.ToInt32(blob[offset..]);
                int type = BitConverter.ToInt32(blob[(offset + 4)..]);
                int valueLength = BitConverter.ToInt32(blob[(offset + 8)..]);
                offset += 12;
                cells.Add(new DecodedCell(column, type, Encoding.UTF8.GetString(blob.Slice(offset, valueLength))));
                offset += valueLength;
            }

            return cells;
        }

        [Fact]
        public void NextRow_Should_Decode_A_Csv_Row_Exactly()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name,qty\nwidget,7\n");
            Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                byte[] buffer = new byte[4096];

                Assert.Equal(NativeStatus.Ok, NativeApi.NextRow(handle, buffer, out int first));
                List<DecodedCell> header = DecodeRow(buffer.AsSpan(0, first));
                Assert.Equal(2, header.Count);
                Assert.Equal(0, header[0].Column);
                Assert.Equal("name", header[0].Value);
                Assert.Equal(1, header[1].Column);
                Assert.Equal("qty", header[1].Value);

                Assert.Equal(NativeStatus.Ok, NativeApi.NextRow(handle, buffer, out int second));
                List<DecodedCell> data = DecodeRow(buffer.AsSpan(0, second));
                Assert.Equal("widget", data[0].Value);
                Assert.Equal("7", data[1].Value);

                Assert.Equal(NativeStatus.Eof, NativeApi.NextRow(handle, buffer, out _));
            }
            finally
            {
                NativeApi.Close(handle);
                File.Delete(path);
            }
        }

        [Fact]
        public void NextRow_Should_Reserve_The_Row_When_The_Buffer_Is_Too_Small()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name,qty\n");
            Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                byte[] tiny = new byte[3];
                Assert.Equal(NativeStatus.BufferTooSmall, NativeApi.NextRow(handle, tiny, out int required));
                Assert.True(required > 3);

                // The same row must come back — a caller that grows its buffer must not lose data.
                byte[] big = new byte[required];
                Assert.Equal(NativeStatus.Ok, NativeApi.NextRow(handle, big, out int written));
                Assert.Equal(required, written);
                Assert.Equal("name", DecodeRow(big.AsSpan(0, written))[0].Value);
            }
            finally
            {
                NativeApi.Close(handle);
                File.Delete(path);
            }
        }

        [Fact]
        public void NextRow_Should_Read_Every_Row_Of_The_Xlsx_Fixture()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                byte[] buffer = new byte[1 << 20];
                int rows = 0;
                while (NativeApi.NextRow(handle, buffer, out int written) == NativeStatus.Ok)
                {
                    Assert.True(written >= 4);
                    rows++;
                }

                Assert.True(rows > 0);
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void NextRow_Should_Restart_After_MoveToSheet()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                byte[] buffer = new byte[1 << 20];
                Assert.Equal(NativeStatus.Ok, NativeApi.NextRow(handle, buffer, out int firstPass));

                Assert.Equal(NativeStatus.Ok, NativeApi.MoveToSheet(handle, 0));

                Assert.Equal(NativeStatus.Ok, NativeApi.NextRow(handle, buffer, out int secondPass));
                Assert.Equal(firstPass, secondPass);
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void NextRow_Should_Reject_A_Null_Handle()
        {
            Assert.Equal(NativeStatus.InvalidHandle, NativeApi.NextRow(null, new byte[16], out _));
        }
    }
}
