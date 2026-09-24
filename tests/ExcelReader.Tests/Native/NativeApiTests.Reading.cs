using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core;
using ExcelReader.Core.Reader;
using ExcelReader.Native;
using ExcelReader.Native.Reading;
using ExcelReader.Tests.Crypto;

namespace ExcelReader.Tests.Native
{
    public sealed partial class NativeApiTests
    {
        private static readonly string XlsbFixture = Path.Combine(AppContext.BaseDirectory, "data", "RealExcel.xlsb");

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

        private static List<List<DecodedCell>> DecodeAllRowsBlob(ReadOnlySpan<byte> blob)
        {
            List<List<DecodedCell>> rows = [];
            int rowCount = BitConverter.ToInt32(blob[..4]);
            int offset = 4;
            for (int i = 0; i < rowCount; i++)
            {
                int rowLength = BitConverter.ToInt32(blob[offset..]);
                offset += 4;
                rows.Add(DecodeRow(blob.Slice(offset, rowLength)));
                offset += rowLength;
            }

            return rows;
        }

        private static List<string> DecodeStringColumn(NativeColumn column)
        {
            int rowCount = (int)column.Length;
            int[] offsets = new int[rowCount + 1];
            Marshal.Copy(column.Values, offsets, 0, rowCount + 1);
            byte[] data = new byte[column.DataLen];
            if (data.Length > 0)
            {
                Marshal.Copy(column.Data, data, 0, data.Length);
            }
            List<string> values = [];
            for (int i = 0; i < rowCount; i++)
            {
                values.Add(Encoding.UTF8.GetString(data, offsets[i], offsets[i + 1] - offsets[i]));
            }
            return values;
        }

        private static MemoryStream BuildInferSchemaFixture()
        {
            return WorkbookBuilder.Build(
                """
                <row r="1">
                    <c r="A1" t="inlineStr"><is><t>Name</t></is></c>
                    <c r="B1" t="inlineStr"><is><t>Qty</t></is></c>
                    <c r="C1" t="inlineStr"><is><t>Price</t></is></c>
                    <c r="D1" t="inlineStr"><is><t>Active</t></is></c>
                    <c r="E1" t="inlineStr"><is><t>Mixed</t></is></c>
                </row>
                <row r="2">
                    <c r="A2" t="inlineStr"><is><t>Alice</t></is></c>
                    <c r="B2"><v>3</v></c>
                    <c r="C2"><v>1.5</v></c>
                    <c r="D2" t="b"><v>1</v></c>
                    <c r="E2" t="inlineStr"><is><t>oops</t></is></c>
                    <c r="F2"><v>10</v></c>
                </row>
                <row r="3">
                    <c r="A3" t="inlineStr"><is><t>Bob</t></is></c>
                    <c r="B3"><v>7</v></c>
                    <c r="C3"><v>4.5</v></c>
                    <c r="D3" t="b"><v>0</v></c>
                    <c r="E3"><v>2</v></c>
                </row>
                """);
        }

        private static void AssertSpec((string? Name, int Index, int Type, bool Nullable) spec, string? name, int type, bool nullable)
        {
            Assert.Equal(name, spec.Name);
            Assert.Equal(type, spec.Type);
            Assert.Equal(nullable, spec.Nullable);
        }

        private static NativeHandle NativeHandle_Create(IExcelRowReader reader)
        {
            return new NativeHandle(reader);
        }

        /// <summary>
        /// Wraps a real <see cref="IExcelRowReader"/>, forwarding everything except row enumeration:
        /// its enumerator yields <paramref name="failAfter"/> real rows and then throws, simulating a
        /// genuine mid-sheet decode failure for <see cref="ReadAllDecoded_Should_Free_Already_Decoded_Rows_When_A_Later_Row_Fails_To_Decode"/>.
        /// </summary>
        private sealed class FailAfterNRowsReader(IExcelRowReader inner, int failAfter) : IExcelRowReader
        {
            public bool IsDate1904 => inner.IsDate1904;
            public string SheetName => inner.SheetName;
            public int SheetCount => inner.SheetCount;
            public string SheetNameAt(int index)
            {
                return inner.SheetNameAt(index);
            }

            public ExcelSheetVisibility SheetVisibility => inner.SheetVisibility;

            public ExcelSheetVisibility SheetVisibilityAt(int index)
            {
                return inner.SheetVisibilityAt(index);
            }

            public bool TryMoveToSheet(ReadOnlySpan<char> name)
            {
                return inner.TryMoveToSheet(name);
            }

            public void MoveToSheet(int index)
            {
                inner.MoveToSheet(index);
            }

            public IExcelRowEnumerator GetEnumerator()
            {
                return new FailAfterNRowsEnumerator(inner.GetEnumerator(), failAfter);
            }

            public IExcelRowEnumerator GetAsyncEnumerator(CancellationToken ct = default)
            {
                return new FailAfterNRowsEnumerator(inner.GetAsyncEnumerator(ct), failAfter);
            }

            public void Dispose()
            {
                inner.Dispose();
            }

            public ValueTask DisposeAsync()
            {
                return inner.DisposeAsync();
            }
        }

        [Theory]
        [InlineData(NativeFormat.Auto)]
        [InlineData(NativeFormat.Xlsx)]
        public void OpenFile_Should_Open_Xlsx(int format)
        {
            int status = OpenPath(XlsxFixture, format, out NativeHandle? handle);

            Assert.Equal(NativeStatus.Ok, status);
            Assert.NotNull(handle);
            Assert.Equal(NativeStatus.Ok, ReadApi.Close(handle));
        }

        [Theory]
        [InlineData(NativeFormat.Auto)]
        [InlineData(NativeFormat.Xlsb)]
        public void OpenFile_Should_Open_Xlsb(int format)
        {
            int status = OpenPath(XlsbFixture, format, out NativeHandle? handle);

            Assert.Equal(NativeStatus.Ok, status);
            Assert.NotNull(handle);
            Assert.Equal(NativeStatus.Ok, ReadApi.Close(handle));
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
                Assert.Equal(NativeStatus.Ok, ReadApi.Close(handle));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void OpenMemory_Should_Open_Xlsx_From_A_Copy_Of_The_Bytes()
        {
            byte[] bytes = File.ReadAllBytes(XlsxFixture);

            int status = ReadApi.OpenMemory(bytes, NativeFormat.Auto, out NativeHandle? handle);

            Assert.Equal(NativeStatus.Ok, status);
            Assert.NotNull(handle);
            Assert.Equal(NativeStatus.Ok, ReadApi.Close(handle));
        }

        [Fact]
        public void OpenFileEx_With_Null_Options_Behaves_Like_OpenFile()
        {
            int status = ReadApi.OpenFileEx(Encoding.UTF8.GetBytes(XlsxFixture), NativeFormat.Auto, null, out NativeHandle? handle);

            Assert.Equal(NativeStatus.Ok, status);
            Assert.NotNull(handle);
            Assert.Equal(NativeStatus.Ok, ReadApi.Close(handle));
        }

        [Fact]
        public void OpenMemoryEx_With_Null_Options_Behaves_Like_OpenMemory()
        {
            byte[] bytes = File.ReadAllBytes(XlsxFixture);

            int status = ReadApi.OpenMemoryEx(bytes, NativeFormat.Auto, null, out NativeHandle? handle);

            Assert.Equal(NativeStatus.Ok, status);
            Assert.NotNull(handle);
            Assert.Equal(NativeStatus.Ok, ReadApi.Close(handle));
        }

        [Theory]
        [InlineData(-1)]
        [InlineData(256)]
        public void OpenFileEx_Rejects_An_Out_Of_Range_Csv_Delimiter(int delimiter)
        {
            NativeOpenOptionsRaw options = DefaultRawOptions() with { CsvDelimiter = delimiter };

            int status = ReadApi.OpenFileEx(Encoding.UTF8.GetBytes(XlsxFixture), NativeFormat.Csv, options, out NativeHandle? handle);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Null(handle);
        }

        [Fact]
        public void OpenFileEx_Rejects_A_Negative_Numeric_Option()
        {
            NativeOpenOptionsRaw options = DefaultRawOptions() with { MaxZipEntries = -5 };

            int status = ReadApi.OpenFileEx(Encoding.UTF8.GetBytes(XlsxFixture), NativeFormat.Auto, options, out NativeHandle? handle);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Null(handle);
        }

        [Fact]
        public void OpenFileEx_Rejects_An_Out_Of_Range_Csv_Sniff_Dialect_State()
        {
            NativeOpenOptionsRaw options = DefaultRawOptions() with { CsvSniffDialect = 99 };

            int status = ReadApi.OpenFileEx(Encoding.UTF8.GetBytes(XlsxFixture), NativeFormat.Csv, options, out NativeHandle? handle);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Null(handle);
        }

        [Fact]
        public void OpenFileEx_Applies_An_Explicit_Csv_Delimiter()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name;qty\nwidget;7\n");
            try
            {
                NativeOpenOptionsRaw options = DefaultRawOptions() with { CsvDelimiter = (byte)';' };
                Assert.Equal(NativeStatus.Ok, ReadApi.OpenFileEx(Encoding.UTF8.GetBytes(path), NativeFormat.Csv, options, out NativeHandle? handle));
                try
                {
                    byte[] buffer = new byte[4096];
                    Assert.Equal(NativeStatus.Ok, ReadApi.NextRow(handle, buffer, out int written));
                    List<DecodedCell> row = DecodeRow(buffer.AsSpan(0, written));
                    Assert.Equal(2, row.Count);
                    Assert.Equal("name", row[0].Value);
                    Assert.Equal("qty", row[1].Value);
                }
                finally
                {
                    ReadApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void OpenFileEx_Sniffs_The_Csv_Dialect_When_Requested()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name;qty\nwidget;7\ngadget;9\n");
            try
            {
                NativeOpenOptionsRaw options = DefaultRawOptions() with { CsvSniffDialect = NativeOptionState.True };
                Assert.Equal(NativeStatus.Ok, ReadApi.OpenFileEx(Encoding.UTF8.GetBytes(path), NativeFormat.Csv, options, out NativeHandle? handle));
                try
                {
                    byte[] buffer = new byte[4096];
                    Assert.Equal(NativeStatus.Ok, ReadApi.NextRow(handle, buffer, out int written));
                    List<DecodedCell> row = DecodeRow(buffer.AsSpan(0, written));
                    Assert.Equal(2, row.Count);
                    Assert.Equal("name", row[0].Value);
                    Assert.Equal("qty", row[1].Value);
                }
                finally
                {
                    ReadApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void OpenFileEx_Applies_A_Tiny_Max_Total_Decompressed_Bytes_To_A_Real_Xlsx()
        {
            NativeOpenOptionsRaw options = DefaultRawOptions() with { MaxTotalDecompressedBytes = 1 };

            int status = ReadApi.OpenFileEx(Encoding.UTF8.GetBytes(XlsxFixture), NativeFormat.Xlsx, options, out NativeHandle? handle);

            Assert.Equal(NativeStatus.Error, status);
            Assert.Null(handle);
        }

        [Fact]
        public void OpenFileEx_Applies_A_Tiny_Csv_Max_Cell_Bytes()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name\n" + new string('x', 100_000) + "\n");
            try
            {
                NativeOpenOptionsRaw options = DefaultRawOptions() with { CsvMaxCellBytes = 4 };
                Assert.Equal(NativeStatus.Ok, ReadApi.OpenFileEx(Encoding.UTF8.GetBytes(path), NativeFormat.Csv, options, out NativeHandle? handle));
                try
                {
                    Assert.Equal(NativeStatus.Ok, ReadApi.NextRow(handle, new byte[128 * 1024], out _));
                    Assert.Equal(NativeStatus.Error, ReadApi.NextRow(handle, new byte[128 * 1024], out _));
                }
                finally
                {
                    ReadApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void Should_Open_When_Password_Passed_Across_Abi()
        {
            byte[] pw = Encoding.UTF8.GetBytes(EncryptedFixtures.Password);
            IntPtr pointer = Marshal.AllocHGlobal(pw.Length);
            try
            {
                Marshal.Copy(pw, 0, pointer, pw.Length);
                NativeOpenOptionsRaw options = DefaultRawOptions() with { Password = pointer, PasswordLen = pw.Length };
                int status = ReadApi.OpenFileEx(
                    Encoding.UTF8.GetBytes(EncryptedFixtures.Path_("agile-aes256-sha512.xlsx")),
                    NativeFormat.Auto, options, out NativeHandle? handle);

                Assert.Equal(NativeStatus.Ok, status);
                Assert.NotNull(handle);
                ReadApi.Close(handle);
            }
            finally
            {
                Marshal.FreeHGlobal(pointer);
            }
        }

        [Fact]
        public void Should_Return_PasswordRequired_When_No_Password_Across_Abi()
        {
            NativeOpenOptionsRaw options = DefaultRawOptions();
            int status = ReadApi.OpenFileEx(
                Encoding.UTF8.GetBytes(EncryptedFixtures.Path_("agile-aes256-sha512.xlsx")),
                NativeFormat.Auto, options, out NativeHandle? handle);

            Assert.Equal(NativeStatus.PasswordRequired, status);
            Assert.Null(handle);
        }

        [Fact]
        public void Should_Return_PasswordIncorrect_When_Password_Wrong_Across_Abi()
        {
            byte[] pw = Encoding.UTF8.GetBytes("wrong");
            IntPtr pointer = Marshal.AllocHGlobal(pw.Length);
            try
            {
                Marshal.Copy(pw, 0, pointer, pw.Length);
                NativeOpenOptionsRaw options = DefaultRawOptions() with { Password = pointer, PasswordLen = pw.Length };
                int status = ReadApi.OpenFileEx(
                    Encoding.UTF8.GetBytes(EncryptedFixtures.Path_("agile-aes256-sha512.xlsx")),
                    NativeFormat.Auto, options, out NativeHandle? handle);

                Assert.Equal(NativeStatus.PasswordIncorrect, status);
                Assert.Null(handle);
            }
            finally
            {
                Marshal.FreeHGlobal(pointer);
            }
        }

        [Fact]
        public void Should_Reject_When_Password_Len_Is_Negative()
        {
            NativeOpenOptionsRaw options = DefaultRawOptions() with { Password = 1, PasswordLen = -1 };
            int status = ReadApi.OpenFileEx(
                Encoding.UTF8.GetBytes(EncryptedFixtures.Path_("agile-aes256-sha512.xlsx")),
                NativeFormat.Auto, options, out NativeHandle? handle);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Null(handle);
        }

        [Fact]
        public void Should_Reject_When_Struct_Size_Is_Stale()
        {
            NativeOpenOptionsRaw options = DefaultRawOptions() with { StructSize = Marshal.SizeOf<NativeOpenOptionsRaw>() - 8 };
            int status = ReadApi.OpenFileEx(
                Encoding.UTF8.GetBytes(EncryptedFixtures.Path_("agile-aes256-sha512.xlsx")),
                NativeFormat.Auto, options, out NativeHandle? handle);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Null(handle);
        }

        [Fact]
        public void Close_Should_Reject_A_Null_Handle()
        {
            Assert.Equal(NativeStatus.InvalidHandle, ReadApi.Close(null));
        }

        [Fact]
        public void SheetCount_Should_Report_At_Least_One_Sheet()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                int status = ReadApi.SheetCount(handle, out int count);

                Assert.Equal(NativeStatus.Ok, status);
                Assert.True(count >= 1);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void SheetName_Should_Return_A_Non_Empty_Utf8_Name()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                Span<byte> buffer = stackalloc byte[256];
                int status = ReadApi.SheetName(handle, buffer, out int length);

                Assert.Equal(NativeStatus.Ok, status);
                Assert.True(length > 0);
                Assert.NotEmpty(Encoding.UTF8.GetString(buffer[..length]));
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void SheetName_Should_Report_Required_Size_When_Buffer_Too_Small()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                int status = ReadApi.SheetName(handle, Span<byte>.Empty, out int length);

                Assert.Equal(NativeStatus.BufferTooSmall, status);
                Assert.True(length > 0);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void MoveToSheet_Should_Accept_The_First_Sheet()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, ReadApi.MoveToSheet(handle, 0));
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void MoveToSheet_Should_Fail_For_An_Out_Of_Range_Index()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Error, ReadApi.MoveToSheet(handle, 9999));
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void SheetNameAt_Should_Return_Each_Name_Without_Moving_The_Cursor()
        {
            using MemoryStream ms = WorkbookBuilder.BuildMultiSheet(
            [
                ("First", """<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>"""),
                ("Second", """<row r="1"><c r="A1"><v>10</v></c></row>"""),
            ]);
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, ReadApi.NextRow(handle, new byte[4096], out _));

                Span<byte> buffer = stackalloc byte[64];
                Assert.Equal(NativeStatus.Ok, ReadApi.SheetNameAt(handle, 0, buffer, out int firstLength));
                Assert.Equal("First", Encoding.UTF8.GetString(buffer[..firstLength]));
                Assert.Equal(NativeStatus.Ok, ReadApi.SheetNameAt(handle, 1, buffer, out int secondLength));
                Assert.Equal("Second", Encoding.UTF8.GetString(buffer[..secondLength]));

                Assert.Equal(NativeStatus.Ok, ReadApi.SheetName(handle, buffer, out int currentLength));
                Assert.Equal("First", Encoding.UTF8.GetString(buffer[..currentLength]));
                byte[] rowBuffer = new byte[4096];
                Assert.Equal(NativeStatus.Ok, ReadApi.NextRow(handle, rowBuffer, out int written));
                Assert.Equal("2", DecodeRow(rowBuffer.AsSpan(0, written))[0].Value);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void SheetNameAt_Should_Report_Required_Size_When_Buffer_Too_Small()
        {
            using MemoryStream ms = WorkbookBuilder.BuildMultiSheet([("VeryLongSheetName", "")]);
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                Span<byte> tiny = stackalloc byte[2];
                Assert.Equal(NativeStatus.BufferTooSmall, ReadApi.SheetNameAt(handle, 0, tiny, out int required));
                Assert.Equal("VeryLongSheetName".Length, required);

                Span<byte> big = stackalloc byte[required];
                Assert.Equal(NativeStatus.Ok, ReadApi.SheetNameAt(handle, 0, big, out int written));
                Assert.Equal("VeryLongSheetName", Encoding.UTF8.GetString(big[..written]));
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void SheetNameAt_Should_Reject_A_Negative_Index()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.InvalidArgument, ReadApi.SheetNameAt(handle, -1, stackalloc byte[64], out _));
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void SheetNameAt_Should_Error_On_An_Index_Past_The_Last_Sheet()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Error, ReadApi.SheetNameAt(handle, 9999, stackalloc byte[64], out _));
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void SheetNameAt_Should_Reject_A_Null_Handle()
        {
            Assert.Equal(NativeStatus.InvalidHandle, ReadApi.SheetNameAt(null, 0, stackalloc byte[64], out _));
        }

        [Fact]
        public void IsDate1904_Should_Report_Zero_For_A_1900_Based_Workbook()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                int status = ReadApi.IsDate1904(handle, out int flag);

                Assert.Equal(NativeStatus.Ok, status);
                Assert.Equal(0, flag);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void Sheet_Functions_Should_Reject_A_Null_Handle()
        {
            Assert.Equal(NativeStatus.InvalidHandle, ReadApi.SheetCount(null, out _));
            Assert.Equal(NativeStatus.InvalidHandle, ReadApi.MoveToSheet(null, 0));
            Assert.Equal(NativeStatus.InvalidHandle, ReadApi.IsDate1904(null, out _));
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

                Assert.Equal(NativeStatus.Ok, ReadApi.NextRow(handle, buffer, out int first));
                List<DecodedCell> header = DecodeRow(buffer.AsSpan(0, first));
                Assert.Equal(2, header.Count);
                Assert.Equal(0, header[0].Column);
                Assert.Equal("name", header[0].Value);
                Assert.Equal(1, header[1].Column);
                Assert.Equal("qty", header[1].Value);

                Assert.Equal(NativeStatus.Ok, ReadApi.NextRow(handle, buffer, out int second));
                List<DecodedCell> data = DecodeRow(buffer.AsSpan(0, second));
                Assert.Equal("widget", data[0].Value);
                Assert.Equal("7", data[1].Value);

                Assert.Equal(NativeStatus.Eof, ReadApi.NextRow(handle, buffer, out _));
            }
            finally
            {
                ReadApi.Close(handle);
                File.Delete(path);
            }
        }

        [Fact]
        public void NextRowDecoded_Should_Expose_Values_Through_A_C_Struct()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name,qty\nwidget,7\n");
            Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, ReadApi.NextRowDecoded(handle, out NativeRow row));
                try
                {
                    Assert.Equal(2, row.CellCount);
                    Assert.NotEqual(IntPtr.Zero, row.Cells);

                    int cellSize = Marshal.SizeOf<NativeRowCell>();
                    NativeRowCell[] cells =
                    [
                        Marshal.PtrToStructure<NativeRowCell>(row.Cells),
                        Marshal.PtrToStructure<NativeRowCell>(IntPtr.Add(row.Cells, cellSize)),
                    ];
                    Assert.Equal(0, cells[0].Column);
                    Assert.Equal(1, cells[1].Column);
                    Assert.Equal("name", Marshal.PtrToStringUTF8(cells[0].Value, cells[0].ValueLength));
                    Assert.Equal("qty", Marshal.PtrToStringUTF8(cells[1].Value, cells[1].ValueLength));
                }
                finally
                {
                    ReadApi.FreeRow(ref row);
                }

                Assert.Equal(NativeStatus.Ok, ReadApi.NextRow(handle, new byte[4096], out _));
                Assert.Equal(NativeStatus.Eof, ReadApi.NextRow(handle, new byte[4096], out _));
            }
            finally
            {
                ReadApi.Close(handle);
                File.Delete(path);
            }
        }

        [Fact]
        public void NextRowDecoded_Should_Resume_A_Row_Pending_From_The_Blob_API()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name\n");
            Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.BufferTooSmall, ReadApi.NextRow(handle, Span<byte>.Empty, out _));

                Assert.Equal(NativeStatus.Ok, ReadApi.NextRowDecoded(handle, out NativeRow row));
                try
                {
                    NativeRowCell cell = Marshal.PtrToStructure<NativeRowCell>(row.Cells);
                    Assert.Equal("name", Marshal.PtrToStringUTF8(cell.Value, cell.ValueLength));
                }
                finally
                {
                    ReadApi.FreeRow(ref row);
                }

                Assert.Equal(NativeStatus.Eof, ReadApi.NextRowDecoded(handle, out _));
            }
            finally
            {
                ReadApi.Close(handle);
                File.Delete(path);
            }
        }

        [Fact]
        public void NextRowDecoded_Should_Place_Every_Value_Inside_The_Row_Allocation()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name,qty,note\nwidget,7,fragile\n");
            Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, ReadApi.NextRow(handle, new byte[4096], out _));
                Assert.Equal(NativeStatus.Ok, ReadApi.NextRowDecoded(handle, out NativeRow row));
                try
                {
                    Assert.Equal(3, row.CellCount);
                    int cellSize = Marshal.SizeOf<NativeRowCell>();
                    NativeRowCell[] cells = new NativeRowCell[row.CellCount];
                    int totalValueBytes = 0;
                    for (int index = 0; index < row.CellCount; index++)
                    {
                        cells[index] = Marshal.PtrToStructure<NativeRowCell>(IntPtr.Add(row.Cells, index * cellSize));
                        totalValueBytes += cells[index].ValueLength + 1;
                    }

                    IntPtr valuesStart = IntPtr.Add(row.Cells, row.CellCount * cellSize);
                    IntPtr blockEnd = IntPtr.Add(valuesStart, totalValueBytes);
                    long previousEnd = valuesStart.ToInt64();
                    foreach (NativeRowCell cell in cells)
                    {
                        Assert.True(cell.Value.ToInt64() >= previousEnd, "value must not overlap the previous cell's value");
                        Assert.True(cell.Value.ToInt64() + cell.ValueLength < blockEnd.ToInt64(), "value must stay inside the row allocation");
                        previousEnd = cell.Value.ToInt64() + cell.ValueLength + 1;
                    }
                    Assert.Equal(blockEnd.ToInt64(), previousEnd);
                }
                finally
                {
                    ReadApi.FreeRow(ref row);
                }
            }
            finally
            {
                ReadApi.Close(handle);
                File.Delete(path);
            }
        }

        [Fact]
        public void NextRowDecoded_Should_NUL_Terminate_Every_Value()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name\nwidget\n");
            Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, ReadApi.NextRow(handle, new byte[4096], out _));
                Assert.Equal(NativeStatus.Ok, ReadApi.NextRowDecoded(handle, out NativeRow row));
                try
                {
                    NativeRowCell cell = Marshal.PtrToStructure<NativeRowCell>(row.Cells);
                    Assert.Equal((byte)0, Marshal.ReadByte(cell.Value, cell.ValueLength));
                }
                finally
                {
                    ReadApi.FreeRow(ref row);
                }
            }
            finally
            {
                ReadApi.Close(handle);
                File.Delete(path);
            }
        }

        [Fact]
        public void NextRowDecoded_Should_Return_A_Null_Cells_Pointer_For_An_Empty_Row()
        {
            using MemoryStream ms = WorkbookBuilder.Build("""<row r="1"/>""");
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, ReadApi.NextRowDecoded(handle, out NativeRow row));
                Assert.Equal(0, row.CellCount);
                Assert.Equal(IntPtr.Zero, row.Cells);

                ReadApi.FreeRow(ref row);
                Assert.Equal(IntPtr.Zero, row.Cells);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void FreeRow_Should_Be_Idempotent_On_A_Zeroed_Row()
        {
            NativeRow row = default;
            ReadApi.FreeRow(ref row);
            ReadApi.FreeRow(ref row);
            Assert.Equal(IntPtr.Zero, row.Cells);
            Assert.Equal(0, row.CellCount);
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
                Assert.Equal(NativeStatus.BufferTooSmall, ReadApi.NextRow(handle, tiny, out int required));
                Assert.True(required > 3);

                byte[] big = new byte[required];
                Assert.Equal(NativeStatus.Ok, ReadApi.NextRow(handle, big, out int written));
                Assert.Equal(required, written);
                Assert.Equal("name", DecodeRow(big.AsSpan(0, written))[0].Value);
            }
            finally
            {
                ReadApi.Close(handle);
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
                while (ReadApi.NextRow(handle, buffer, out int written) == NativeStatus.Ok)
                {
                    Assert.True(written >= 4);
                    rows++;
                }

                Assert.True(rows > 0);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void NextRow_Should_Restart_After_MoveToSheet()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                byte[] buffer = new byte[1 << 20];
                Assert.Equal(NativeStatus.Ok, ReadApi.NextRow(handle, buffer, out int firstPass));

                Assert.Equal(NativeStatus.Ok, ReadApi.MoveToSheet(handle, 0));

                Assert.Equal(NativeStatus.Ok, ReadApi.NextRow(handle, buffer, out int secondPass));
                Assert.Equal(firstPass, secondPass);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void NextRow_Should_Reject_A_Null_Handle()
        {
            Assert.Equal(NativeStatus.InvalidHandle, ReadApi.NextRow(null, new byte[16], out _));
            Assert.Equal(NativeStatus.InvalidHandle, ReadApi.NextRowDecoded(null, out _));
        }

        [Fact]
        public void NextRow_Should_Serialize_Xlsb_Numeric_Cells_With_A_Nonempty_Value()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsbFixture, NativeFormat.Xlsb, out NativeHandle? handle));
            try
            {
                byte[] buffer = new byte[1 << 20];
                bool foundNumericCell = false;

                while (ReadApi.NextRow(handle, buffer, out int written) == NativeStatus.Ok)
                {
                    List<DecodedCell> cells = DecodeRow(buffer.AsSpan(0, written));
                    foreach (DecodedCell cell in cells)
                    {
                        if ((cell.Type == 0 || cell.Type == 2) && !string.IsNullOrEmpty(cell.Value))
                        {
                            foundNumericCell = true;
                            break;
                        }
                    }
                    if (foundNumericCell)
                    {
                        break;
                    }
                }

                Assert.True(foundNumericCell, "XLSB fixture must contain at least one numeric cell with a non-empty serialized value");
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void ReadAllDecoded_Should_Return_Every_Remaining_Row_In_One_Call()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name,qty\nwidget,7\ngadget,9\n");
            Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, ReadApi.ReadAllDecoded(handle, out NativeRows rows));
                try
                {
                    Assert.Equal(3, rows.RowCount);
                    Assert.NotEqual(IntPtr.Zero, rows.Rows);

                    int rowSize = Marshal.SizeOf<NativeRow>();
                    NativeRow first = Marshal.PtrToStructure<NativeRow>(rows.Rows);
                    Assert.Equal(2, first.CellCount);
                    NativeRowCell firstCell = Marshal.PtrToStructure<NativeRowCell>(first.Cells);
                    Assert.Equal("name", Marshal.PtrToStringUTF8(firstCell.Value, firstCell.ValueLength));

                    NativeRow last = Marshal.PtrToStructure<NativeRow>(IntPtr.Add(rows.Rows, 2 * rowSize));
                    NativeRowCell lastCell = Marshal.PtrToStructure<NativeRowCell>(last.Cells);
                    Assert.Equal("gadget", Marshal.PtrToStringUTF8(lastCell.Value, lastCell.ValueLength));
                }
                finally
                {
                    ReadApi.FreeRows(ref rows);
                }
            }
            finally
            {
                ReadApi.Close(handle);
                File.Delete(path);
            }
        }

        [Fact]
        public void ReadAllDecoded_Should_Return_Zero_Rows_At_End_Of_Sheet()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name\n");
            Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, ReadApi.ReadAllDecoded(handle, out NativeRows first));
                ReadApi.FreeRows(ref first);

                Assert.Equal(NativeStatus.Ok, ReadApi.ReadAllDecoded(handle, out NativeRows second));
                Assert.Equal(0, second.RowCount);
                Assert.Equal(IntPtr.Zero, second.Rows);
                ReadApi.FreeRows(ref second);
            }
            finally
            {
                ReadApi.Close(handle);
                File.Delete(path);
            }
        }

        [Fact]
        public void ReadAllDecoded_Should_Return_InvalidHandle_For_A_Null_Handle()
        {
            Assert.Equal(NativeStatus.InvalidHandle, ReadApi.ReadAllDecoded(null, out _));
        }

        [Fact]
        public void ReadAllDecoded_Should_Free_Already_Decoded_Rows_When_A_Later_Row_Fails_To_Decode()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name,qty\nwidget,7\ngadget,9\ndoohickey,3\n");
            try
            {
                using (FileStream stream = File.OpenRead(path))
                {
                    var reader = new FailAfterNRowsReader(Excel.FromCsv(stream, leaveOpen: true), failAfter: 2);
                    NativeHandle handle = NativeHandle_Create(reader);
                    try
                    {
                        int status = ReadApi.ReadAllDecoded(handle, out NativeRows rows);

                        Assert.Equal(NativeStatus.Error, status);
                        Assert.Equal(0, rows.RowCount);
                        Assert.Equal(IntPtr.Zero, rows.Rows);

                        ReadApi.FreeRows(ref rows);
                    }
                    finally
                    {
                        ReadApi.Close(handle);
                    }
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ReadAllBlob_Should_Contain_Every_Remaining_Row()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name,qty\nwidget,7\ngadget,9\n");
            Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                byte[] buffer = new byte[8192];
                Assert.Equal(NativeStatus.Ok, ReadApi.ReadAllBlob(handle, buffer, out int written));

                List<List<DecodedCell>> rows = DecodeAllRowsBlob(buffer.AsSpan(0, written));
                Assert.Equal(3, rows.Count);
                Assert.Equal("name", rows[0][0].Value);
                Assert.Equal("qty", rows[0][1].Value);
                Assert.Equal("widget", rows[1][0].Value);
                Assert.Equal("gadget", rows[2][0].Value);
            }
            finally
            {
                ReadApi.Close(handle);
                File.Delete(path);
            }
        }

        [Fact]
        public void ReadAllBlob_Should_Include_A_Row_Already_Pending_From_NextRow()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name\nwidget\ngadget\n");
            Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.BufferTooSmall, ReadApi.NextRow(handle, Span<byte>.Empty, out _));

                byte[] buffer = new byte[8192];
                Assert.Equal(NativeStatus.Ok, ReadApi.ReadAllBlob(handle, buffer, out int written));

                List<List<DecodedCell>> rows = DecodeAllRowsBlob(buffer.AsSpan(0, written));
                Assert.Equal(3, rows.Count);
                Assert.Equal("name", rows[0][0].Value);
                Assert.Equal("widget", rows[1][0].Value);
                Assert.Equal("gadget", rows[2][0].Value);
            }
            finally
            {
                ReadApi.Close(handle);
                File.Delete(path);
            }
        }

        [Fact]
        public void ReadAllBlob_Should_Not_Lose_Rows_When_The_Buffer_Is_Too_Small()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name,qty\nwidget,7\ngadget,9\n");
            Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.BufferTooSmall, ReadApi.ReadAllBlob(handle, Span<byte>.Empty, out int required));
                Assert.True(required > 0);

                byte[] big = new byte[required];
                Assert.Equal(NativeStatus.Ok, ReadApi.ReadAllBlob(handle, big, out int written));
                Assert.Equal(required, written);

                List<List<DecodedCell>> rows = DecodeAllRowsBlob(big.AsSpan(0, written));
                Assert.Equal(3, rows.Count);
                Assert.Equal("gadget", rows[2][0].Value);
            }
            finally
            {
                ReadApi.Close(handle);
                File.Delete(path);
            }
        }

        [Fact]
        public void ReadAllBlob_Should_Return_Zero_Rows_At_End_Of_Sheet()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name\n");
            Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, ReadApi.ReadAllBlob(handle, new byte[4096], out _));

                byte[] buffer = new byte[4096];
                Assert.Equal(NativeStatus.Ok, ReadApi.ReadAllBlob(handle, buffer, out int written));
                Assert.Equal(0, BitConverter.ToInt32(buffer.AsSpan(0, written)));
            }
            finally
            {
                ReadApi.Close(handle);
                File.Delete(path);
            }
        }

        [Fact]
        public void ReadAllBlob_Should_Drop_Pending_Bytes_On_Sheet_Change()
        {
            using MemoryStream ms = WorkbookBuilder.BuildMultiSheet(
            [
                ("First", """<row r="1"><c r="A1"><v>1</v></c></row>"""),
                ("Second", """<row r="1"><c r="A1"><v>99</v></c></row>"""),
            ]);
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.BufferTooSmall, ReadApi.ReadAllBlob(handle, Span<byte>.Empty, out _));

                Assert.Equal(NativeStatus.Ok, ReadApi.MoveToSheet(handle, 1));

                byte[] buffer = new byte[4096];
                Assert.Equal(NativeStatus.Ok, ReadApi.ReadAllBlob(handle, buffer, out int written));
                List<List<DecodedCell>> rows = DecodeAllRowsBlob(buffer.AsSpan(0, written));
                Assert.Single(rows);
                Assert.Equal("99", rows[0][0].Value);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void ReadAllBlob_Should_Reject_A_Null_Handle()
        {
            Assert.Equal(NativeStatus.InvalidHandle, ReadApi.ReadAllBlob(null, new byte[64], out _));
        }

        [Fact]
        public void InferSchema_Should_Guess_Types_From_Sampled_Cells()
        {
            using MemoryStream ms = BuildInferSchemaFixture();
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, ReadApi.InferSchema(handle, headerRow: 1, sampleSize: 100, out NativeInferredSchema schema));
                try
                {
                    Assert.Equal(6, schema.ColumnCount);
                    (string? Name, int Index, int Type, bool Nullable)[] columns = DecodeSchema(schema);

                    AssertSpec(columns[0], "Name", NativeColumnType.String, nullable: false);
                    AssertSpec(columns[1], "Qty", NativeColumnType.Int64, nullable: false);
                    AssertSpec(columns[2], "Price", NativeColumnType.Float64, nullable: false);
                    AssertSpec(columns[3], "Active", NativeColumnType.Bool, nullable: false);
                    AssertSpec(columns[4], "Mixed", NativeColumnType.String, nullable: false);
                    AssertSpec(columns[5], name: null, NativeColumnType.Int64, nullable: true);
                    Assert.Equal(5, columns[5].Index);
                }
                finally
                {
                    ReadApi.FreeSchema(ref schema);
                }
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void InferSchema_Should_Resolve_By_Index_When_Header_Row_Is_Zero()
        {
            using MemoryStream ms = BuildInferSchemaFixture();
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, ReadApi.InferSchema(handle, headerRow: 0, sampleSize: 100, out NativeInferredSchema schema));
                try
                {
                    foreach ((string? name, _, _, _) in DecodeSchema(schema))
                    {
                        Assert.Null(name);
                    }
                }
                finally
                {
                    ReadApi.FreeSchema(ref schema);
                }
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void InferSchema_Should_Reject_A_Negative_Header_Row()
        {
            using MemoryStream ms = BuildInferSchemaFixture();
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.InvalidArgument, ReadApi.InferSchema(handle, headerRow: -1, sampleSize: 100, out NativeInferredSchema schema));
                Assert.Equal(IntPtr.Zero, schema.Columns);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Theory]
        [InlineData(0)]
        [InlineData(-1)]
        public void InferSchema_Should_Reject_A_NonPositive_Sample_Size(int sampleSize)
        {
            using MemoryStream ms = BuildInferSchemaFixture();
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.InvalidArgument, ReadApi.InferSchema(handle, headerRow: 1, sampleSize, out NativeInferredSchema schema));
                Assert.Equal(IntPtr.Zero, schema.Columns);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void InferSchema_Should_Reject_A_Null_Handle()
        {
            Assert.Equal(NativeStatus.InvalidHandle, ReadApi.InferSchema(null, headerRow: 1, sampleSize: 100, out _));
        }

        [Fact]
        public void InferSchema_Should_Report_An_Error_When_The_Sheet_Has_Fewer_Rows_Than_Header_Row()
        {
            using MemoryStream ms = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>1</v></c></row>""");
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                int status = ReadApi.InferSchema(handle, headerRow: 5, sampleSize: 100, out NativeInferredSchema schema);

                Assert.Equal(NativeStatus.InvalidArgument, status);
                Assert.Equal(IntPtr.Zero, schema.Columns);
                Span<byte> buffer = stackalloc byte[256];
                Assert.Equal(NativeStatus.Ok, NativeApi.LastError(buffer, out int length));
                Assert.Contains("fewer than", Encoding.UTF8.GetString(buffer[..length]), StringComparison.Ordinal);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void InferSchema_Should_Not_Disturb_The_Row_Cursor()
        {
            using MemoryStream ms = BuildInferSchemaFixture();
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, ReadApi.NextRow(handle, new byte[4096], out _));

                Assert.Equal(NativeStatus.Ok, ReadApi.InferSchema(handle, headerRow: 1, sampleSize: 100, out NativeInferredSchema schema));
                ReadApi.FreeSchema(ref schema);

                byte[] buffer = new byte[4096];
                Assert.Equal(NativeStatus.Ok, ReadApi.NextRow(handle, buffer, out int written));
                Assert.Equal("Alice", DecodeRow(buffer.AsSpan(0, written))[0].Value);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void FreeSchema_Should_Be_Idempotent_On_A_Zeroed_Schema()
        {
            NativeInferredSchema schema = default;
            ReadApi.FreeSchema(ref schema);
            ReadApi.FreeSchema(ref schema);
            Assert.Equal(IntPtr.Zero, schema.Columns);
        }
    }
}
