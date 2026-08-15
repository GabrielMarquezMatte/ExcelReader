using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using ExcelReader.Core.Writer;
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

        [Fact]
        public void LastErrorPtr_Should_Return_The_Stored_Message()
        {
            NativeApi.SetLastError("boom");

            nint pointer = NativeApi.LastErrorPtr(out int length);

            Assert.NotEqual(IntPtr.Zero, pointer);
            Assert.Equal(4, length);
            byte[] bytes = new byte[length];
            Marshal.Copy(pointer, bytes, 0, length);
            Assert.Equal("boom", Encoding.UTF8.GetString(bytes));
        }

        [Fact]
        public void LastErrorPtr_Should_Return_Zero_Length_When_Cleared()
        {
            NativeApi.SetLastError("boom");
            NativeApi.ClearLastError();

            nint pointer = NativeApi.LastErrorPtr(out int length);

            Assert.Equal(IntPtr.Zero, pointer);
            Assert.Equal(0, length);
        }

        [Fact]
        [SuppressMessage("Reliability", "S1215:GC.Collect should not be called",
            Justification = "This test's entire purpose is to prove the pointer survives a forced gen2 collection — that is the property under test, not incidental cleanup.")]
        public void LastErrorPtr_Should_Survive_A_Gen2_Collection()
        {
            // Catches an unpinned implementation: if the byte[] backing the pointer weren't allocated
            // pinned, a blocking gen2 collection could relocate it and the pointer taken before the
            // collection would now point at stale/reused memory.
            NativeApi.SetLastError("boom");
            nint pointer = NativeApi.LastErrorPtr(out int length);

            GC.Collect(2, GCCollectionMode.Forced, blocking: true);
            GC.WaitForPendingFinalizers();

            byte[] bytes = new byte[length];
            Marshal.Copy(pointer, bytes, 0, length);
            Assert.Equal("boom", Encoding.UTF8.GetString(bytes));
        }

        [Fact]
        public void LastErrorPtr_Should_Be_Thread_Local()
        {
            nint[] pointers = new nint[2];
            int[] lengths = new int[2];

            var first = new Thread(() =>
            {
                NativeApi.SetLastError("thread-one");
                pointers[0] = NativeApi.LastErrorPtr(out lengths[0]);
            });
            var second = new Thread(() =>
            {
                NativeApi.SetLastError("thread-two-error");
                pointers[1] = NativeApi.LastErrorPtr(out lengths[1]);
            });

            first.Start();
            first.Join();
            second.Start();
            second.Join();

            byte[] firstBytes = new byte[lengths[0]];
            Marshal.Copy(pointers[0], firstBytes, 0, lengths[0]);
            byte[] secondBytes = new byte[lengths[1]];
            Marshal.Copy(pointers[1], secondBytes, 0, lengths[1]);

            Assert.Equal("thread-one", Encoding.UTF8.GetString(firstBytes));
            Assert.Equal("thread-two-error", Encoding.UTF8.GetString(secondBytes));
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

        private static NativeOpenOptionsRaw DefaultRawOptions()
        {
            return new NativeOpenOptionsRaw { StructSize = Marshal.SizeOf<NativeOpenOptionsRaw>() };
        }

        [Fact]
        public void OpenFileEx_With_Null_Options_Behaves_Like_OpenFile()
        {
            int status = NativeApi.OpenFileEx(Encoding.UTF8.GetBytes(XlsxFixture), NativeFormat.Auto, null, out NativeHandle? handle);

            Assert.Equal(NativeStatus.Ok, status);
            Assert.NotNull(handle);
            Assert.Equal(NativeStatus.Ok, NativeApi.Close(handle));
        }

        [Fact]
        public void OpenMemoryEx_With_Null_Options_Behaves_Like_OpenMemory()
        {
            byte[] bytes = File.ReadAllBytes(XlsxFixture);

            int status = NativeApi.OpenMemoryEx(bytes, NativeFormat.Auto, null, out NativeHandle? handle);

            Assert.Equal(NativeStatus.Ok, status);
            Assert.NotNull(handle);
            Assert.Equal(NativeStatus.Ok, NativeApi.Close(handle));
        }

        [Fact]
        public void OpenFileEx_With_An_Unrecognized_Struct_Size_Is_Invalid_Argument()
        {
            NativeOpenOptionsRaw options = DefaultRawOptions() with { StructSize = 1 };

            int status = NativeApi.OpenFileEx(Encoding.UTF8.GetBytes(XlsxFixture), NativeFormat.Auto, options, out NativeHandle? handle);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Null(handle);
            Span<byte> buffer = stackalloc byte[256];
            Assert.Equal(NativeStatus.Ok, NativeApi.LastError(buffer, out int length));
            Assert.Contains("struct_size", Encoding.UTF8.GetString(buffer[..length]), StringComparison.Ordinal);
        }

        [Theory]
        [InlineData(-1)]
        [InlineData(256)]
        public void OpenFileEx_Rejects_An_Out_Of_Range_Csv_Delimiter(int delimiter)
        {
            NativeOpenOptionsRaw options = DefaultRawOptions() with { CsvDelimiter = delimiter };

            int status = NativeApi.OpenFileEx(Encoding.UTF8.GetBytes(XlsxFixture), NativeFormat.Csv, options, out NativeHandle? handle);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Null(handle);
        }

        [Fact]
        public void OpenFileEx_Rejects_A_Negative_Numeric_Option()
        {
            NativeOpenOptionsRaw options = DefaultRawOptions() with { MaxZipEntries = -5 };

            int status = NativeApi.OpenFileEx(Encoding.UTF8.GetBytes(XlsxFixture), NativeFormat.Auto, options, out NativeHandle? handle);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Null(handle);
        }

        [Fact]
        public void OpenFileEx_Rejects_An_Out_Of_Range_Csv_Sniff_Dialect_State()
        {
            NativeOpenOptionsRaw options = DefaultRawOptions() with { CsvSniffDialect = 99 };

            int status = NativeApi.OpenFileEx(Encoding.UTF8.GetBytes(XlsxFixture), NativeFormat.Csv, options, out NativeHandle? handle);

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
                Assert.Equal(NativeStatus.Ok, NativeApi.OpenFileEx(Encoding.UTF8.GetBytes(path), NativeFormat.Csv, options, out NativeHandle? handle));
                try
                {
                    byte[] buffer = new byte[4096];
                    Assert.Equal(NativeStatus.Ok, NativeApi.NextRow(handle, buffer, out int written));
                    List<DecodedCell> row = DecodeRow(buffer.AsSpan(0, written));
                    Assert.Equal(2, row.Count);
                    Assert.Equal("name", row[0].Value);
                    Assert.Equal("qty", row[1].Value);
                }
                finally
                {
                    NativeApi.Close(handle);
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
            // Same semicolon-delimited file as the explicit-delimiter test above, but with no delimiter
            // given at all — csv_sniff_dialect must infer it via Excel.SniffCsvDialectFromFile.
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name;qty\nwidget;7\ngadget;9\n");
            try
            {
                NativeOpenOptionsRaw options = DefaultRawOptions() with { CsvSniffDialect = NativeOptionState.True };
                Assert.Equal(NativeStatus.Ok, NativeApi.OpenFileEx(Encoding.UTF8.GetBytes(path), NativeFormat.Csv, options, out NativeHandle? handle));
                try
                {
                    byte[] buffer = new byte[4096];
                    Assert.Equal(NativeStatus.Ok, NativeApi.NextRow(handle, buffer, out int written));
                    List<DecodedCell> row = DecodeRow(buffer.AsSpan(0, written));
                    Assert.Equal(2, row.Count);
                    Assert.Equal("name", row[0].Value);
                    Assert.Equal("qty", row[1].Value);
                }
                finally
                {
                    NativeApi.Close(handle);
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
            // Proves max_total_decompressed_bytes actually reaches XlsxReader: a cap this small fails
            // before even xl/workbook.xml can be read.
            NativeOpenOptionsRaw options = DefaultRawOptions() with { MaxTotalDecompressedBytes = 1 };

            int status = NativeApi.OpenFileEx(Encoding.UTF8.GetBytes(XlsxFixture), NativeFormat.Xlsx, options, out NativeHandle? handle);

            Assert.Equal(NativeStatus.Error, status);
            Assert.Null(handle);
        }

        [Fact]
        public void OpenFileEx_Applies_A_Tiny_Csv_Max_Cell_Bytes()
        {
            // Proves csv_max_cell_bytes actually reaches CsvReader. A plain (unquoted) field is read
            // zero-copy straight out of the stream's own read buffer (CsvReader.Enumerator's "simple
            // record" fast path) without ever touching CellAccumulator's separate value buffer — so the
            // cap only has anything to enforce once the record forces BufferedStreamCursor's raw read
            // buffer to grow past its 64 KiB initial size. The field below (100,000 bytes) guarantees
            // that grow happens; the cap (4) guarantees it throws when it does.
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name\n" + new string('x', 100_000) + "\n");
            try
            {
                NativeOpenOptionsRaw options = DefaultRawOptions() with { CsvMaxCellBytes = 4 };
                Assert.Equal(NativeStatus.Ok, NativeApi.OpenFileEx(Encoding.UTF8.GetBytes(path), NativeFormat.Csv, options, out NativeHandle? handle));
                try
                {
                    Assert.Equal(NativeStatus.Ok, NativeApi.NextRow(handle, new byte[128 * 1024], out _)); // header row, short enough
                    Assert.Equal(NativeStatus.Error, NativeApi.NextRow(handle, new byte[128 * 1024], out _)); // data row, forces a grow past the cap
                }
                finally
                {
                    NativeApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
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
        public void SheetNameAt_Should_Return_Each_Name_Without_Moving_The_Cursor()
        {
            using MemoryStream ms = WorkbookBuilder.BuildMultiSheet(
            [
                ("First", """<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>"""),
                ("Second", """<row r="1"><c r="A1"><v>10</v></c></row>"""),
            ]);
            Assert.Equal(NativeStatus.Ok, NativeApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                // Read the first row of sheet 0 before listing names, then confirm the second row of the
                // SAME sheet still comes back afterward — SheetNameAt must not move the cursor or the
                // current sheet, unlike xl_move_to_sheet.
                Assert.Equal(NativeStatus.Ok, NativeApi.NextRow(handle, new byte[4096], out _));

                Span<byte> buffer = stackalloc byte[64];
                Assert.Equal(NativeStatus.Ok, NativeApi.SheetNameAt(handle, 0, buffer, out int firstLength));
                Assert.Equal("First", Encoding.UTF8.GetString(buffer[..firstLength]));
                Assert.Equal(NativeStatus.Ok, NativeApi.SheetNameAt(handle, 1, buffer, out int secondLength));
                Assert.Equal("Second", Encoding.UTF8.GetString(buffer[..secondLength]));

                // The current sheet is still 0, and its second row is still next — SheetNameAt must not
                // have reset row enumeration the way xl_move_to_sheet does.
                Assert.Equal(NativeStatus.Ok, NativeApi.SheetName(handle, buffer, out int currentLength));
                Assert.Equal("First", Encoding.UTF8.GetString(buffer[..currentLength]));
                byte[] rowBuffer = new byte[4096];
                Assert.Equal(NativeStatus.Ok, NativeApi.NextRow(handle, rowBuffer, out int written));
                Assert.Equal("2", DecodeRow(rowBuffer.AsSpan(0, written))[0].Value);
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void SheetNameAt_Should_Report_Required_Size_When_Buffer_Too_Small()
        {
            using MemoryStream ms = WorkbookBuilder.BuildMultiSheet([("VeryLongSheetName", "")]);
            Assert.Equal(NativeStatus.Ok, NativeApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                Span<byte> tiny = stackalloc byte[2];
                Assert.Equal(NativeStatus.BufferTooSmall, NativeApi.SheetNameAt(handle, 0, tiny, out int required));
                Assert.Equal("VeryLongSheetName".Length, required);

                Span<byte> big = stackalloc byte[required];
                Assert.Equal(NativeStatus.Ok, NativeApi.SheetNameAt(handle, 0, big, out int written));
                Assert.Equal("VeryLongSheetName", Encoding.UTF8.GetString(big[..written]));
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void SheetNameAt_Should_Reject_A_Negative_Index()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.InvalidArgument, NativeApi.SheetNameAt(handle, -1, stackalloc byte[64], out _));
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void SheetNameAt_Should_Error_On_An_Index_Past_The_Last_Sheet()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Error, NativeApi.SheetNameAt(handle, 9999, stackalloc byte[64], out _));
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void SheetNameAt_Should_Reject_A_Null_Handle()
        {
            Assert.Equal(NativeStatus.InvalidHandle, NativeApi.SheetNameAt(null, 0, stackalloc byte[64], out _));
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
        public void NextRowDecoded_Should_Expose_Values_Through_A_C_Struct()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name,qty\nwidget,7\n");
            Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, NativeApi.NextRowDecoded(handle, out NativeRow row));
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
                    NativeApi.FreeRow(ref row);
                }

                Assert.Equal(NativeStatus.Ok, NativeApi.NextRow(handle, new byte[4096], out _));
                Assert.Equal(NativeStatus.Eof, NativeApi.NextRow(handle, new byte[4096], out _));
            }
            finally
            {
                NativeApi.Close(handle);
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
                Assert.Equal(NativeStatus.BufferTooSmall, NativeApi.NextRow(handle, Span<byte>.Empty, out _));

                Assert.Equal(NativeStatus.Ok, NativeApi.NextRowDecoded(handle, out NativeRow row));
                try
                {
                    NativeRowCell cell = Marshal.PtrToStructure<NativeRowCell>(row.Cells);
                    Assert.Equal("name", Marshal.PtrToStringUTF8(cell.Value, cell.ValueLength));
                }
                finally
                {
                    NativeApi.FreeRow(ref row);
                }

                Assert.Equal(NativeStatus.Eof, NativeApi.NextRowDecoded(handle, out _));
            }
            finally
            {
                NativeApi.Close(handle);
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
                // Skip the header row; assert on the data row, which has three non-empty cells.
                Assert.Equal(NativeStatus.Ok, NativeApi.NextRow(handle, new byte[4096], out _));
                Assert.Equal(NativeStatus.Ok, NativeApi.NextRowDecoded(handle, out NativeRow row));
                try
                {
                    Assert.Equal(3, row.CellCount);
                    int cellSize = Marshal.SizeOf<NativeRowCell>();
                    NativeRowCell[] cells = new NativeRowCell[row.CellCount];
                    int totalValueBytes = 0;
                    for (int index = 0; index < row.CellCount; index++)
                    {
                        cells[index] = Marshal.PtrToStructure<NativeRowCell>(IntPtr.Add(row.Cells, index * cellSize));
                        totalValueBytes += cells[index].ValueLength + 1; // +1 for the NUL terminator.
                    }

                    // Every value pointer must land inside the single row allocation: at or after where the
                    // cell array ends, and before the block's end (cell array + every value + its NUL).
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
                    NativeApi.FreeRow(ref row);
                }
            }
            finally
            {
                NativeApi.Close(handle);
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
                Assert.Equal(NativeStatus.Ok, NativeApi.NextRow(handle, new byte[4096], out _));
                Assert.Equal(NativeStatus.Ok, NativeApi.NextRowDecoded(handle, out NativeRow row));
                try
                {
                    NativeRowCell cell = Marshal.PtrToStructure<NativeRowCell>(row.Cells);
                    Assert.Equal((byte)0, Marshal.ReadByte(cell.Value, cell.ValueLength));
                }
                finally
                {
                    NativeApi.FreeRow(ref row);
                }
            }
            finally
            {
                NativeApi.Close(handle);
                File.Delete(path);
            }
        }

        [Fact]
        public void NextRowDecoded_Should_Return_A_Null_Cells_Pointer_For_An_Empty_Row()
        {
            // Self-closing <row/> is a raw-XML shape with zero cells (see CellVariantTests.SelfClosingRowYieldsZeroColumnCount).
            // A blank CSV line does not test this: it yields one empty field, not zero cells (see
            // CsvReaderTests.BlankLineYieldsOneEmptyField).
            using MemoryStream ms = WorkbookBuilder.Build("""<row r="1"/>""");
            Assert.Equal(NativeStatus.Ok, NativeApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, NativeApi.NextRowDecoded(handle, out NativeRow row));
                Assert.Equal(0, row.CellCount);
                Assert.Equal(IntPtr.Zero, row.Cells);

                // FreeRow on an already-empty row must be a harmless no-op, not a crash.
                NativeApi.FreeRow(ref row);
                Assert.Equal(IntPtr.Zero, row.Cells);
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void FreeRow_Should_Be_Idempotent_On_A_Zeroed_Row()
        {
            NativeRow row = default;
            NativeApi.FreeRow(ref row);
            NativeApi.FreeRow(ref row);
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
            Assert.Equal(NativeStatus.InvalidHandle, NativeApi.NextRowDecoded(null, out _));
        }

        [Fact]
        public void NextRow_Should_Serialize_Xlsb_Numeric_Cells_With_A_Nonempty_Value()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsbFixture, NativeFormat.Xlsb, out NativeHandle? handle));
            try
            {
                byte[] buffer = new byte[1 << 20];
                bool foundNumericCell = false;

                // Read rows until we find a Number or Date typed cell with a non-empty value
                while (NativeApi.NextRow(handle, buffer, out int written) == NativeStatus.Ok)
                {
                    List<DecodedCell> cells = DecodeRow(buffer.AsSpan(0, written));
#pragma warning disable HLQ012
                    foreach (DecodedCell cell in cells)
                    {
                        // CellType.Number = 0, CellType.Date = 2 (from the XLSX/XLSB reader)
                        if ((cell.Type == 0 || cell.Type == 2) && !string.IsNullOrEmpty(cell.Value))
                        {
                            foundNumericCell = true;
                            break;
                        }
                    }
#pragma warning restore HLQ012
                    if (foundNumericCell)
                    {
                        break;
                    }
                }

                // RealExcel.xlsb is known to contain numeric cells; we must find at least one
                Assert.True(foundNumericCell, "XLSB fixture must contain at least one numeric cell with a non-empty serialized value");
            }
            finally
            {
                NativeApi.Close(handle);
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
                Assert.Equal(NativeStatus.Ok, NativeApi.ReadAllDecoded(handle, out NativeRows rows));
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
                    NativeApi.FreeRows(ref rows);
                }
            }
            finally
            {
                NativeApi.Close(handle);
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
                Assert.Equal(NativeStatus.Ok, NativeApi.ReadAllDecoded(handle, out NativeRows first));
                NativeApi.FreeRows(ref first);

                Assert.Equal(NativeStatus.Ok, NativeApi.ReadAllDecoded(handle, out NativeRows second));
                Assert.Equal(0, second.RowCount);
                Assert.Equal(IntPtr.Zero, second.Rows);
                NativeApi.FreeRows(ref second); // must be a no-op on a zeroed value, not throw
            }
            finally
            {
                NativeApi.Close(handle);
                File.Delete(path);
            }
        }

        [Fact]
        public void ReadAllDecoded_Should_Return_InvalidHandle_For_A_Null_Handle()
        {
            Assert.Equal(NativeStatus.InvalidHandle, NativeApi.ReadAllDecoded(null, out _));
        }

        [Fact]
        public void ReadAllDecoded_Should_Free_Already_Decoded_Rows_When_A_Later_Row_Fails_To_Decode()
        {
            // Regression test for a leak: a mid-loop decode error used to `return status;` straight out
            // of ReadAllDecoded's try block, skipping the catch block that frees every row already
            // decoded. This reader succeeds for the first two rows, then throws on the third row's
            // MoveNext(), reproducing that exact "some rows decoded, then a real failure" shape without
            // needing a genuinely malformed file. NextRow's own try/catch turns that thrown exception
            // into a plain NativeStatus.Error return, so this exercises the non-exceptional mid-loop
            // error path in ReadAllDecoded, not its outer catch block.
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name,qty\nwidget,7\ngadget,9\ndoohickey,3\n");
            try
            {
                using (FileStream stream = File.OpenRead(path))
                {
                    // leaveOpen: true — the enclosing `using FileStream` owns the stream; NativeHandle.Dispose
                    // (via NativeApi.Close below) owns disposing the FailAfterNRowsReader/CsvReader chain.
                    var reader = new FailAfterNRowsReader(Excel.FromCsv(stream, leaveOpen: true), failAfter: 2);
                    NativeHandle handle = NativeHandle_Create(reader);
                    try
                    {
                        int status = NativeApi.ReadAllDecoded(handle, out NativeRows rows);

                        Assert.Equal(NativeStatus.Error, status);
                        Assert.Equal(0, rows.RowCount);
                        Assert.Equal(IntPtr.Zero, rows.Rows);

                        NativeApi.FreeRows(ref rows); // must still be safe/no-op on the zeroed result
                    }
                    finally
                    {
                        NativeApi.Close(handle);
                    }
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        // Decodes the xl_read_all_blob layout: int32 row_count, then row_count * {int32 row_length, row blob}.
        // Reuses DecodeRow (the single-row blob decoder already used above) for each entry.
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

        [Fact]
        public void ReadAllBlob_Should_Contain_Every_Remaining_Row()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name,qty\nwidget,7\ngadget,9\n");
            Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                byte[] buffer = new byte[8192];
                Assert.Equal(NativeStatus.Ok, NativeApi.ReadAllBlob(handle, buffer, out int written));

                List<List<DecodedCell>> rows = DecodeAllRowsBlob(buffer.AsSpan(0, written));
                Assert.Equal(3, rows.Count);
                Assert.Equal("name", rows[0][0].Value);
                Assert.Equal("qty", rows[0][1].Value);
                Assert.Equal("widget", rows[1][0].Value);
                Assert.Equal("gadget", rows[2][0].Value);
            }
            finally
            {
                NativeApi.Close(handle);
                File.Delete(path);
            }
        }

        [Fact]
        public void ReadAllBlob_Should_Include_A_Row_Already_Pending_From_NextRow()
        {
            // A row held pending from a prior xl_next_row that returned XL_BUFFER_TOO_SMALL must still
            // show up in xl_read_all_blob's result — it hasn't been consumed by any successful call yet.
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name\nwidget\ngadget\n");
            Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.BufferTooSmall, NativeApi.NextRow(handle, Span<byte>.Empty, out _));

                byte[] buffer = new byte[8192];
                Assert.Equal(NativeStatus.Ok, NativeApi.ReadAllBlob(handle, buffer, out int written));

                List<List<DecodedCell>> rows = DecodeAllRowsBlob(buffer.AsSpan(0, written));
                Assert.Equal(3, rows.Count);
                Assert.Equal("name", rows[0][0].Value);
                Assert.Equal("widget", rows[1][0].Value);
                Assert.Equal("gadget", rows[2][0].Value);
            }
            finally
            {
                NativeApi.Close(handle);
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
                Assert.Equal(NativeStatus.BufferTooSmall, NativeApi.ReadAllBlob(handle, Span<byte>.Empty, out int required));
                Assert.True(required > 0);

                byte[] big = new byte[required];
                Assert.Equal(NativeStatus.Ok, NativeApi.ReadAllBlob(handle, big, out int written));
                Assert.Equal(required, written);

                List<List<DecodedCell>> rows = DecodeAllRowsBlob(big.AsSpan(0, written));
                Assert.Equal(3, rows.Count);
                Assert.Equal("gadget", rows[2][0].Value);
            }
            finally
            {
                NativeApi.Close(handle);
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
                Assert.Equal(NativeStatus.Ok, NativeApi.ReadAllBlob(handle, new byte[4096], out _));

                byte[] buffer = new byte[4096];
                Assert.Equal(NativeStatus.Ok, NativeApi.ReadAllBlob(handle, buffer, out int written));
                Assert.Equal(0, BitConverter.ToInt32(buffer.AsSpan(0, written)));
            }
            finally
            {
                NativeApi.Close(handle);
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
            Assert.Equal(NativeStatus.Ok, NativeApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                // Force a too-small result so accumulated bytes for sheet 0 are held pending.
                Assert.Equal(NativeStatus.BufferTooSmall, NativeApi.ReadAllBlob(handle, Span<byte>.Empty, out _));

                Assert.Equal(NativeStatus.Ok, NativeApi.MoveToSheet(handle, 1));

                byte[] buffer = new byte[4096];
                Assert.Equal(NativeStatus.Ok, NativeApi.ReadAllBlob(handle, buffer, out int written));
                List<List<DecodedCell>> rows = DecodeAllRowsBlob(buffer.AsSpan(0, written));
                Assert.Single(rows);
                Assert.Equal("99", rows[0][0].Value);
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void ReadAllBlob_Should_Reject_A_Null_Handle()
        {
            Assert.Equal(NativeStatus.InvalidHandle, NativeApi.ReadAllBlob(null, new byte[64], out _));
        }

        // Builds a native table whose buffers live in unmanaged memory, so write-path tests exercise
        // the same pointer walk the ABI does. Every allocation is released by FreeBuiltTable.
        private static NativeTable BuildInt64Table(long[] values, byte[]? validity = null)
        {
            NativeColumn column = new()
            {
                Type = NativeColumnType.Int64,
                Length = values.LongLength,
                Values = Marshal.AllocHGlobal(values.Length * sizeof(long)),
                Validity = IntPtr.Zero,
                Data = IntPtr.Zero,
                DataLen = 0,
            };
            Marshal.Copy(values, 0, column.Values, values.Length);
            if (validity is not null)
            {
                column.Validity = Marshal.AllocHGlobal(validity.Length);
                Marshal.Copy(validity, 0, column.Validity, validity.Length);
            }

            IntPtr columns = Marshal.AllocHGlobal(Marshal.SizeOf<NativeColumn>());
            Marshal.StructureToPtr(column, columns, false);
            return new NativeTable { ColumnCount = 1, RowCount = values.LongLength, Columns = columns };
        }

        // Mirrors BuildInt64Table for XL_T_STRING: `offsets` and `data` are two INDEPENDENT
        // allocations here, which the write direction explicitly permits (unlike ParseTyped's output,
        // where Data is interior to Values).
        private static NativeTable BuildStringTable(int[] offsets, byte[] data)
        {
            NativeColumn column = new()
            {
                Type = NativeColumnType.String,
                Length = offsets.Length - 1,
                Values = Marshal.AllocHGlobal(offsets.Length * sizeof(int)),
                Validity = IntPtr.Zero,
                Data = data.Length == 0 ? IntPtr.Zero : Marshal.AllocHGlobal(data.Length),
                DataLen = data.Length,
            };
            Marshal.Copy(offsets, 0, column.Values, offsets.Length);
            if (data.Length > 0)
            {
                Marshal.Copy(data, 0, column.Data, data.Length);
            }

            IntPtr columns = Marshal.AllocHGlobal(Marshal.SizeOf<NativeColumn>());
            Marshal.StructureToPtr(column, columns, false);
            return new NativeTable { ColumnCount = 1, RowCount = offsets.Length - 1, Columns = columns };
        }

        private static void FreeBuiltTable(ref NativeTable table)
        {
            // Deliberately NOT NativeApi.FreeTable: that one knows Data is interior to Values, which is
            // true of ParseTyped's output but not of the independently-allocated tables built above.
            for (int index = 0; index < table.ColumnCount; index++)
            {
                NativeColumn column = Marshal.PtrToStructure<NativeColumn>(
                    IntPtr.Add(table.Columns, index * Marshal.SizeOf<NativeColumn>()));
                foreach (IntPtr block in new[] { column.Values, column.Validity, column.Data })
                {
                    if (block != IntPtr.Zero)
                    {
                        Marshal.FreeHGlobal(block);
                    }
                }
            }
            Marshal.FreeHGlobal(table.Columns);
            table = default;
        }

        [Fact]
        public void ValidateWriteTable_Should_Accept_A_Well_Formed_Named_Table()
        {
            NativeTable table = BuildInt64Table([1L, 2L]);
            try
            {
                NativeColumnSpec[] specs = [new() { Name = "qty", Type = NativeColumnType.Int64 }];

                Assert.True(NativeApi.TryValidateWriteTable(specs, table, out bool hasHeader, out string? error));
                Assert.True(hasHeader);
                Assert.Null(error);
            }
            finally
            {
                FreeBuiltTable(ref table);
            }
        }

        [Fact]
        public void ValidateWriteTable_Should_Report_No_Header_When_Every_Spec_Is_Unnamed()
        {
            NativeTable table = BuildInt64Table([1L]);
            try
            {
                NativeColumnSpec[] specs = [new() { Name = null, Type = NativeColumnType.Int64 }];

                Assert.True(NativeApi.TryValidateWriteTable(specs, table, out bool hasHeader, out _));
                Assert.False(hasHeader);
            }
            finally
            {
                FreeBuiltTable(ref table);
            }
        }

        [Fact]
        public void ValidateWriteTable_Should_Reject_A_Mix_Of_Named_And_Unnamed_Specs()
        {
            NativeTable table = BuildInt64Table([1L]);
            table.ColumnCount = 2;
            try
            {
                NativeColumnSpec[] specs =
                [
                    new() { Name = "qty", Type = NativeColumnType.Int64 },
                    new() { Name = null, Type = NativeColumnType.Int64 },
                ];

                Assert.False(NativeApi.TryValidateWriteTable(specs, table, out _, out string? error));
                Assert.Contains("name", error, StringComparison.Ordinal);
            }
            finally
            {
                table.ColumnCount = 1;
                FreeBuiltTable(ref table);
            }
        }

        [Fact]
        public void ValidateWriteTable_Should_Reject_A_Spec_Type_That_Disagrees_With_Its_Column()
        {
            NativeTable table = BuildInt64Table([1L]);
            try
            {
                NativeColumnSpec[] specs = [new() { Name = "qty", Type = NativeColumnType.Float64 }];

                Assert.False(NativeApi.TryValidateWriteTable(specs, table, out _, out string? error));
                Assert.Contains("type", error, StringComparison.Ordinal);
            }
            finally
            {
                FreeBuiltTable(ref table);
            }
        }

        [Fact]
        public void ValidateWriteTable_Should_Reject_A_Column_Whose_Length_Is_Not_The_Row_Count()
        {
            NativeTable table = BuildInt64Table([1L, 2L]);
            table.RowCount = 3;
            try
            {
                NativeColumnSpec[] specs = [new() { Name = "qty", Type = NativeColumnType.Int64 }];

                Assert.False(NativeApi.TryValidateWriteTable(specs, table, out _, out string? error));
                Assert.Contains("length", error, StringComparison.Ordinal);
            }
            finally
            {
                table.RowCount = 2;
                FreeBuiltTable(ref table);
            }
        }

        [Fact]
        public void ValidateWriteTable_Should_Accept_Well_Formed_String_Offsets()
        {
            NativeTable table = BuildStringTable([0, 6, 12], "widgetgadget"u8.ToArray());
            try
            {
                NativeColumnSpec[] specs = [new() { Name = "name", Type = NativeColumnType.String }];

                Assert.True(NativeApi.TryValidateWriteTable(specs, table, out _, out string? error));
                Assert.Null(error);
            }
            finally
            {
                FreeBuiltTable(ref table);
            }
        }

        [Theory]
        [InlineData(new[] { 1, 6, 12 })]                 // does not start at 0
        [InlineData(new[] { 0, -1, 12 })]                // negative offset
        [InlineData(new[] { 0, 9, 6 })]                  // not monotonic
        [InlineData(new[] { 0, 6, 13 })]                 // last offset past data_len
        [InlineData(new[] { 0, 6, 11 })]                 // last offset short of data_len
        public void ValidateWriteTable_Should_Reject_Malformed_String_Offsets(int[] offsets)
        {
            NativeTable table = BuildStringTable(offsets, "widgetgadget"u8.ToArray());
            try
            {
                NativeColumnSpec[] specs = [new() { Name = "name", Type = NativeColumnType.String }];

                Assert.False(NativeApi.TryValidateWriteTable(specs, table, out _, out string? error));
                Assert.Contains("offset", error, StringComparison.Ordinal);
            }
            finally
            {
                FreeBuiltTable(ref table);
            }
        }

        [Fact]
        public void ValidateWriteTable_Should_Reject_A_Null_Values_Pointer_With_Rows()
        {
            NativeTable table = BuildInt64Table([1L]);
            try
            {
                IntPtr original = Marshal.ReadIntPtr(table.Columns, sizeof(long) * 2);
                NativeColumn column = Marshal.PtrToStructure<NativeColumn>(table.Columns);
                column.Values = IntPtr.Zero;
                Marshal.StructureToPtr(column, table.Columns, false);
                NativeColumnSpec[] specs = [new() { Name = "qty", Type = NativeColumnType.Int64 }];

                Assert.False(NativeApi.TryValidateWriteTable(specs, table, out _, out string? error));
                Assert.Contains("values", error, StringComparison.Ordinal);

                column.Values = original;
                Marshal.StructureToPtr(column, table.Columns, false);
            }
            finally
            {
                FreeBuiltTable(ref table);
            }
        }

        private static NativeColumn ColumnAt(NativeTable table, int index)
        {
            int columnSize = Marshal.SizeOf<NativeColumn>();
            return Marshal.PtrToStructure<NativeColumn>(IntPtr.Add(table.Columns, index * columnSize));
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

        private static bool[] DecodeValidity(NativeColumn column)
        {
            int rowCount = (int)column.Length;
            bool[] result = new bool[rowCount];
            if (column.Validity == IntPtr.Zero)
            {
                Array.Fill(result, true); // NULL validity means every value is valid
                return result;
            }
            byte[] bitmap = new byte[(rowCount + 7) / 8];
            Marshal.Copy(column.Validity, bitmap, 0, bitmap.Length);
            for (int i = 0; i < rowCount; i++)
            {
                result[i] = (bitmap[i >> 3] & (1 << (i & 7))) != 0;
            }
            return result;
        }

        [Fact]
        public void ParseTyped_Should_Return_Typed_Columns_By_Name()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name,qty,price,active,joined\nwidget,3,9.99,true,2024-01-15\ngadget,7,4.5,false,2024-02-20\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs =
                [
                    new() { Name = "name", Type = NativeColumnType.String },
                    new() { Name = "qty", Type = NativeColumnType.Int64 },
                    new() { Name = "price", Type = NativeColumnType.Float64 },
                    new() { Name = "active", Type = NativeColumnType.Bool },
                    new() { Name = "joined", Type = NativeColumnType.Date },
                ];
                Assert.Equal(NativeStatus.Ok, NativeApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable table));
                try
                {
                    Assert.Equal(5, table.ColumnCount);
                    Assert.Equal(2, table.RowCount);
                    Assert.Equal(["widget", "gadget"], DecodeStringColumn(ColumnAt(table, 0)));

                    long[] qty = new long[2];
                    Marshal.Copy(ColumnAt(table, 1).Values, qty, 0, 2);
                    Assert.Equal([3L, 7L], qty);

                    double[] prices = new double[2];
                    Marshal.Copy(ColumnAt(table, 2).Values, prices, 0, 2);
                    Assert.Equal([9.99, 4.5], prices);

                    byte[] flags = new byte[2];
                    Marshal.Copy(ColumnAt(table, 3).Values, flags, 0, 2);
                    Assert.Equal([(byte)1, (byte)0], flags);

                    int[] days = new int[2];
                    Marshal.Copy(ColumnAt(table, 4).Values, days, 0, 2);
                    int epoch = new DateOnly(1970, 1, 1).DayNumber;
                    Assert.Equal(new DateOnly(2024, 1, 15).DayNumber - epoch, days[0]);
                    Assert.Equal(new DateOnly(2024, 2, 20).DayNumber - epoch, days[1]);
                }
                finally
                {
                    NativeApi.FreeTable(ref table);
                    NativeApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseTyped_Should_Resolve_By_Index_When_Header_Row_Is_Zero()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "1,2\n3,4\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs =
                [
                    new() { Index = 0, Type = NativeColumnType.Int64 },
                    new() { Index = 1, Type = NativeColumnType.Int64 },
                ];
                Assert.Equal(NativeStatus.Ok, NativeApi.ParseTyped(handle, specs, headerRow: 0, out NativeTable table));
                try
                {
                    Assert.Equal(2, table.RowCount); // header_row == 0 means BOTH rows are data
                    long[] first = new long[2];
                    Marshal.Copy(ColumnAt(table, 0).Values, first, 0, 2);
                    long[] second = new long[2];
                    Marshal.Copy(ColumnAt(table, 1).Values, second, 0, 2);
                    Assert.Equal([1L, 3L], first);
                    Assert.Equal([2L, 4L], second);
                }
                finally
                {
                    NativeApi.FreeTable(ref table);
                    NativeApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseTyped_Should_Compute_Time_And_Timestamp_As_Microseconds()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "at,logged\n13:45:30,2024-01-15T13:45:30\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs =
                [
                    new() { Name = "at", Type = NativeColumnType.Time },
                    new() { Name = "logged", Type = NativeColumnType.Timestamp },
                ];
                Assert.Equal(NativeStatus.Ok, NativeApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable table));
                try
                {
                    long[] time = new long[1];
                    Marshal.Copy(ColumnAt(table, 0).Values, time, 0, 1);
                    Assert.Equal(new TimeOnly(13, 45, 30).ToTimeSpan().Ticks / 10, time[0]);

                    long[] timestamp = new long[1];
                    Marshal.Copy(ColumnAt(table, 1).Values, timestamp, 0, 1);
                    DateTime expected = new(2024, 1, 15, 13, 45, 30, DateTimeKind.Unspecified);
                    Assert.Equal((expected - DateTime.UnixEpoch).Ticks / 10, timestamp[0]);
                }
                finally
                {
                    NativeApi.FreeTable(ref table);
                    NativeApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        // The ABI's spec_count/name_len ceilings. xl_parse_typed and xl_parse_arrow enforce these
        // before either value sizes an allocation or drives a walk over caller memory, but those entry
        // points are [UnmanagedCallersOnly] and unreachable from managed code — so the predicate is
        // pinned here and the entry points themselves are covered by tests/ExcelReader.NativeSmoke.
        [Theory]
        [InlineData(int.MinValue, false)]
        [InlineData(-1, false)]
        [InlineData(0, false)] // a parse of zero columns is a caller mistake, not an empty result
        [InlineData(1, true)]
        [InlineData(16_384, true)] // A..XFD, the widest a real sheet can be
        [InlineData(16_385, false)]
        [InlineData(int.MaxValue, false)] // the shape that used to reach `new NativeColumnSpec[specCount]`
        public void IsValidSpecCount_Should_Accept_Only_One_Through_Excels_Column_Ceiling(int specCount, bool expected)
        {
            Assert.Equal(expected, NativeApi.IsValidSpecCount(specCount));
        }

        [Theory]
        [InlineData(int.MinValue, false)]
        [InlineData(-1, false)] // would reach Encoding.UTF8.GetString as a negative length
        [InlineData(0, true)] // an empty name is length-valid; TryValidateArguments rejects it later
        [InlineData(131_068, true)] // 32,767 chars at UTF-8's 4-byte worst case
        [InlineData(131_069, false)]
        [InlineData(int.MaxValue, false)]
        public void IsValidNameLength_Should_Bound_What_Becomes_A_Read_Length(int nameLength, bool expected)
        {
            Assert.Equal(expected, NativeApi.IsValidNameLength(nameLength));
        }

        [Fact]
        public void ParseTyped_Should_Reject_A_Blank_Column_Name()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            // A blank header sits in column 1: without the guard, the blank spec name would trim to ""
            // and resolve to it, silently reading a column the caller never named.
            File.WriteAllText(path, "name,,qty\nwidget,x,3\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs = [new() { Name = "   ", Type = NativeColumnType.String }];
                try
                {
                    Assert.Equal(NativeStatus.InvalidArgument, NativeApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable table));
                    Assert.Equal(IntPtr.Zero, table.Columns);

                    Span<byte> buffer = stackalloc byte[256];
                    Assert.Equal(NativeStatus.Ok, NativeApi.LastError(buffer, out int length));
                    Assert.Contains("blank name", Encoding.UTF8.GetString(buffer[..length]), StringComparison.Ordinal);
                }
                finally
                {
                    NativeApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        // The validity bitmap is accumulated eight rows to a byte, so every bug it can have lives at a
        // byte boundary: the last bit of a byte, the first bit of the next, and a final partial byte.
        // 20 rows with nulls at 0, 7, 8, 15, 16 and 19 put a null on each of those, which a three-row
        // fixture (see the test below) can never reach.
        [Fact]
        public void ParseTyped_Validity_Bitmap_Should_Survive_Byte_Boundaries()
        {
            const int rowCount = 20;
            int[] nullRows = [0, 7, 8, 15, 16, 19];

            StringBuilder csv = new("qty\n");
            for (int i = 0; i < rowCount; i++)
            {
                csv.Append(nullRows.Contains(i) ? "notanumber" : i.ToString(CultureInfo.InvariantCulture)).Append('\n');
            }

            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, csv.ToString());
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs = [new() { Name = "qty", Type = NativeColumnType.Int64, Nullable = true }];
                Assert.Equal(NativeStatus.Ok, NativeApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable table));
                try
                {
                    NativeColumn column = ColumnAt(table, 0);
                    Assert.Equal(rowCount, column.Length);
                    Assert.NotEqual(IntPtr.Zero, column.Validity);

                    bool[] expected = new bool[rowCount];
                    Array.Fill(expected, true);
                    foreach (int row in nullRows)
                    {
                        expected[row] = false;
                    }
                    Assert.Equal(expected, DecodeValidity(column));

                    // The values themselves must stay row-aligned with the bitmap: a null still occupies
                    // its slot, so a packing bug that shifted rows would show up here and not above.
                    long[] values = new long[rowCount];
                    Marshal.Copy(column.Values, values, 0, rowCount);
                    for (int i = 0; i < rowCount; i++)
                    {
                        Assert.Equal(nullRows.Contains(i) ? 0L : i, values[i]);
                    }
                }
                finally
                {
                    NativeApi.FreeTable(ref table);
                    NativeApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseTyped_Should_Mark_Failed_Nullable_Conversions_In_The_Validity_Bitmap()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "qty\n5\n\nnotanumber\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs = [new() { Name = "qty", Type = NativeColumnType.Int64, Nullable = true }];
                Assert.Equal(NativeStatus.Ok, NativeApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable table));
                try
                {
                    NativeColumn column = ColumnAt(table, 0);
                    Assert.NotEqual(IntPtr.Zero, column.Validity);
                    Assert.Equal([true, false, false], DecodeValidity(column));

                    long[] values = new long[3];
                    Marshal.Copy(column.Values, values, 0, 3);
                    Assert.Equal(5L, values[0]);
                }
                finally
                {
                    NativeApi.FreeTable(ref table);
                    NativeApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseTyped_Should_Fail_For_A_Non_Nullable_Conversion_Failure()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "qty\n5\nnotanumber\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs = [new() { Name = "qty", Type = NativeColumnType.Int64, Nullable = false }];

                int status = NativeApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable table);

                Assert.Equal(NativeStatus.Error, status);
                Assert.Equal(IntPtr.Zero, table.Columns);
                Span<byte> buffer = stackalloc byte[256];
                Assert.Equal(NativeStatus.Ok, NativeApi.LastError(buffer, out int length));
                Assert.True(length > 0);
                NativeApi.Close(handle);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseTyped_Should_Reject_A_Name_Based_Spec_When_Header_Row_Is_Zero()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                NativeColumnSpec[] specs = [new() { Name = "anything", Type = NativeColumnType.String }];
                Assert.Equal(NativeStatus.InvalidArgument, NativeApi.ParseTyped(handle, specs, headerRow: 0, out NativeTable table));
                Assert.Equal(IntPtr.Zero, table.Columns);
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void ParseTyped_Should_Reject_An_Unmatched_Header_Name()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name,qty\nwidget,3\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs = [new() { Name = "does-not-exist", Type = NativeColumnType.String }];

                int status = NativeApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable table);

                Assert.Equal(NativeStatus.InvalidArgument, status);
                Assert.Equal(IntPtr.Zero, table.Columns);
                NativeApi.Close(handle);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseTyped_Should_Reject_Zero_Specs()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.InvalidArgument, NativeApi.ParseTyped(handle, [], headerRow: 1, out NativeTable table));
                Assert.Equal(IntPtr.Zero, table.Columns);
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void ParseTyped_Should_Reject_A_Null_Handle()
        {
            NativeColumnSpec[] specs = [new() { Index = 0, Type = NativeColumnType.String }];
            Assert.Equal(NativeStatus.InvalidHandle, NativeApi.ParseTyped(null, specs, headerRow: 1, out _));
        }

        [Fact]
        public void ParseTyped_Should_Not_Disturb_The_Row_Cursor()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name\nfirst\nsecond\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                try
                {
                    Assert.Equal(NativeStatus.Ok, NativeApi.NextRow(handle, new byte[4096], out _)); // header

                    NativeColumnSpec[] specs = [new() { Name = "name", Type = NativeColumnType.String }];
                    Assert.Equal(NativeStatus.Ok, NativeApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable table));
                    NativeApi.FreeTable(ref table);

                    // ParseTyped reads the WHOLE sheet through its own independent enumerator - the
                    // xl_next_row cursor above must still be sitting right after the header row.
                    byte[] buffer = new byte[4096];
                    Assert.Equal(NativeStatus.Ok, NativeApi.NextRow(handle, buffer, out int written));
                    Assert.Equal("first", DecodeRow(buffer.AsSpan(0, written))[0].Value);
                }
                finally
                {
                    NativeApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void FreeTable_Should_Be_Idempotent_On_A_Zeroed_Table()
        {
            NativeTable table = default;
            NativeApi.FreeTable(ref table);
            NativeApi.FreeTable(ref table);
            Assert.Equal(IntPtr.Zero, table.Columns);
        }

        // Column layout: A=Name (string), B=Qty (whole numbers), C=Price (fractional numbers),
        // D=Active (bool), E=Mixed (a string in one row, a number in the other), F=Extra (a number
        // that only row 2 populates, and that never appears in the header at all).
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

        [Fact]
        public void InferSchema_Should_Guess_Types_From_Sampled_Cells()
        {
            using MemoryStream ms = BuildInferSchemaFixture();
            Assert.Equal(NativeStatus.Ok, NativeApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, NativeApi.InferSchema(handle, headerRow: 1, sampleSize: 100, out NativeInferredSchema schema));
                try
                {
                    Assert.Equal(6, schema.ColumnCount);
                    (string? Name, int Index, int Type, bool Nullable)[] columns = DecodeSchema(schema);

                    AssertSpec(columns[0], "Name", NativeColumnType.String, nullable: false);
                    AssertSpec(columns[1], "Qty", NativeColumnType.Int64, nullable: false);
                    AssertSpec(columns[2], "Price", NativeColumnType.Float64, nullable: false);
                    AssertSpec(columns[3], "Active", NativeColumnType.Bool, nullable: false);
                    // A string in one sampled row and a number in the other is a real mix - no single
                    // type describes both, so this falls back to STRING.
                    AssertSpec(columns[4], "Mixed", NativeColumnType.String, nullable: false);
                    // Never named in the header, and only row 2 populates it - the missing value in
                    // row 3 must be caught even though no row's cells enumerator ever visits column F.
                    AssertSpec(columns[5], name: null, NativeColumnType.Int64, nullable: true);
                    Assert.Equal(5, columns[5].Index);
                }
                finally
                {
                    NativeApi.FreeSchema(ref schema);
                }
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void InferSchema_Should_Resolve_By_Index_When_Header_Row_Is_Zero()
        {
            using MemoryStream ms = BuildInferSchemaFixture();
            Assert.Equal(NativeStatus.Ok, NativeApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, NativeApi.InferSchema(handle, headerRow: 0, sampleSize: 100, out NativeInferredSchema schema));
                try
                {
                    // Every row (including what would have been the header) is now sampled as data, so
                    // row 1's inline strings make every column look like STRING with no header names.
                    foreach ((string? name, _, _, _) in DecodeSchema(schema))
                    {
                        Assert.Null(name);
                    }
                }
                finally
                {
                    NativeApi.FreeSchema(ref schema);
                }
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void InferSchema_Should_Reject_A_Negative_Header_Row()
        {
            using MemoryStream ms = BuildInferSchemaFixture();
            Assert.Equal(NativeStatus.Ok, NativeApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.InvalidArgument, NativeApi.InferSchema(handle, headerRow: -1, sampleSize: 100, out NativeInferredSchema schema));
                Assert.Equal(IntPtr.Zero, schema.Columns);
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Theory]
        [InlineData(0)]
        [InlineData(-1)]
        public void InferSchema_Should_Reject_A_NonPositive_Sample_Size(int sampleSize)
        {
            using MemoryStream ms = BuildInferSchemaFixture();
            Assert.Equal(NativeStatus.Ok, NativeApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.InvalidArgument, NativeApi.InferSchema(handle, headerRow: 1, sampleSize, out NativeInferredSchema schema));
                Assert.Equal(IntPtr.Zero, schema.Columns);
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void InferSchema_Should_Reject_A_Null_Handle()
        {
            Assert.Equal(NativeStatus.InvalidHandle, NativeApi.InferSchema(null, headerRow: 1, sampleSize: 100, out _));
        }

        [Fact]
        public void InferSchema_Should_Report_An_Error_When_The_Sheet_Has_Fewer_Rows_Than_Header_Row()
        {
            using MemoryStream ms = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>1</v></c></row>""");
            Assert.Equal(NativeStatus.Ok, NativeApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                int status = NativeApi.InferSchema(handle, headerRow: 5, sampleSize: 100, out NativeInferredSchema schema);

                Assert.Equal(NativeStatus.InvalidArgument, status);
                Assert.Equal(IntPtr.Zero, schema.Columns);
                Span<byte> buffer = stackalloc byte[256];
                Assert.Equal(NativeStatus.Ok, NativeApi.LastError(buffer, out int length));
                Assert.Contains("fewer than", Encoding.UTF8.GetString(buffer[..length]), StringComparison.Ordinal);
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void InferSchema_Should_Not_Disturb_The_Row_Cursor()
        {
            using MemoryStream ms = BuildInferSchemaFixture();
            Assert.Equal(NativeStatus.Ok, NativeApi.OpenMemory(ms.ToArray(), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, NativeApi.NextRow(handle, new byte[4096], out _)); // header

                Assert.Equal(NativeStatus.Ok, NativeApi.InferSchema(handle, headerRow: 1, sampleSize: 100, out NativeInferredSchema schema));
                NativeApi.FreeSchema(ref schema);

                // InferSchema reads the WHOLE sheet through its own independent enumerator - the
                // xl_next_row cursor above must still be sitting right after the header row.
                byte[] buffer = new byte[4096];
                Assert.Equal(NativeStatus.Ok, NativeApi.NextRow(handle, buffer, out int written));
                Assert.Equal("Alice", DecodeRow(buffer.AsSpan(0, written))[0].Value);
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void FreeSchema_Should_Be_Idempotent_On_A_Zeroed_Schema()
        {
            NativeInferredSchema schema = default;
            NativeApi.FreeSchema(ref schema);
            NativeApi.FreeSchema(ref schema);
            Assert.Equal(IntPtr.Zero, schema.Columns);
        }

        private static void AssertSpec((string? Name, int Index, int Type, bool Nullable) spec, string? name, int type, bool nullable)
        {
            Assert.Equal(name, spec.Name);
            Assert.Equal(type, spec.Type);
            Assert.Equal(nullable, spec.Nullable);
        }

        // Reads the native xl_column_spec array by hand rather than via NativeColumnSpecRaw's `byte*
        // Name` field - this project has no AllowUnsafeBlocks, and Marshal.PtrToStringUTF8 already
        // decodes exactly `length` bytes regardless of a trailing NUL (there isn't one; see
        // NativeApi.Schema.cs's BuildSpec), so there is nothing an unsafe pointer read would add here.
        private static (string? Name, int Index, int Type, bool Nullable)[] DecodeSchema(NativeInferredSchema schema)
        {
            int specSize = Marshal.SizeOf<NativeColumnSpecRaw>(); // name(8) + name_len(4) + index(4) + type(4) + nullable(4)
            var columns = new (string?, int, int, bool)[schema.ColumnCount];
            for (int i = 0; i < columns.Length; i++)
            {
                IntPtr spec = IntPtr.Add(schema.Columns, i * specSize);
                IntPtr namePtr = Marshal.ReadIntPtr(spec, 0);
                int nameLen = Marshal.ReadInt32(spec, 8);
                int index = Marshal.ReadInt32(spec, 12);
                int type = Marshal.ReadInt32(spec, 16);
                int nullable = Marshal.ReadInt32(spec, 20);
                string? name = namePtr == IntPtr.Zero ? null : Marshal.PtrToStringUTF8(namePtr, nameLen);
                columns[i] = (name, index, type, nullable != 0);
            }
            return columns;
        }

        private static ArrowSchema ArrowChildSchema(ArrowSchema schema, int index)
        {
            return Marshal.PtrToStructure<ArrowSchema>(Marshal.ReadIntPtr(schema.Children, index * IntPtr.Size));
        }

        private static ArrowArray ArrowChildArray(ArrowArray array, int index)
        {
            return Marshal.PtrToStructure<ArrowArray>(Marshal.ReadIntPtr(array.Children, index * IntPtr.Size));
        }

        private static IntPtr ArrowBuffer(ArrowArray array, int index)
        {
            return Marshal.ReadIntPtr(array.Buffers, index * IntPtr.Size);
        }

        [Fact]
        public void ParseArrow_Should_Return_A_Struct_Array_With_A_Matching_Schema()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name,qty\nwidget,3\ngadget,7\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs =
                [
                    new() { Name = "name", Type = NativeColumnType.String },
                    new() { Name = "qty", Type = NativeColumnType.Int64 },
                ];
                Assert.Equal(NativeStatus.Ok, NativeApi.ParseArrow(handle, specs, headerRow: 1, out ArrowArray array, out ArrowSchema schema));
                try
                {
                    Assert.Equal("+s", Marshal.PtrToStringUTF8(schema.Format));
                    Assert.Equal(2, schema.NChildren);
                    Assert.Equal("u", Marshal.PtrToStringUTF8(ArrowChildSchema(schema, 0).Format));
                    Assert.Equal("name", Marshal.PtrToStringUTF8(ArrowChildSchema(schema, 0).Name));
                    Assert.Equal("l", Marshal.PtrToStringUTF8(ArrowChildSchema(schema, 1).Format));

                    Assert.Equal(2, array.Length); // row count
                    Assert.Equal(2, array.NChildren);
                    Assert.Equal(1, array.NBuffers); // top-level struct array: validity only, always absent here
                    Assert.Equal(IntPtr.Zero, ArrowBuffer(array, 0));

                    ArrowArray nameColumn = ArrowChildArray(array, 0);
                    Assert.Equal(2, nameColumn.Length);
                    Assert.Equal(0, nameColumn.NullCount);
                    Assert.Equal(3, nameColumn.NBuffers); // validity, offsets, data
                    Assert.Equal(IntPtr.Zero, ArrowBuffer(nameColumn, 0)); // never null
                    int[] offsets = new int[3];
                    Marshal.Copy(ArrowBuffer(nameColumn, 1), offsets, 0, 3);
                    byte[] data = new byte[offsets[2]];
                    Marshal.Copy(ArrowBuffer(nameColumn, 2), data, 0, data.Length);
                    Assert.Equal("widget", Encoding.UTF8.GetString(data, offsets[0], offsets[1] - offsets[0]));
                    Assert.Equal("gadget", Encoding.UTF8.GetString(data, offsets[1], offsets[2] - offsets[1]));

                    ArrowArray qtyColumn = ArrowChildArray(array, 1);
                    Assert.Equal(2, qtyColumn.NBuffers); // validity, values
                    long[] qty = new long[2];
                    Marshal.Copy(ArrowBuffer(qtyColumn, 1), qty, 0, 2);
                    Assert.Equal([3L, 7L], qty);
                }
                finally
                {
                    ExercisedReleaseArrow(ref array, ref schema);
                    NativeApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseArrow_Should_Bit_Pack_Bool_Columns()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            // 10 values so the bitmap spans two bytes: true,false alternating plus a tail.
            File.WriteAllText(path, "flag\ntrue\nfalse\ntrue\nfalse\ntrue\nfalse\ntrue\nfalse\ntrue\ntrue\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs = [new() { Name = "flag", Type = NativeColumnType.Bool }];
                Assert.Equal(NativeStatus.Ok, NativeApi.ParseArrow(handle, specs, headerRow: 1, out ArrowArray array, out ArrowSchema schema));
                try
                {
                    ArrowArray column = ArrowChildArray(array, 0);
                    Assert.Equal(10, column.Length);
                    Assert.Equal("b", Marshal.PtrToStringUTF8(ArrowChildSchema(schema, 0).Format));

                    byte[] bitmap = new byte[2];
                    Marshal.Copy(ArrowBuffer(column, 1), bitmap, 0, 2);
                    bool[] expected = [true, false, true, false, true, false, true, false, true, true];
                    for (int i = 0; i < expected.Length; i++)
                    {
                        bool bit = (bitmap[i >> 3] & (1 << (i & 7))) != 0;
                        Assert.Equal(expected[i], bit);
                    }
                }
                finally
                {
                    ExercisedReleaseArrow(ref array, ref schema);
                    NativeApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseArrow_Should_Report_Null_Count_From_The_Validity_Bitmap()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "qty\n5\n\nnotanumber\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs = [new() { Name = "qty", Type = NativeColumnType.Int64, Nullable = true }];
                Assert.Equal(NativeStatus.Ok, NativeApi.ParseArrow(handle, specs, headerRow: 1, out ArrowArray array, out ArrowSchema schema));
                try
                {
                    ArrowArray column = ArrowChildArray(array, 0);
                    Assert.NotEqual(IntPtr.Zero, ArrowBuffer(column, 0));
                    Assert.Equal(2, column.NullCount);
                    Assert.True((ArrowChildSchema(schema, 0).Flags & ArrowFlags.Nullable) != 0);
                }
                finally
                {
                    ExercisedReleaseArrow(ref array, ref schema);
                    NativeApi.Close(handle);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseArrow_Should_Propagate_A_Non_Nullable_Conversion_Failure()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "qty\nnotanumber\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs = [new() { Name = "qty", Type = NativeColumnType.Int64 }];

                int status = NativeApi.ParseArrow(handle, specs, headerRow: 1, out ArrowArray array, out ArrowSchema schema);

                Assert.Equal(NativeStatus.Error, status);
                Assert.Equal(IntPtr.Zero, array.Release);
                Assert.Equal(IntPtr.Zero, schema.Release);
                NativeApi.Close(handle);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ParseArrow_Should_Reject_A_Null_Handle()
        {
            NativeColumnSpec[] specs = [new() { Index = 0, Type = NativeColumnType.String }];
            Assert.Equal(NativeStatus.InvalidHandle, NativeApi.ParseArrow(null, specs, headerRow: 1, out _, out _));
        }

        [Fact]
        public void ReleaseArrowArray_And_ReleaseArrowSchema_Are_Idempotent()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name\nwidget\n");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
                NativeColumnSpec[] specs = [new() { Name = "name", Type = NativeColumnType.String }];
                Assert.Equal(NativeStatus.Ok, NativeApi.ParseArrow(handle, specs, headerRow: 1, out ArrowArray array, out ArrowSchema schema));

                ExercisedReleaseArrow(ref array, ref schema);
                // A second release on the same (now-zeroed-release) struct must be a harmless no-op.
                ExercisedReleaseArrow(ref array, ref schema);

                Assert.Equal(IntPtr.Zero, array.Release);
                Assert.Equal(IntPtr.Zero, schema.Release);
                NativeApi.Close(handle);
            }
            finally
            {
                File.Delete(path);
            }
        }

        private static NativeWriteOptionsRaw DefaultWriteOptionsRaw()
        {
            return new NativeWriteOptionsRaw { StructSize = Marshal.SizeOf<NativeWriteOptionsRaw>() };
        }

        [Fact]
        public void WriteOptions_Should_Decode_An_All_Defaults_Struct()
        {
            Assert.True(NativeWriteOptions.TryDecode(DefaultWriteOptionsRaw(), null, out NativeWriteOptions options, out string? error));

            Assert.Null(error);
            Assert.Null(options.SheetName);
            Assert.Null(options.CsvDelimiter);
            Assert.Null(options.CsvQuote);
            Assert.Null(options.Date1904);
            Assert.Null(options.UseSharedStrings);
        }

        [Fact]
        public void WriteOptions_Should_Reject_An_Unrecognized_Struct_Size()
        {
            NativeWriteOptionsRaw raw = DefaultWriteOptionsRaw();
            raw.StructSize = 1;

            Assert.False(NativeWriteOptions.TryDecode(raw, null, out _, out string? error));
            Assert.Contains("struct_size", error, StringComparison.Ordinal);
        }

        [Fact]
        public void WriteOptions_Should_Reject_A_Csv_Delimiter_Outside_A_Byte()
        {
            NativeWriteOptionsRaw raw = DefaultWriteOptionsRaw();
            raw.CsvDelimiter = 300;

            Assert.False(NativeWriteOptions.TryDecode(raw, null, out _, out string? error));
            Assert.Contains("csv_delimiter", error, StringComparison.Ordinal);
        }

        [Fact]
        public void WriteOptions_Should_Reject_A_Sheet_Name_Excel_Cannot_Store()
        {
            NativeWriteOptionsRaw raw = DefaultWriteOptionsRaw();

            Assert.False(NativeWriteOptions.TryDecode(raw, "has/slash", out _, out string? error));
            Assert.Contains("sheet_name", error, StringComparison.Ordinal);
            Assert.False(NativeWriteOptions.TryDecode(raw, "", out _, out error));
            Assert.Contains("sheet_name", error, StringComparison.Ordinal);
            Assert.False(NativeWriteOptions.TryDecode(raw, new string('x', 32), out _, out error));
            Assert.Contains("sheet_name", error, StringComparison.Ordinal);
        }

        [Fact]
        public void WriteOptions_Should_Carry_Overrides_Into_CsvWriterOptions()
        {
            NativeWriteOptionsRaw raw = DefaultWriteOptionsRaw();
            raw.CsvDelimiter = ';';
            raw.CsvQuote = '\'';
            raw.Date1904 = NativeOptionState.True;
            raw.UseSharedStrings = NativeOptionState.True;

            Assert.True(NativeWriteOptions.TryDecode(raw, "Data", out NativeWriteOptions options, out _));

            Assert.Equal("Data", options.SheetName);
            Assert.True(options.Date1904);
            Assert.True(options.UseSharedStrings);
            CsvWriterOptions csv = options.ToCsvWriterOptions();
            Assert.Equal((byte)';', csv.Delimiter);
            Assert.Equal((byte)'\'', csv.Quote);
        }

        private static NativeWriteOptions DefaultWriteOptions()
        {
            Assert.True(NativeWriteOptions.TryDecode(DefaultWriteOptionsRaw(), null, out NativeWriteOptions options, out _));
            return options;
        }

        [Theory]
        [InlineData(NativeFormat.Xlsx, "xlsx")]
        [InlineData(NativeFormat.Xlsb, "xlsb")]
        [InlineData(NativeFormat.Xls, "xls")]
        [InlineData(NativeFormat.Csv, "csv")]
        public void WriteTyped_Should_Round_Trip_Through_ParseTyped(int format, string extension)
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.{extension}");
            NativeTable table = BuildInt64Table([3L, 7L]);
            try
            {
                NativeColumnSpec[] specs = [new() { Name = "qty", Type = NativeColumnType.Int64 }];

                Assert.Equal(NativeStatus.Ok, NativeApi.WriteTyped(
                    Encoding.UTF8.GetBytes(path), format, specs, table, DefaultWriteOptions()));

                Assert.Equal(NativeStatus.Ok, OpenPath(path, format, out NativeHandle? handle));
                try
                {
                    Assert.Equal(NativeStatus.Ok, NativeApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable read));
                    try
                    {
                        Assert.Equal(2, read.RowCount);
                        long[] values = new long[2];
                        Marshal.Copy(ColumnAt(read, 0).Values, values, 0, 2);
                        Assert.Equal([3L, 7L], values);
                    }
                    finally
                    {
                        NativeApi.FreeTable(ref read);
                    }
                }
                finally
                {
                    NativeApi.Close(handle);
                }
            }
            finally
            {
                FreeBuiltTable(ref table);
                File.Delete(path);
            }
        }

        [Fact]
        public void WriteTyped_Should_Round_Trip_Strings_And_Nulls()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.xlsx");
            NativeTable table = BuildStringTable([0, 6, 6, 12], "widgetgadget"u8.ToArray());
            try
            {
                // Row 1 is null (bit 1 clear), rows 0 and 2 are valid: 0b101.
                NativeColumn column = Marshal.PtrToStructure<NativeColumn>(table.Columns);
                column.Validity = Marshal.AllocHGlobal(1);
                Marshal.WriteByte(column.Validity, 0b101);
                Marshal.StructureToPtr(column, table.Columns, false);

                NativeColumnSpec[] specs = [new() { Name = "name", Type = NativeColumnType.String }];
                Assert.Equal(NativeStatus.Ok, NativeApi.WriteTyped(
                    Encoding.UTF8.GetBytes(path), NativeFormat.Xlsx, specs, table, DefaultWriteOptions()));

                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Xlsx, out NativeHandle? handle));
                try
                {
                    NativeColumnSpec[] readSpecs = [new() { Name = "name", Type = NativeColumnType.String, Nullable = true }];
                    Assert.Equal(NativeStatus.Ok, NativeApi.ParseTyped(handle, readSpecs, headerRow: 1, out NativeTable read));
                    try
                    {
                        Assert.Equal(["widget", "", "gadget"], DecodeStringColumn(ColumnAt(read, 0)));
                    }
                    finally
                    {
                        NativeApi.FreeTable(ref read);
                    }
                }
                finally
                {
                    NativeApi.Close(handle);
                }
            }
            finally
            {
                FreeBuiltTable(ref table);
                File.Delete(path);
            }
        }

        [Fact]
        public void WriteTyped_Should_Round_Trip_Every_Temporal_Type()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.xlsx");
            int epoch = new DateOnly(1970, 1, 1).DayNumber;
            int day = new DateOnly(2024, 1, 15).DayNumber - epoch;
            long clock = 3_600_000_000L;                                   // 01:00:00
            long stamp = 1_705_280_400_000_000L;                           // 2024-01-15T01:00:00Z

            NativeTable table = BuildTemporalTable(day, clock, stamp);
            try
            {
                NativeColumnSpec[] specs =
                [
                    new() { Name = "day", Type = NativeColumnType.Date },
                    new() { Name = "clock", Type = NativeColumnType.Time },
                    new() { Name = "stamp", Type = NativeColumnType.Timestamp },
                ];
                Assert.Equal(NativeStatus.Ok, NativeApi.WriteTyped(
                    Encoding.UTF8.GetBytes(path), NativeFormat.Xlsx, specs, table, DefaultWriteOptions()));

                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Xlsx, out NativeHandle? handle));
                try
                {
                    Assert.Equal(NativeStatus.Ok, NativeApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable read));
                    try
                    {
                        int[] days = new int[1];
                        Marshal.Copy(ColumnAt(read, 0).Values, days, 0, 1);
                        long[] clocks = new long[1];
                        Marshal.Copy(ColumnAt(read, 1).Values, clocks, 0, 1);
                        long[] stamps = new long[1];
                        Marshal.Copy(ColumnAt(read, 2).Values, stamps, 0, 1);

                        Assert.Equal(day, days[0]);
                        Assert.Equal(clock, clocks[0]);
                        Assert.Equal(stamp, stamps[0]);
                    }
                    finally
                    {
                        NativeApi.FreeTable(ref read);
                    }
                }
                finally
                {
                    NativeApi.Close(handle);
                }
            }
            finally
            {
                FreeBuiltTable(ref table);
                File.Delete(path);
            }
        }

        private static NativeTable BuildTemporalTable(int day, long clock, long stamp)
        {
            NativeColumn[] columns =
            [
                new() { Type = NativeColumnType.Date, Length = 1, Values = Marshal.AllocHGlobal(sizeof(int)) },
                new() { Type = NativeColumnType.Time, Length = 1, Values = Marshal.AllocHGlobal(sizeof(long)) },
                new() { Type = NativeColumnType.Timestamp, Length = 1, Values = Marshal.AllocHGlobal(sizeof(long)) },
            ];
            Marshal.WriteInt32(columns[0].Values, day);
            Marshal.WriteInt64(columns[1].Values, clock);
            Marshal.WriteInt64(columns[2].Values, stamp);

            int size = Marshal.SizeOf<NativeColumn>();
            IntPtr block = Marshal.AllocHGlobal(size * columns.Length);
            for (int index = 0; index < columns.Length; index++)
            {
                Marshal.StructureToPtr(columns[index], IntPtr.Add(block, index * size), false);
            }
            return new NativeTable { ColumnCount = columns.Length, RowCount = 1, Columns = block };
        }

        [Fact]
        public void WriteTyped_Should_Write_No_Header_Row_When_Specs_Are_Unnamed()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            NativeTable table = BuildInt64Table([3L, 7L]);
            try
            {
                NativeColumnSpec[] specs = [new() { Name = null, Type = NativeColumnType.Int64 }];
                Assert.Equal(NativeStatus.Ok, NativeApi.WriteTyped(
                    Encoding.UTF8.GetBytes(path), NativeFormat.Csv, specs, table, DefaultWriteOptions()));

                Assert.Equal("3\n7\n", File.ReadAllText(path).ReplaceLineEndings("\n"));
            }
            finally
            {
                FreeBuiltTable(ref table);
                File.Delete(path);
            }
        }

        [Fact]
        public void WriteTyped_Should_Apply_The_Csv_Delimiter_Override()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            NativeTable table = BuildInt64Table([3L]);
            try
            {
                NativeWriteOptionsRaw raw = DefaultWriteOptionsRaw();
                raw.CsvDelimiter = ';';
                Assert.True(NativeWriteOptions.TryDecode(raw, null, out NativeWriteOptions options, out _));
                NativeColumnSpec[] specs = [new() { Name = "qty", Type = NativeColumnType.Int64 }];

                Assert.Equal(NativeStatus.Ok, NativeApi.WriteTyped(
                    Encoding.UTF8.GetBytes(path), NativeFormat.Csv, specs, table, options));

                Assert.StartsWith("qty", File.ReadAllText(path), StringComparison.Ordinal);
            }
            finally
            {
                FreeBuiltTable(ref table);
                File.Delete(path);
            }
        }

        [Fact]
        public void WriteTyped_Should_Reject_Auto_Format_And_Create_No_File()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.xlsx");
            NativeTable table = BuildInt64Table([3L]);
            try
            {
                NativeColumnSpec[] specs = [new() { Name = "qty", Type = NativeColumnType.Int64 }];

                Assert.Equal(NativeStatus.InvalidArgument, NativeApi.WriteTyped(
                    Encoding.UTF8.GetBytes(path), NativeFormat.Auto, specs, table, DefaultWriteOptions()));
                Assert.False(File.Exists(path));
            }
            finally
            {
                FreeBuiltTable(ref table);
            }
        }

        [Fact]
        public void WriteTyped_Should_Create_No_File_When_The_Table_Is_Rejected()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.xlsx");
            NativeTable table = BuildStringTable([0, 9, 6], "widgetgadget"u8.ToArray());
            try
            {
                NativeColumnSpec[] specs = [new() { Name = "name", Type = NativeColumnType.String }];

                Assert.Equal(NativeStatus.InvalidArgument, NativeApi.WriteTyped(
                    Encoding.UTF8.GetBytes(path), NativeFormat.Xlsx, specs, table, DefaultWriteOptions()));
                Assert.False(File.Exists(path));
            }
            finally
            {
                FreeBuiltTable(ref table);
            }
        }

        [Fact]
        public void WriteTyped_Should_Use_The_Requested_Sheet_Name()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.xlsx");
            NativeTable table = BuildInt64Table([3L]);
            try
            {
                Assert.True(NativeWriteOptions.TryDecode(DefaultWriteOptionsRaw(), "Vendas", out NativeWriteOptions options, out _));
                NativeColumnSpec[] specs = [new() { Name = "qty", Type = NativeColumnType.Int64 }];

                Assert.Equal(NativeStatus.Ok, NativeApi.WriteTyped(
                    Encoding.UTF8.GetBytes(path), NativeFormat.Xlsx, specs, table, options));

                Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Xlsx, out NativeHandle? handle));
                try
                {
                    byte[] name = new byte[64];
                    Assert.Equal(NativeStatus.Ok, NativeApi.SheetNameAt(handle, 0, name, out int written));
                    Assert.Equal("Vendas", Encoding.UTF8.GetString(name, 0, written));
                }
                finally
                {
                    NativeApi.Close(handle);
                }
            }
            finally
            {
                FreeBuiltTable(ref table);
                File.Delete(path);
            }
        }

        // Releases via the same IntPtr round-trip real Arrow consumers use (their own storage, not a
        // pointer to this ref struct) - NativeApi.ReleaseArrowArray/Schema take an address, and `array`/
        // `schema` here are plain locals with no fixed address of their own.
        private static void ExercisedReleaseArrow(ref ArrowArray array, ref ArrowSchema schema)
        {
            IntPtr arrayBlock = Marshal.AllocHGlobal(Marshal.SizeOf<ArrowArray>());
            IntPtr schemaBlock = Marshal.AllocHGlobal(Marshal.SizeOf<ArrowSchema>());
            try
            {
                Marshal.StructureToPtr(array, arrayBlock, false);
                Marshal.StructureToPtr(schema, schemaBlock, false);
                NativeApi.ReleaseArrowArray(arrayBlock);
                NativeApi.ReleaseArrowSchema(schemaBlock);
                array = Marshal.PtrToStructure<ArrowArray>(arrayBlock);
                schema = Marshal.PtrToStructure<ArrowSchema>(schemaBlock);
            }
            finally
            {
                Marshal.FreeHGlobal(arrayBlock);
                Marshal.FreeHGlobal(schemaBlock);
            }
        }

        // NativeHandle's constructor is internal; this project has InternalsVisibleTo access to it, so
        // tests can hand it a reader that isn't produced by NativeApi.OpenFile/OpenMemory.
        private static NativeHandle NativeHandle_Create(IExcelRowReader reader)
        {
            return new NativeHandle(reader);
        }

        /// <summary>
        /// Wraps a real <see cref="IExcelRowReader"/>, forwarding everything except row enumeration:
        /// its enumerator yields <paramref name="failAfter"/> real rows and then throws, simulating a
        /// genuine mid-sheet decode failure for <see cref="ReadAllDecoded_Should_Free_Already_Decoded_Rows_When_A_Later_Row_Fails_To_Decode"/>.
        /// </summary>
#pragma warning disable IDISP007 // Don't dispose injected: this wrapper owns `inner` for the test's duration by construction — nothing else disposes it.
#pragma warning disable HLQ006 // A reference-type enumerator is intentional here: this decorates the format-agnostic IExcelRowEnumerator interface, not a zero-allocation hot path.
        private sealed class FailAfterNRowsReader(IExcelRowReader inner, int failAfter) : IExcelRowReader
        {
            public bool IsDate1904 => inner.IsDate1904;
            public string SheetName => inner.SheetName;
            public int SheetCount => inner.SheetCount;
            public string SheetNameAt(int index) => inner.SheetNameAt(index);

            public bool TryMoveToSheet(ReadOnlySpan<char> name) => inner.TryMoveToSheet(name);
            public void MoveToSheet(int index) => inner.MoveToSheet(index);

            public IExcelRowEnumerator GetEnumerator() => new FailAfterNRowsEnumerator(inner.GetEnumerator(), failAfter);
            public IExcelRowEnumerator GetAsyncEnumerator() => new FailAfterNRowsEnumerator(inner.GetAsyncEnumerator(), failAfter);
            public ValueTask<IExcelRowEnumerator> GetAsyncEnumeratorAsync(CancellationToken ct = default) =>
                new(new FailAfterNRowsEnumerator(inner.GetAsyncEnumerator(), failAfter));

            public void Dispose() => inner.Dispose();
            public ValueTask DisposeAsync() => inner.DisposeAsync();
        }
#pragma warning restore HLQ006
#pragma warning restore IDISP007

#pragma warning disable IDISP007 // Don't dispose injected: this wrapper owns `inner` for the test's duration by construction — nothing else disposes it.
        private sealed class FailAfterNRowsEnumerator(IExcelRowEnumerator inner, int failAfter) : IExcelRowEnumerator
        {
            private int _moveNextCalls;

            public Row Current => inner.Current;

            public bool MoveNext()
            {
                if (_moveNextCalls++ >= failAfter)
                {
                    throw new InvalidOperationException("Forced decode failure for test purposes.");
                }
                return inner.MoveNext();
            }

            public ValueTask<bool> MoveNextAsync()
            {
                if (_moveNextCalls++ >= failAfter)
                {
                    throw new InvalidOperationException("Forced decode failure for test purposes.");
                }
                return inner.MoveNextAsync();
            }

            public void Dispose() => inner.Dispose();
            public ValueTask DisposeAsync() => inner.DisposeAsync();
        }
#pragma warning restore IDISP007
    }
}
