using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
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
