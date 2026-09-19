using System.Text;
using ExcelReader.Native;
using ExcelReader.Native.Writer;

namespace ExcelReader.Tests
{
    /// <summary>
    /// Drives <see cref="NativeApi.OpenWriteHandle"/>/<see cref="NativeApi.CloseWriteHandle"/> and
    /// <see cref="NativeWriterHandle"/>'s state machine directly — the layer <c>Exports</c>'s
    /// [UnmanagedCallersOnly] entry points delegate to, and the only layer this test host can call.
    /// </summary>
    public sealed class NativeWriterHandleTests
    {
        private static string TempPath(string extension)
        {
            return Path.Combine(Path.GetTempPath(), $"excelreader-writer-handle-{Guid.NewGuid():N}.{extension}");
        }

        private static int OpenWriteHandle(string path, int format, out NativeWriterHandle? handle, NativeWriteOptions options = default)
        {
            return NativeApi.OpenWriteHandle(Encoding.UTF8.GetBytes(path), format, options, out handle);
        }

        private static int OpenWriteHandleToMemory(int format, out NativeWriterHandle? handle, NativeWriteOptions options = default)
        {
            return NativeApi.OpenWriteHandleToMemory(format, options, out handle);
        }

        [Theory]
        [InlineData(NativeFormat.Xlsx, "xlsx")]
        [InlineData(NativeFormat.Xlsb, "xlsb")]
        [InlineData(NativeFormat.Xls, "xls")]
        [InlineData(NativeFormat.Csv, "csv")]
        public void OpenWriteHandle_Should_Round_Trip_A_Streamed_Row_Through_OpenFile(int format, string extension)
        {
            string path = TempPath(extension);
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenWriteHandle(path, format, out NativeWriterHandle? handle));
                Assert.NotNull(handle);

                handle.StartSheet("Dados");
                handle.StartRow();
                handle.WriteString("uma");
                handle.WriteInt64(3);
                handle.EndRow();

                Assert.Equal(NativeStatus.Ok, NativeApi.CloseWriteHandle(handle));

                Assert.Equal(NativeStatus.Ok, NativeApi.OpenFile(Encoding.UTF8.GetBytes(path), format, out NativeHandle? reader));
                Assert.NotNull(reader);
                try
                {
                    Span<byte> buffer = stackalloc byte[256];
                    Assert.Equal(NativeStatus.Ok, NativeApi.NextRow(reader, buffer, out int written));
                    string row = Encoding.UTF8.GetString(buffer[..written]);
                    Assert.Contains("uma", row, StringComparison.Ordinal);
                    Assert.Contains("3", row, StringComparison.Ordinal);
                }
                finally
                {
                    NativeApi.Close(reader);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void OpenWriteHandle_Should_Reject_An_Empty_Path()
        {
            Assert.Equal(NativeStatus.InvalidArgument,
                NativeApi.OpenWriteHandle(ReadOnlySpan<byte>.Empty, NativeFormat.Xlsx, default, out NativeWriterHandle? handle));
            Assert.Null(handle);
        }

        [Fact]
        public void OpenWriteHandle_Should_Reject_Auto_Format_And_Create_No_File()
        {
            string path = TempPath("xlsx");
            try
            {
                Assert.Equal(NativeStatus.InvalidArgument, OpenWriteHandle(path, NativeFormat.Auto, out NativeWriterHandle? handle));
                Assert.Null(handle);
                Assert.False(File.Exists(path));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void StartRow_Should_Throw_Before_StartSheet()
        {
            string path = TempPath("xlsx");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenWriteHandle(path, NativeFormat.Xlsx, out NativeWriterHandle? handle));
                Assert.Throws<InvalidOperationException>(() => handle!.StartRow());

                handle!.StartSheet("S");
                Assert.Equal(NativeStatus.Ok, NativeApi.CloseWriteHandle(handle));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void WriteInt64_Should_Throw_Before_StartRow()
        {
            string path = TempPath("xlsx");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenWriteHandle(path, NativeFormat.Xlsx, out NativeWriterHandle? handle));
                handle!.StartSheet("S");
                Assert.Throws<InvalidOperationException>(() => handle.WriteInt64(1));
                Assert.Equal(NativeStatus.Ok, NativeApi.CloseWriteHandle(handle));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void EndRow_Should_Throw_Without_A_Matching_StartRow()
        {
            string path = TempPath("xlsx");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenWriteHandle(path, NativeFormat.Xlsx, out NativeWriterHandle? handle));
                handle!.StartSheet("S");
                Assert.Throws<InvalidOperationException>(handle.EndRow);
                Assert.Equal(NativeStatus.Ok, NativeApi.CloseWriteHandle(handle));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void EndSheet_Should_Throw_Without_A_Matching_StartSheet()
        {
            string path = TempPath("xlsx");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenWriteHandle(path, NativeFormat.Xlsx, out NativeWriterHandle? handle));
                Assert.Throws<InvalidOperationException>(() => handle!.EndSheet());

                handle!.StartSheet("S");
                Assert.Equal(NativeStatus.Ok, NativeApi.CloseWriteHandle(handle));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void EndSheet_Should_Throw_While_A_Row_Is_Still_Open()
        {
            string path = TempPath("xlsx");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenWriteHandle(path, NativeFormat.Xlsx, out NativeWriterHandle? handle));
                handle!.StartSheet("S");
                handle.StartRow();
                Assert.Throws<InvalidOperationException>(handle.EndSheet);
                handle.EndRow();
                handle.EndSheet();
                Assert.Equal(NativeStatus.Ok, NativeApi.CloseWriteHandle(handle));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void CloseWriteHandle_Should_Produce_A_Valid_File_Even_When_EndSheet_Was_Never_Called()
        {
            string path = TempPath("xlsx");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenWriteHandle(path, NativeFormat.Xlsx, out NativeWriterHandle? handle));
                handle!.StartSheet("S");
                handle.StartRow();
                handle.WriteString("pending");

                Assert.Equal(NativeStatus.Ok, NativeApi.CloseWriteHandle(handle));

                Assert.Equal(NativeStatus.Ok, NativeApi.OpenFile(Encoding.UTF8.GetBytes(path), NativeFormat.Xlsx, out NativeHandle? reader));
                NativeApi.Close(reader);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void CloseWriteHandle_Should_Return_InvalidHandle_On_A_Null_Handle()
        {
            Assert.Equal(NativeStatus.InvalidHandle, NativeApi.CloseWriteHandle(null));
        }

        [Fact]
        public void CloseWriteHandle_Should_Not_Reuse_A_Handle_That_Was_Already_Closed()
        {
            string path = TempPath("xlsx");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenWriteHandle(path, NativeFormat.Xlsx, out NativeWriterHandle? handle));
                handle!.StartSheet("S");
                Assert.Equal(NativeStatus.Ok, NativeApi.CloseWriteHandle(handle));

                int status = NativeApi.CloseWriteHandle(handle);
                Assert.True(status is NativeStatus.Ok or NativeStatus.Error);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void StartSheet_Should_Throw_When_A_Sheet_Is_Already_Open()
        {
            string path = TempPath("xlsx");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenWriteHandle(path, NativeFormat.Xlsx, out NativeWriterHandle? handle));
                handle!.StartSheet("First");
                Assert.Throws<InvalidOperationException>(() => handle.StartSheet("Second"));
                Assert.Equal(NativeStatus.Ok, NativeApi.CloseWriteHandle(handle));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void StartRow_Should_Throw_When_A_Row_Is_Already_Open()
        {
            string path = TempPath("xlsx");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenWriteHandle(path, NativeFormat.Xlsx, out NativeWriterHandle? handle));
                handle!.StartSheet("S");
                handle.StartRow();
                Assert.Throws<InvalidOperationException>(handle.StartRow);
                handle.EndRow();
                Assert.Equal(NativeStatus.Ok, NativeApi.CloseWriteHandle(handle));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void OpenWriteHandle_Should_Carry_UseSharedStrings_Into_An_Xlsb_Writer()
        {
            string path = TempPath("xlsb");
            NativeWriteOptionsRaw raw = new()
            {
                StructSize = System.Runtime.InteropServices.Marshal.SizeOf<NativeWriteOptionsRaw>(),
                UseSharedStrings = NativeOptionState.True,
            };
            Assert.True(NativeWriteOptions.TryDecode(raw, null, out NativeWriteOptions options, out _));
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenWriteHandle(path, NativeFormat.Xlsb, out NativeWriterHandle? handle, options));
                handle!.StartSheet("S");
                Assert.Equal(NativeStatus.Ok, NativeApi.CloseWriteHandle(handle));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void NativeHandleTable_Should_Never_Resolve_A_Writer_Id_As_A_Reader_Or_Vice_Versa()
        {
            string xlsxFixture = Path.Combine(AppContext.BaseDirectory, "data", "sample.xlsx");
            string writePath = TempPath("xlsx");
            try
            {
                Assert.Equal(NativeStatus.Ok, NativeApi.OpenFile(Encoding.UTF8.GetBytes(xlsxFixture), NativeFormat.Auto, out NativeHandle? reader));
                nint readerId = NativeHandleTable.Register(reader!);

                Assert.Equal(NativeStatus.Ok, OpenWriteHandle(writePath, NativeFormat.Xlsx, out NativeWriterHandle? writer));
                nint writerId = NativeHandleTable.Register(writer!);

                Assert.Null(NativeHandleTable.Resolve<NativeWriterHandle>(readerId));
                Assert.Null(NativeHandleTable.Resolve<NativeHandle>(writerId));
                Assert.False(NativeHandleTable.TryUnregister(readerId, out NativeWriterHandle? _));
                Assert.False(NativeHandleTable.TryUnregister(writerId, out NativeHandle? _));

                Assert.Same(reader, NativeHandleTable.Resolve<NativeHandle>(readerId));
                Assert.Same(writer, NativeHandleTable.Resolve<NativeWriterHandle>(writerId));

                Assert.True(NativeHandleTable.TryUnregister(readerId, out NativeHandle? freedReader));
                NativeApi.Close(freedReader);
                writer!.StartSheet("S");
                Assert.Equal(NativeStatus.Ok, NativeApi.CloseWriteHandle(writer));
                NativeHandleTable.TryUnregister(writerId, out NativeWriterHandle? _);
            }
            finally
            {
                File.Delete(writePath);
            }
        }

        [Theory]
        [InlineData(NativeFormat.Xlsx)]
        [InlineData(NativeFormat.Xlsb)]
        [InlineData(NativeFormat.Xls)]
        [InlineData(NativeFormat.Csv)]
        public void GetWriteHandleBytes_Should_Round_Trip_A_Streamed_Row_Through_OpenMemory(int format)
        {
            Assert.Equal(NativeStatus.Ok, OpenWriteHandleToMemory(format, out NativeWriterHandle? handle));
            Assert.NotNull(handle);

            handle.StartSheet("Dados");
            handle.StartRow();
            handle.WriteString("uma");
            handle.WriteInt64(3);
            handle.EndRow();

            Assert.Equal(NativeStatus.Ok, NativeApi.GetWriteHandleBytes(handle, out byte[]? bytes));
            Assert.NotNull(bytes);
            Assert.NotEmpty(bytes);

            Assert.Equal(NativeStatus.Ok, NativeApi.OpenMemory(bytes, format, out NativeHandle? reader));
            Assert.NotNull(reader);
            try
            {
                Span<byte> buffer = stackalloc byte[256];
                Assert.Equal(NativeStatus.Ok, NativeApi.NextRow(reader, buffer, out int written));
                string row = Encoding.UTF8.GetString(buffer[..written]);
                Assert.Contains("uma", row, StringComparison.Ordinal);
                Assert.Contains("3", row, StringComparison.Ordinal);
            }
            finally
            {
                NativeApi.Close(reader);
            }

            Assert.Equal(NativeStatus.Ok, NativeApi.CloseWriteHandle(handle));
        }

        [Fact]
        public void GetWriteHandleBytes_Should_Not_Require_EndSheet_First()
        {
            Assert.Equal(NativeStatus.Ok, OpenWriteHandleToMemory(NativeFormat.Xlsx, out NativeWriterHandle? handle));
            handle!.StartSheet("S");
            handle.StartRow();
            handle.WriteString("pending");

            Assert.Equal(NativeStatus.Ok, NativeApi.GetWriteHandleBytes(handle, out byte[]? bytes));
            Assert.NotNull(bytes);

            Assert.Equal(NativeStatus.Ok, NativeApi.OpenMemory(bytes, NativeFormat.Xlsx, out NativeHandle? reader));
            NativeApi.Close(reader);

            Assert.Equal(NativeStatus.Ok, NativeApi.CloseWriteHandle(handle));
        }

        [Fact]
        public void GetWriteHandleBytes_Called_Twice_Should_Return_The_Same_Content()
        {
            Assert.Equal(NativeStatus.Ok, OpenWriteHandleToMemory(NativeFormat.Xlsx, out NativeWriterHandle? handle));
            handle!.StartSheet("S");

            Assert.Equal(NativeStatus.Ok, NativeApi.GetWriteHandleBytes(handle, out byte[]? first));
            Assert.Equal(NativeStatus.Ok, NativeApi.GetWriteHandleBytes(handle, out byte[]? second));
            Assert.NotNull(first);
            Assert.Equal(first, second);

            Assert.Equal(NativeStatus.Ok, NativeApi.CloseWriteHandle(handle));
        }

        [Fact]
        public void GetWriteHandleBytes_Should_Reject_A_File_Backed_Handle()
        {
            string path = TempPath("xlsx");
            try
            {
                Assert.Equal(NativeStatus.Ok, OpenWriteHandle(path, NativeFormat.Xlsx, out NativeWriterHandle? handle));
                handle!.StartSheet("S");

                Assert.Equal(NativeStatus.InvalidArgument, NativeApi.GetWriteHandleBytes(handle, out byte[]? bytes));
                Assert.Null(bytes);

                Assert.Equal(NativeStatus.Ok, NativeApi.CloseWriteHandle(handle));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void GetWriteHandleBytes_Should_Return_InvalidHandle_For_A_Null_Handle()
        {
            Assert.Equal(NativeStatus.InvalidHandle, NativeApi.GetWriteHandleBytes(null, out byte[]? bytes));
            Assert.Null(bytes);
        }

        [Fact]
        public void OpenWriteHandleToMemory_Should_Reject_Auto_Format()
        {
            Assert.Equal(NativeStatus.InvalidArgument, OpenWriteHandleToMemory(NativeFormat.Auto, out NativeWriterHandle? handle));
            Assert.Null(handle);
        }

        [Fact]
        public void CloseWriteHandle_After_GetWriteHandleBytes_Should_Not_Reopen_The_Workbook()
        {
            Assert.Equal(NativeStatus.Ok, OpenWriteHandleToMemory(NativeFormat.Xlsx, out NativeWriterHandle? handle));
            handle!.StartSheet("S");
            Assert.Equal(NativeStatus.Ok, NativeApi.GetWriteHandleBytes(handle, out _));
            Assert.Equal(NativeStatus.Ok, NativeApi.CloseWriteHandle(handle));
        }
    }
}
