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

                // A workbook with zero sheets is not a file IWorkbookWriter.End can finalize, so the
                // handle needs a sheet before it can be closed successfully - this is the underlying
                // writer's own contract, not something xl_close_write_handle relaxes.
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

                // See StartRow_Should_Throw_Before_StartSheet: a sheet-less workbook can't be closed
                // successfully, so add one before asserting the close status.
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
            // The row/sheet are left open on purpose: xl_close_write_handle must still leave a
            // readable file behind, unlike EndRow/EndSheet's strict single-step guards above.
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

                // Close() a second time on the same, already-disposed handle: whatever the workbook
                // writer does with a repeat End()/Dispose() must not throw past NativeApi's own
                // try/catch and must not resurrect the file.
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
            // Regression guard for the bug where NativeWriterHandle.Create dropped
            // useSharedStrings for the XLSB branch: this only asserts the handle opens and closes
            // cleanly with the option set, since the shared-strings table itself is an internal
            // implementation detail XlsbWorkbookWriter doesn't expose for direct inspection.
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

        // Regression guard for the id-collision bug: NativeHandleTable used to be generic
        // (NativeHandleTable<THandle>), which gave reader ids and writer ids independent counters
        // that both start at 1 - a live reader and a live writer could then share the same id, and
        // an entry point for one kind would resolve the other kind's object instead of failing
        // cleanly. The shared, single-counter, runtime-type-checked table must keep that from ever
        // happening again, for handles that are simultaneously live AND for a retired id reused by
        // the other kind.
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

                // A live reader id resolved as a writer, and a live writer id resolved as a reader,
                // must both come back null - never the other kind's object.
                Assert.Null(NativeHandleTable.Resolve<NativeWriterHandle>(readerId));
                Assert.Null(NativeHandleTable.Resolve<NativeHandle>(writerId));
                Assert.False(NativeHandleTable.TryUnregister(readerId, out NativeWriterHandle? _));
                Assert.False(NativeHandleTable.TryUnregister(writerId, out NativeHandle? _));

                // The failed cross-type TryUnregister calls above must not have removed either
                // handle from the table - both must still resolve as their own, correct kind.
                Assert.Same(reader, NativeHandleTable.Resolve<NativeHandle>(readerId));
                Assert.Same(writer, NativeHandleTable.Resolve<NativeWriterHandle>(writerId));

                Assert.True(NativeHandleTable.TryUnregister(readerId, out NativeHandle? freedReader));
                NativeApi.Close(freedReader);
                writer!.StartSheet("S"); // a sheet-less workbook cannot be closed successfully
                Assert.Equal(NativeStatus.Ok, NativeApi.CloseWriteHandle(writer));
                NativeHandleTable.TryUnregister(writerId, out NativeWriterHandle? _); // already closed above; just drop the id
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

            // GetWriteHandleBytes must not have released the handle: xl_close_write_handle is still
            // required, same contract as a file-backed handle.
            Assert.Equal(NativeStatus.Ok, NativeApi.CloseWriteHandle(handle));
        }

        [Fact]
        public void GetWriteHandleBytes_Should_Not_Require_EndSheet_First()
        {
            // Same idempotent-Close contract CloseWriteHandle relies on (see
            // CloseWriteHandle_Should_Produce_A_Valid_File_Even_When_EndSheet_Was_Never_Called): the
            // row/sheet are left open on purpose.
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
            // Exercises Close()'s new idempotency guard directly: a second GetWriteHandleBytes call
            // must not re-run _workbook.End() and must not throw or corrupt the buffer.
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
            // GetWriteHandleBytes's internal Close() call must be indistinguishable, from
            // CloseWriteHandle's perspective, from having already been closed once - no exception, no
            // attempt to append more workbook structure.
            Assert.Equal(NativeStatus.Ok, OpenWriteHandleToMemory(NativeFormat.Xlsx, out NativeWriterHandle? handle));
            handle!.StartSheet("S");
            Assert.Equal(NativeStatus.Ok, NativeApi.GetWriteHandleBytes(handle, out _));
            Assert.Equal(NativeStatus.Ok, NativeApi.CloseWriteHandle(handle));
        }
    }
}
