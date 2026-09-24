using System.Globalization;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using ExcelReader.Core.Reader;
using ExcelReader.Native;
using ExcelReader.Native.Reading;
using ExcelReader.Tests.Crypto;

namespace ExcelReader.Tests.Native
{
    public sealed partial class NativeApiTests
    {
        private static readonly string XlsxFixture = Path.Combine(AppContext.BaseDirectory, "data", "sample.xlsx");

        internal static int OpenPath(string path, int format, out NativeHandle? handle)
        {
            return ReadApi.OpenFile(Encoding.UTF8.GetBytes(path), format, out handle);
        }

        private static NativeOpenOptionsRaw DefaultRawOptions()
        {
            return new NativeOpenOptionsRaw { StructSize = Marshal.SizeOf<NativeOpenOptionsRaw>() };
        }

        private static NativeTable SingleColumnTable(NativeColumn column, long rowCount)
        {
            IntPtr columns = Marshal.AllocHGlobal(Marshal.SizeOf<NativeColumn>());
            Marshal.StructureToPtr(column, columns, false);
            return new NativeTable { ColumnCount = 1, RowCount = rowCount, Columns = columns };
        }

        private static NativeColumn ColumnAt(NativeTable table, int index)
        {
            int columnSize = Marshal.SizeOf<NativeColumn>();
            return Marshal.PtrToStructure<NativeColumn>(IntPtr.Add(table.Columns, index * columnSize));
        }

        private static bool[] DecodeValidity(NativeColumn column)
        {
            int rowCount = (int)column.Length;
            bool[] result = new bool[rowCount];
            if (column.Validity == IntPtr.Zero)
            {
                Array.Fill(result, true);
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

        private static (string? Name, int Index, int Type, bool Nullable)[] DecodeSchema(NativeInferredSchema schema)
        {
            int specSize = Marshal.SizeOf<NativeColumnSpecRaw>();
            int namesOffset = (int)Marshal.OffsetOf<NativeColumnSpecRaw>(nameof(NativeColumnSpecRaw.Names));
            int nameLensOffset = (int)Marshal.OffsetOf<NativeColumnSpecRaw>(nameof(NativeColumnSpecRaw.NameLens));
            int nameCountOffset = (int)Marshal.OffsetOf<NativeColumnSpecRaw>(nameof(NativeColumnSpecRaw.NameCount));
            int indexOffset = (int)Marshal.OffsetOf<NativeColumnSpecRaw>(nameof(NativeColumnSpecRaw.Index));
            int typeOffset = (int)Marshal.OffsetOf<NativeColumnSpecRaw>(nameof(NativeColumnSpecRaw.Type));
            int nullableOffset = (int)Marshal.OffsetOf<NativeColumnSpecRaw>(nameof(NativeColumnSpecRaw.Nullable));

            var columns = new (string?, int, int, bool)[schema.ColumnCount];
            for (int i = 0; i < columns.Length; i++)
            {
                IntPtr spec = IntPtr.Add(schema.Columns, i * specSize);
                int nameCount = Marshal.ReadInt32(spec, nameCountOffset);
                string? name = null;
                if (nameCount > 0)
                {
                    IntPtr namesArray = Marshal.ReadIntPtr(spec, namesOffset);
                    IntPtr nameLensArray = Marshal.ReadIntPtr(spec, nameLensOffset);
                    IntPtr namePtr = Marshal.ReadIntPtr(namesArray, 0);
                    int nameLen = Marshal.ReadInt32(nameLensArray, 0);
                    name = Marshal.PtrToStringUTF8(namePtr, nameLen);
                }
                int index = Marshal.ReadInt32(spec, indexOffset);
                int type = Marshal.ReadInt32(spec, typeOffset);
                int nullable = Marshal.ReadInt32(spec, nullableOffset);
                columns[i] = (name, index, type, nullable != 0);
            }
            return columns;
        }

        private static string ReadZipEntry(ZipArchive archive, string entryName)
        {
            ZipArchiveEntry entry = Assert.IsType<ZipArchiveEntry>(archive.GetEntry(entryName));
            using StreamReader reader = new(entry.Open(), Encoding.UTF8);
            return reader.ReadToEnd();
        }

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

            public void Dispose()
            {
                inner.Dispose();
            }

            public ValueTask DisposeAsync()
            {
                return inner.DisposeAsync();
            }
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
        public void LastErrorPtr_Should_Survive_A_Gen2_Collection()
        {
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
            string?[] seen = new string?[2];
            using var firstSet = new ManualResetEventSlim();

            var first = new Thread(() =>
            {
                NativeApi.SetLastError("thread-one");
                firstSet.Set();
                SpinWait.SpinUntil(() => Volatile.Read(ref seen[1]) is not null, TimeSpan.FromSeconds(10));
                seen[0] = ReadLastErrorPtr();
            });
            var second = new Thread(() =>
            {
                firstSet.Wait();
                NativeApi.SetLastError("thread-two-error");
                seen[1] = ReadLastErrorPtr();
            });

            first.Start();
            second.Start();
            first.Join();
            second.Join();

            Assert.Equal("thread-one", seen[0]);
            Assert.Equal("thread-two-error", seen[1]);
        }

        [Fact]
        public void LastError_Should_Not_Keep_Memory_Pinned_After_Its_Thread_Exits()
        {
            const int threads = 64;
            long before = PinnedObjectsAfterFullCollection();

            for (int i = 0; i < threads; i++)
            {
                var thread = new Thread(static () => NativeApi.SetLastError("boom"));
                thread.Start();
                thread.Join();
            }

            long leaked = PinnedObjectsAfterFullCollection() - before;
            Assert.True(leaked < threads / 2, $"{leaked} pinned objects outlived the {threads} threads that set an error.");
        }

        private static string ReadLastErrorPtr()
        {
            nint pointer = NativeApi.LastErrorPtr(out int length);
            byte[] bytes = new byte[length];
            Marshal.Copy(pointer, bytes, 0, length);
            return Encoding.UTF8.GetString(bytes);
        }

        private static long PinnedObjectsAfterFullCollection()
        {
            GC.Collect(2, GCCollectionMode.Forced, blocking: true);
            GC.WaitForPendingFinalizers();
            GC.Collect(2, GCCollectionMode.Forced, blocking: true);
            return GC.GetGCMemoryInfo(GCKind.FullBlocking).PinnedObjectsCount;
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
        public void OpenFileEx_With_An_Unrecognized_Struct_Size_Is_Invalid_Argument()
        {
            NativeOpenOptionsRaw options = DefaultRawOptions() with { StructSize = 1 };

            int status = ReadApi.OpenFileEx(Encoding.UTF8.GetBytes(XlsxFixture), NativeFormat.Auto, options, out NativeHandle? handle);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Null(handle);
            Span<byte> buffer = stackalloc byte[256];
            Assert.Equal(NativeStatus.Ok, NativeApi.LastError(buffer, out int length));
            Assert.Contains("struct_size", Encoding.UTF8.GetString(buffer[..length]), StringComparison.Ordinal);
        }

        [Fact]
        public void EncryptPackage_Should_Produce_A_File_Openable_With_The_Same_Password()
        {
            string plainPath = EncryptedFixtures.PlainPath("agile-aes256-sha512.xlsx");
            string encryptedPath = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.xlsx");
            try
            {
                int status = NativeApi.EncryptPackage(
                    Encoding.UTF8.GetBytes(plainPath),
                    Encoding.UTF8.GetBytes(encryptedPath),
                    Encoding.UTF8.GetBytes(EncryptedFixtures.Password));

                Assert.Equal(NativeStatus.Ok, status);

                using IExcelRowReader plain = Excel.Open(plainPath);
                using IExcelRowReader roundTripped = Excel.Open(
                    encryptedPath, new ExcelReaderOptions { Password = EncryptedFixtures.Password });
                using IExcelRowEnumerator plainRows = plain.GetEnumerator();
                using IExcelRowEnumerator roundTrippedRows = roundTripped.GetEnumerator();

                Assert.True(plainRows.MoveNext());
                Assert.True(roundTrippedRows.MoveNext());
                Assert.Equal(plainRows.Current[0].GetString(), roundTrippedRows.Current[0].GetString());
            }
            finally
            {
                File.Delete(encryptedPath);
            }
        }

        [Fact]
        public void EncryptPackage_Should_Fail_When_The_Password_Is_Empty()
        {
            string encryptedPath = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.xlsx");
            try
            {
                int status = NativeApi.EncryptPackage(
                    Encoding.UTF8.GetBytes(EncryptedFixtures.PlainPath("agile-aes256-sha512.xlsx")),
                    Encoding.UTF8.GetBytes(encryptedPath),
                    ReadOnlySpan<byte>.Empty);

                Assert.Equal(NativeStatus.Error, status);
            }
            finally
            {
                File.Delete(encryptedPath);
            }
        }

        [Fact]
        public void EncryptPackage_Should_Reject_A_Password_Past_The_Length_Ceiling()
        {
            byte[] tooLong = new byte[4097];
            int status = NativeApi.EncryptPackage(
                Encoding.UTF8.GetBytes(EncryptedFixtures.PlainPath("agile-aes256-sha512.xlsx")),
                Encoding.UTF8.GetBytes(Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.xlsx")),
                tooLong);

            Assert.Equal(NativeStatus.InvalidArgument, status);
        }

        [Fact]
        public void Should_Report_The_Abi_Version_Declared_In_The_C_Header()
        {
            DirectoryInfo? dir = new(AppContext.BaseDirectory);
            while (dir is not null && !File.Exists(Path.Combine(dir.FullName, "ExcelReader.slnx")))
            {
                dir = dir.Parent;
            }
            Assert.NotNull(dir);

            string headerPath = Path.Combine(dir.FullName, "src", "ExcelReader.Native", "include", "excelreader.h");
            string header = File.ReadAllText(headerPath);
            Match match = Regex.Match(header, @"#define\s+XL_ABI_VERSION\s+(?<version>\d+)", RegexOptions.None, TimeSpan.FromSeconds(1));
            Assert.True(match.Success, "XL_ABI_VERSION not found in excelreader.h");

            Assert.Equal(NativeStatus.AbiVersion, int.Parse(match.Groups["version"].Value, CultureInfo.InvariantCulture));
        }
    }
}
