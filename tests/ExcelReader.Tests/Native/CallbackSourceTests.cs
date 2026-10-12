using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Writer.Xlsx;
using ExcelReader.Native;
using ExcelReader.Native.Reading;
using ExcelReader.Tests.Reader;

namespace ExcelReader.Tests.Native
{
    public sealed unsafe class CallbackSourceTests
    {
        internal sealed class TestSource(byte[] bytes)
        {
            internal readonly byte[] Bytes = bytes;
            internal int Reads;
            internal int Releases;
            internal string? Error;
            internal long ShortAt = -1;
            internal int FailOffThread;
        }

        private sealed class Pinned(TestSource source) : IDisposable
        {
            private GCHandle _handle = GCHandle.Alloc(source);

            internal NativeSourceRaw Raw(long? length = null) => new()
            {
                StructSize = sizeof(NativeSourceRaw),
                UserData = (void*)GCHandle.ToIntPtr(_handle),
                Length = length ?? source.Bytes.Length,
                ReadAt = &ReadAt,
                Release = &Release,
            };

            public void Dispose()
            {
                _handle.Free();
            }
        }

        [UnmanagedCallersOnly]
        private static long ReadAt(void* userData, long offset, byte* buffer, long length)
        {
            TestSource source = (TestSource)GCHandle.FromIntPtr((nint)userData).Target!;
            Interlocked.Increment(ref source.Reads);
            if (source.Error is string message && Environment.CurrentManagedThreadId != source.FailOffThread)
            {
                SourceErrors.Set(message);
                return -1;
            }
            if (source.ShortAt >= 0 && offset >= source.ShortAt)
            {
                return 0;
            }
            long count = Math.Min(length, source.Bytes.Length - offset);
            source.Bytes.AsSpan((int)offset, (int)count).CopyTo(new Span<byte>(buffer, (int)count));
            return count;
        }

        [UnmanagedCallersOnly]
        private static void Release(void* userData)
        {
            TestSource source = (TestSource)GCHandle.FromIntPtr((nint)userData).Target!;
            Interlocked.Increment(ref source.Releases);
        }

        private static int Open(Pinned pinned, int format, out NativeHandle? handle, NativeOpenOptionsRaw? options = null, long? length = null)
        {
            NativeSourceRaw raw = pinned.Raw(length);
            Assert.True(CallbackByteSource.TryCreate(&raw, hasOutHandle: true, out CallbackByteSource? source, out string? error), error);
            return ReadApi.OpenSource(source!, format, options, out handle);
        }

        private static NativeOpenOptionsRaw Options() => new() { StructSize = sizeof(NativeOpenOptionsRaw) };

        private static byte[] BuildPrefetchedXlsx()
        {
            using MemoryStream buffer = new();
            using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(buffer, leaveOpen: true))
            {
                XlsxSheetWriter sheet = workbook.AddSheet("big");
                for (int r = 0; r < 20_000; r++)
                {
                    using XlsxRowWriter row = sheet.StartRow();
                    row.Write("row-" + r.ToString(CultureInfo.InvariantCulture));
                    row.Write(r);
                }
                sheet.End();
                sheet.Dispose();
                workbook.End();
            }
            return buffer.ToArray();
        }

        private static List<string> ReadSheet(NativeHandle handle, int sheet)
        {
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenRows(handle, sheet, out NativeRowCursor? opened));
            using NativeRowCursor cursor = opened!;
            List<string> lines = [];
            while (true)
            {
                int status = ReadApi.NextRowView(cursor, out NativeRow row);
                if (status == NativeStatus.Eof)
                {
                    return lines;
                }
                Assert.Equal(NativeStatus.Ok, status);
                StringBuilder line = new();
                int cellSize = Marshal.SizeOf<NativeRowCell>();
                for (int i = 0; i < row.CellCount; i++)
                {
                    NativeRowCell cell = Marshal.PtrToStructure<NativeRowCell>(IntPtr.Add(row.Cells, i * cellSize));
                    line.Append(Marshal.PtrToStringUTF8(cell.Value, cell.ValueLength)).Append('|');
                }
                lines.Add(line.ToString());
            }
        }

        [Theory]
        [InlineData(ConcurrentSheetTests.Format.Xlsx, NativeFormat.Xlsx)]
        [InlineData(ConcurrentSheetTests.Format.Xlsb, NativeFormat.Xlsb)]
        [InlineData(ConcurrentSheetTests.Format.Xls, NativeFormat.Xls)]
        [InlineData(ConcurrentSheetTests.Format.Xlsx, NativeFormat.Auto)]
        public void Every_Sheet_Read_In_Parallel_Matches_The_Memory_Read(ConcurrentSheetTests.Format format, int nativeFormat)
        {
            byte[] bytes = ConcurrentSheetTests.Build(format);
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(bytes, nativeFormat, out NativeHandle? memory));
            List<string>[] expected;
            using (memory)
            {
                expected = [.. Enumerable.Range(0, ConcurrentSheetTests.Sheets).Select(s => ReadSheet(memory!, s))];
            }

            TestSource test = new(bytes);
            using Pinned pinned = new(test);
            Assert.Equal(NativeStatus.Ok, Open(pinned, nativeFormat, out NativeHandle? handle));
            using (NativeHandle live = handle!)
            {
                List<string>[] parallel = new List<string>[expected.Length * 2];
                Parallel.For(0, parallel.Length, i => parallel[i] = ReadSheet(live, i % expected.Length));
                for (int i = 0; i < parallel.Length; i++)
                {
                    Assert.Equal(expected[i % expected.Length], parallel[i]);
                }
            }
            Assert.Equal(1, test.Releases);
        }

        [Fact]
        public void A_Csv_Is_Read_Through_The_Source()
        {
            TestSource test = new(Encoding.UTF8.GetBytes("a,b\n1,2\n"));
            using Pinned pinned = new(test);
            Assert.Equal(NativeStatus.Ok, Open(pinned, NativeFormat.Csv, out NativeHandle? handle));
            using (handle)
            {
                Assert.Equal(["a|b|", "1|2|"], ReadSheet(handle!, 0));
            }
            Assert.Equal(1, test.Releases);
        }

        [Fact]
        public void The_Source_Is_Released_Only_When_The_Last_Cursor_Closes_After_The_Workbook()
        {
            TestSource test = new(ConcurrentSheetTests.Build(ConcurrentSheetTests.Format.Xlsx));
            using Pinned pinned = new(test);
            Assert.Equal(NativeStatus.Ok, Open(pinned, NativeFormat.Xlsx, out NativeHandle? handle));
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenRows(handle, 0, out NativeRowCursor? cursor));
            Assert.Equal(NativeStatus.Ok, ReadApi.Close(handle));
            Assert.Equal(0, test.Releases);
            cursor!.Dispose();
            Assert.Equal(1, test.Releases);
        }

        [Fact]
        public void A_Failed_Open_Releases_The_Source_Once()
        {
            TestSource test = new(Encoding.UTF8.GetBytes("not a workbook at all"));
            using Pinned pinned = new(test);
            Assert.Equal(NativeStatus.Error, Open(pinned, NativeFormat.Auto, out NativeHandle? handle));
            Assert.Null(handle);
            Assert.Equal(1, test.Releases);
        }

        [Fact]
        public void An_Empty_Source_Fails_With_The_Format_Message_And_Is_Released()
        {
            TestSource test = new([]);
            using Pinned pinned = new(test);
            Assert.Equal(NativeStatus.Error, Open(pinned, NativeFormat.Auto, out _));
            Assert.Contains("Unrecognized file format", NativeApi.LastErrorText(), StringComparison.Ordinal);
            Assert.Equal(1, test.Releases);
        }

        [Fact]
        public void Invalid_Options_Release_The_Source_Once()
        {
            TestSource test = new(ConcurrentSheetTests.Build(ConcurrentSheetTests.Format.Xlsx));
            using Pinned pinned = new(test);
            NativeOpenOptionsRaw options = Options() with { MaxBufferedBytes = -1 };
            Assert.Equal(NativeStatus.InvalidArgument, Open(pinned, NativeFormat.Xlsx, out _, options));
            Assert.Equal(1, test.Releases);
        }

        [Fact]
        public void The_Callback_Message_Reaches_The_Caller()
        {
            TestSource test = new(ConcurrentSheetTests.Build(ConcurrentSheetTests.Format.Xlsx)) { Error = "disk on fire" };
            using Pinned pinned = new(test);
            Assert.Equal(NativeStatus.Error, Open(pinned, NativeFormat.Xlsx, out _));
            Assert.Contains("disk on fire", NativeApi.LastErrorText(), StringComparison.Ordinal);
            Assert.Equal(1, test.Releases);
        }

        [Fact]
        public void The_Callback_Message_Reaches_The_Caller_When_Prefetch_Reads_On_Another_Thread()
        {
            byte[] bytes = BuildPrefetchedXlsx();
            TestSource test = new(bytes);
            using Pinned pinned = new(test);
            NativeOpenOptionsRaw options = Options() with { PrefetchDecompression = 2, SourceBlockSize = -1 };
            Assert.Equal(NativeStatus.Ok, Open(pinned, NativeFormat.Xlsx, out NativeHandle? handle, options));
            using (handle)
            {
                test.FailOffThread = Environment.CurrentManagedThreadId;
                test.Error = "disk on fire";
                Assert.Equal(NativeStatus.Ok, ReadApi.OpenRows(handle, 0, out NativeRowCursor? opened));
                using NativeRowCursor cursor = opened!;
                int status;
                do
                {
                    status = ReadApi.NextRowView(cursor, out _);
                }
                while (status == NativeStatus.Ok);
                Assert.Equal(NativeStatus.Error, status);
                Assert.Contains("disk on fire", NativeApi.LastErrorText(), StringComparison.Ordinal);
            }
        }

        [Fact]
        public void A_Source_That_Ends_Early_Fails_Instead_Of_Looping()
        {
            byte[] bytes = ConcurrentSheetTests.Build(ConcurrentSheetTests.Format.Xlsx);
            TestSource test = new(bytes) { ShortAt = bytes.Length / 2 };
            using Pinned pinned = new(test);
            NativeOpenOptionsRaw options = Options() with { SourceBlockSize = -1 };
            Assert.Equal(NativeStatus.Error, Open(pinned, NativeFormat.Xlsx, out _, options));
            Assert.Contains("ended at offset", NativeApi.LastErrorText(), StringComparison.Ordinal);
        }

        [Fact]
        public void The_Default_Cache_Fetches_Each_Block_Once()
        {
            byte[] bytes = ConcurrentSheetTests.Build(ConcurrentSheetTests.Format.Xlsx);
            long blocks = (bytes.Length + (4L * 1024 * 1024) - 1) / (4L * 1024 * 1024);

            TestSource cached = new(bytes);
            using (Pinned pinned = new(cached))
            {
                Assert.Equal(NativeStatus.Ok, Open(pinned, NativeFormat.Xlsx, out NativeHandle? handle));
                using (handle)
                {
                    ReadSheet(handle!, 0);
                }
            }
            Assert.True(cached.Reads <= blocks, $"{cached.Reads} reads for {blocks} blocks");

            TestSource uncached = new(bytes);
            using (Pinned pinned = new(uncached))
            {
                Assert.Equal(NativeStatus.Ok, Open(pinned, NativeFormat.Xlsx, out NativeHandle? handle, Options() with { SourceBlockSize = -1 }));
                using (handle)
                {
                    ReadSheet(handle!, 0);
                }
            }
            Assert.True(uncached.Reads > blocks, $"{uncached.Reads} reads without the cache");
        }

        [Fact]
        public void Invalid_Source_Structs_Are_Rejected_Without_Taking_Ownership()
        {
            TestSource test = new([1, 2, 3]);
            using Pinned pinned = new(test);
            NativeSourceRaw good = pinned.Raw();
            NativeSourceRaw wrongSize = good with { StructSize = 8 };
            NativeSourceRaw noRead = good with { ReadAt = null };
            NativeSourceRaw negative = good with { Length = -1 };

            Assert.False(CallbackByteSource.TryCreate(null, true, out _, out _));
            Assert.False(CallbackByteSource.TryCreate(&good, false, out _, out _));
            Assert.False(CallbackByteSource.TryCreate(&wrongSize, true, out _, out _));
            Assert.False(CallbackByteSource.TryCreate(&noRead, true, out _, out _));
            Assert.False(CallbackByteSource.TryCreate(&negative, true, out _, out string? error));
            Assert.Contains("xl_open_source", error, StringComparison.Ordinal);
            Assert.Equal(0, test.Releases);
        }
    }
}
