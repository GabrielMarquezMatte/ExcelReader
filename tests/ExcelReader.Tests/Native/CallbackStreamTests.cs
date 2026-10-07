using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Native;
using ExcelReader.Native.Reading;
using ExcelReader.Tests.Reader;

namespace ExcelReader.Tests.Native
{
    public sealed unsafe class CallbackStreamTests
    {
        internal sealed class TestStream(byte[] bytes)
        {
            internal readonly byte[] Bytes = bytes;
            internal int Position;
            internal int Releases;
            internal string? Error;
        }

        private sealed class Pinned(TestStream stream) : IDisposable
        {
            private GCHandle _handle = GCHandle.Alloc(stream);

            internal NativeStreamRaw Raw()
            {
                return new()
                {
                    StructSize = sizeof(NativeStreamRaw),
                    UserData = (void*)GCHandle.ToIntPtr(_handle),
                    Read = &Read,
                    Release = &Release,
                };
            }

            public void Dispose()
            {
                _handle.Free();
            }
        }

        [UnmanagedCallersOnly]
        private static long Read(void* userData, byte* buffer, long length)
        {
            TestStream stream = (TestStream)GCHandle.FromIntPtr((nint)userData).Target!;
            if (stream.Error is string message)
            {
                SourceErrors.Set(message);
                return -1;
            }
            int count = (int)Math.Min(Math.Min(length, 7), stream.Bytes.Length - stream.Position);
            stream.Bytes.AsSpan(stream.Position, count).CopyTo(new Span<byte>(buffer, count));
            stream.Position += count;
            return count;
        }

        [UnmanagedCallersOnly]
        private static void Release(void* userData)
        {
            TestStream stream = (TestStream)GCHandle.FromIntPtr((nint)userData).Target!;
            Interlocked.Increment(ref stream.Releases);
        }

        private static int Open(Pinned pinned, int format, out NativeHandle? handle, NativeOpenOptionsRaw? options = null)
        {
            NativeStreamRaw raw = pinned.Raw();
            Assert.True(CallbackReadStream.TryCreate(&raw, hasOutHandle: true, out CallbackReadStream? stream, out string? error), error);
            return ReadApi.OpenStream(stream!, format, options, out handle);
        }

        private static NativeOpenOptionsRaw Options() => new() { StructSize = sizeof(NativeOpenOptionsRaw) };

        private static int CountRows(NativeHandle handle, int sheet)
        {
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenRows(handle, sheet, out NativeRowCursor? opened));
            using NativeRowCursor cursor = opened!;
            int rows = 0;
            while (ReadApi.NextRowView(cursor, out _) == NativeStatus.Ok)
            {
                rows++;
            }
            return rows;
        }

        [Fact]
        public void A_Csv_Streams_Once_And_Is_Released_When_Its_Cursor_Closes()
        {
            TestStream test = new(Encoding.UTF8.GetBytes("a,b\n1,2\n3,4\n"));
            using Pinned pinned = new(test);
            Assert.Equal(NativeStatus.Ok, Open(pinned, NativeFormat.Csv, out NativeHandle? handle));
            Assert.Equal(3, CountRows(handle!, 0));

            Assert.Equal(NativeStatus.Error, ReadApi.OpenRows(handle, 0, out _));
            Assert.Contains("only be enumerated once", NativeApi.LastErrorText(), StringComparison.Ordinal);

            Assert.Equal(0, test.Releases);
            Assert.Equal(NativeStatus.Ok, ReadApi.Close(handle));
            Assert.Equal(1, test.Releases);
        }

        [Theory]
        [InlineData(ConcurrentSheetTests.Format.Xlsx, NativeFormat.Xlsx)]
        [InlineData(ConcurrentSheetTests.Format.Xls, NativeFormat.Auto)]
        public void An_Excel_Workbook_Is_Buffered_And_The_Stream_Released_Before_The_Open_Returns(
            ConcurrentSheetTests.Format format, int nativeFormat)
        {
            TestStream test = new(ConcurrentSheetTests.Build(format));
            using Pinned pinned = new(test);
            Assert.Equal(NativeStatus.Ok, Open(pinned, nativeFormat, out NativeHandle? handle));
            Assert.Equal(1, test.Releases);
            using (handle)
            {
                Assert.Equal(ConcurrentSheetTests.Rows, CountRows(handle!, 5));
            }
            Assert.Equal(1, test.Releases);
        }

        [Fact]
        public void A_Stream_Past_The_Buffer_Limit_Fails_And_Is_Released_Once()
        {
            TestStream test = new(ConcurrentSheetTests.Build(ConcurrentSheetTests.Format.Xlsx));
            using Pinned pinned = new(test);
            Assert.Equal(NativeStatus.Error, Open(pinned, NativeFormat.Xlsx, out _, Options() with { MaxBufferedBytes = 1000 }));
            Assert.Contains("MaxBufferedBytes", NativeApi.LastErrorText(), StringComparison.Ordinal);
            Assert.Equal(1, test.Releases);
        }

        [Fact]
        public void Sniffing_A_Csv_Stream_Fails_With_A_Message_And_Is_Released_Once()
        {
            TestStream test = new(Encoding.UTF8.GetBytes("a;b\n1;2\n"));
            using Pinned pinned = new(test);
            Assert.Equal(NativeStatus.Error, Open(pinned, NativeFormat.Csv, out _, Options() with { CsvSniffDialect = 2 }));
            Assert.Contains("seekable", NativeApi.LastErrorText(), StringComparison.Ordinal);
            Assert.Equal(1, test.Releases);
        }

        [Fact]
        public void The_Callback_Message_Reaches_The_Caller()
        {
            TestStream test = new([1, 2, 3]) { Error = "pipe broke" };
            using Pinned pinned = new(test);
            Assert.Equal(NativeStatus.Error, Open(pinned, NativeFormat.Xlsx, out _));
            Assert.Contains("pipe broke", NativeApi.LastErrorText(), StringComparison.Ordinal);
            Assert.Equal(1, test.Releases);
        }

        [Fact]
        public void Invalid_Stream_Structs_Are_Rejected_Without_Taking_Ownership()
        {
            TestStream test = new([]);
            using Pinned pinned = new(test);
            NativeStreamRaw good = pinned.Raw();
            NativeStreamRaw noRead = good with { Read = null };
            NativeStreamRaw wrongSize = good with { StructSize = 4 };
            Assert.False(CallbackReadStream.TryCreate(null, true, out _, out _));
            Assert.False(CallbackReadStream.TryCreate(&good, false, out _, out _));
            Assert.False(CallbackReadStream.TryCreate(&noRead, true, out _, out _));
            Assert.False(CallbackReadStream.TryCreate(&wrongSize, true, out _, out string? error));
            Assert.Contains("xl_open_stream", error, StringComparison.Ordinal);
            Assert.Equal(0, test.Releases);
        }
    }
}
