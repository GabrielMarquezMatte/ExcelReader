using ExcelReader.Core.Writer.Csv;

namespace ExcelReader.Tests.Writer.Csv
{
    public class CsvWriterFailureTests
    {
        public static TheoryData<bool, bool> AsyncAndOwnership => new()
        {
            { false, false },
            { false, true },
            { true, false },
            { true, true },
        };

        private static void WriteRow(CsvWriter writer, string text)
        {
            using CsvRowWriter row = writer.StartRow();
            row.Write(text);
        }

        private static async Task WriteRowAsync(CsvWriter writer, string text)
        {
            await using CsvRowWriter row = writer.StartRow();
            row.Write(text);
        }

        private static Task DisposeWriter(CsvWriter writer, bool async)
        {
            if (async)
            {
                return writer.DisposeAsync().AsTask();
            }
            writer.Dispose();
            return Task.CompletedTask;
        }

        [Theory]
        [MemberData(nameof(AsyncAndOwnership))]
        public async Task Should_ReleaseBufferAndOwnedStream_When_TheFinalFlushFails(bool async, bool leaveOpen)
        {
            using var pool = CountingPool.Install();
            var stream = new TrackingStream();
            CsvWriter writer = CsvWriter.Create(stream, leaveOpen);
            WriteRow(writer, "buffered");
            stream.OnWrite = static () => throw new IOException("disk full");

            await Assert.ThrowsAsync<IOException>(() => DisposeWriter(writer, async));

            Assert.Equal(async ? 1 : 0, stream.AsyncWriteCount);
            Assert.Equal(leaveOpen ? 0 : 1, stream.DisposeCount);
            pool.AssertEveryBufferReturnedOnce();
            Assert.Null(await Record.ExceptionAsync(() => DisposeWriter(writer, async)));
            Assert.Equal(leaveOpen ? 0 : 1, stream.DisposeCount);
            Assert.Throws<ObjectDisposedException>(() => writer.StartRow());
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task Should_ReportTheFlushFailure_When_ClosingTheOwnedStreamAlsoFails(bool async)
        {
            using var pool = CountingPool.Install();
            var stream = new TrackingStream();
            CsvWriter writer = CsvWriter.Create(stream, leaveOpen: false);
            WriteRow(writer, "buffered");
            stream.OnWrite = static () => throw new IOException("disk full");
            stream.OnDispose = static () => throw new InvalidOperationException("close failed");

            await Assert.ThrowsAsync<IOException>(() => DisposeWriter(writer, async));

            Assert.Equal(1, stream.DisposeCount);
            pool.AssertEveryBufferReturnedOnce();
        }

        [Theory]
        [MemberData(nameof(AsyncAndOwnership))]
        public async Task Should_FaultTheWriter_When_ARowFlushFails(bool async, bool leaveOpen)
        {
            using var pool = CountingPool.Install();
            var stream = new TrackingStream();
            CsvWriter writer = CsvWriter.Create(stream, leaveOpen);
            stream.OnWrite = static () => throw new IOException("disk full");

            Exception? rowFailure = null;
            for (int i = 0; i < 100_000 && rowFailure is null; i++)
            {
                string text = $"row {i} with enough text to cross the flush threshold";
                rowFailure = async
                    ? await Record.ExceptionAsync(() => WriteRowAsync(writer, text))
                    : Record.Exception(() => WriteRow(writer, text));
            }
            Assert.IsType<IOException>(rowFailure);

            stream.OnWrite = null;
            Assert.Throws<ObjectDisposedException>(() => writer.StartRow());
            Assert.Throws<ObjectDisposedException>(writer.Flush);
            Assert.Null(await Record.ExceptionAsync(() => DisposeWriter(writer, async)));

            Assert.Empty(stream.ToArray());
            Assert.Equal(leaveOpen ? 0 : 1, stream.DisposeCount);
            pool.AssertEveryBufferReturnedOnce();
        }

        [Theory]
        [MemberData(nameof(AsyncAndOwnership))]
        public async Task Should_StayEndedAndClean_When_TheWorkbookEndFails(bool async, bool leaveOpen)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using var pool = CountingPool.Install();
            var stream = new TrackingStream();
            CsvWorkbookWriter workbook = CsvWorkbookWriter.Create(stream, leaveOpen);
            using (CsvRowWriter row = workbook.AddSheet("S1").StartRow())
            {
                row.Write("value");
            }
            stream.OnFlush = static () => throw new IOException("flush failed");

            if (async)
            {
                await Assert.ThrowsAsync<IOException>(() => workbook.EndAsync(ct).AsTask());
                await Assert.ThrowsAsync<ObjectDisposedException>(() => workbook.EndAsync(ct).AsTask());
                await workbook.DisposeAsync();
                await workbook.DisposeAsync();
            }
            else
            {
                Assert.Throws<IOException>(workbook.End);
                Assert.Throws<ObjectDisposedException>(workbook.End);
                workbook.Dispose();
                workbook.Dispose();
            }

            Assert.Throws<ObjectDisposedException>(() => workbook.AddSheet("S2"));
            Assert.Equal(leaveOpen ? 0 : 1, stream.DisposeCount);
            pool.AssertEveryBufferReturnedOnce();
        }

        [Theory]
        [MemberData(nameof(AsyncAndOwnership))]
        public async Task Should_FlushReleaseAndCloseOnce_When_DisposedSuccessfully(bool async, bool leaveOpen)
        {
            using var pool = CountingPool.Install();
            var stream = new TrackingStream();
            CsvWriter writer = CsvWriter.Create(stream, leaveOpen);
            WriteRow(writer, "a");
            _ = writer.StartRow();

            await DisposeWriter(writer, async);
            await DisposeWriter(writer, async);

            Assert.Equal("a\r\n\r\n", System.Text.Encoding.UTF8.GetString(stream.ToArray()));
            Assert.Equal(leaveOpen ? 0 : 1, stream.DisposeCount);
            pool.AssertEveryBufferReturnedOnce();
        }
    }
}
