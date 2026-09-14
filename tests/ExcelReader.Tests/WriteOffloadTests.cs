using System.Text;
using System.Threading.Channels;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Tests
{
    // WriteOffloadStream overlaps a ZIP entry's deflate with row-serialization on a background thread
    // (the write-side mirror of PrefetchDecompressionTests). Each test here pins one specific way that
    // can go wrong: a cell-content mismatch versus the non-offloaded path, or a sheet too small to ever
    // engage the offload path at all (XLSB only spills past SpillThreshold).
    //
    // Deliberately NOT a raw byte-for-byte comparison of the two builds' zip output: ZipArchiveEntry
    // stamps LastWriteTime from DateTime.Now at CreateEntry time, so two independent builds a few
    // milliseconds apart can legitimately cross a 2-second DOS-time rounding boundary and differ in a
    // header byte or two — a real, pre-existing quirk of System.IO.Compression, not something
    // WriteOffloadStream introduces. Comparing decoded cell content (as PrefetchDecompressionTests does
    // on the read side) is the correctness property that actually matters here.
    public class WriteOffloadTests
    {
        // Enough rows that the XLSX row buffer crosses its 64 KiB flush threshold and the XLSB record
        // buffer crosses its 64 KiB spill threshold many times over, not just once.
        private const int Rows = 20_000;

        [Fact]
        public async Task XlsxPrefetchWriteReadsBackIdenticalCells()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] bytes = await BuildXlsxAsync(prefetchWrite: true, ct);

            using MemoryStream ms = new(bytes, writable: false);
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();
            int rowIndex = 0;
            while (e.MoveNext())
            {
                Row row = e.Current;
                Assert.Equal(rowIndex.ToString(System.Globalization.CultureInfo.InvariantCulture), row[0].GetString());
                Assert.Equal($"row {rowIndex} text", row[1].GetString());
                rowIndex++;
            }
            Assert.Equal(Rows, rowIndex);
        }

        [Fact]
        public async Task XlsbPrefetchWriteReadsBackIdenticalCells()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] bytes = await BuildXlsbAsync(prefetchWrite: true, ct);

            using MemoryStream ms = new(bytes, writable: false);
            using XlsbReader reader = Excel.FromXlsb(ms);
            using XlsbReader.Enumerator e = reader.GetEnumerator();
            int rowIndex = 0;
            while (e.MoveNext())
            {
                Row row = e.Current;
                Assert.True(row[0].TryGetDouble(out double id));
                Assert.Equal(rowIndex, (int)id);
                Assert.Equal($"row {rowIndex} text", row[1].GetString());
                rowIndex++;
            }
            Assert.Equal(Rows, rowIndex);
        }

        // Uses the sync row-write API (StartRow/RowWriter.Dispose), which is the path that exercises
        // WriteOffloadStream's synchronous Write(ReadOnlySpan<byte>) override rather than WriteAsync.
        [Fact]
        public async Task XlsxPrefetchWriteWorksThroughSyncRowApi()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using MemoryStream ms = new();
            await using (XlsxWorkbookWriter wb = await XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true, prefetchWrite: true, ct: ct))
            {
                await wb.StartAsync(ct);
                XlsxSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(ct);
                for (int r = 0; r < Rows; r++)
                {
                    using XlsxRowWriter row = sheet.StartRow(ct);
                    row.Write(r);
                    row.Write($"row {r} text");
                }
                await sheet.EndAsync(ct);
            }

            ms.Position = 0;
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();
            int rowIndex = 0;
            while (e.MoveNext())
            {
                Row row = e.Current;
                Assert.Equal(rowIndex.ToString(System.Globalization.CultureInfo.InvariantCulture), row[0].GetString());
                rowIndex++;
            }
            Assert.Equal(Rows, rowIndex);
        }

        // XLSB's offload path only ever engages once the record buffer actually spills past
        // SpillThreshold (see XlsbSheetWriter.EnsureStream) — a small sheet must still round-trip
        // correctly with prefetchWrite: true even though WriteOffloadStream is never constructed.
        [Fact]
        public async Task XlsbPrefetchWriteRoundTripsASmallSheetThatNeverSpills()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using MemoryStream ms = new();
            await using (XlsbWorkbookWriter wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true, prefetchWrite: true, ct: ct))
            {
                await wb.StartAsync(ct);
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(ct);
                for (int r = 0; r < 5; r++)
                {
                    await using XlsbRowWriter row = await sheet.StartRowAsync(ct);
                    row.Write(r);
                    row.Write($"row {r}");
                }
                await sheet.EndAsync(ct);
            }

            ms.Position = 0;
            using XlsbReader reader = Excel.FromXlsb(ms);
            using XlsbReader.Enumerator e = reader.GetEnumerator();
            int rowIndex = 0;
            while (e.MoveNext())
            {
                Assert.Equal($"row {rowIndex}", e.Current[1].GetString());
                rowIndex++;
            }
            Assert.Equal(5, rowIndex);
        }

        [Fact]
        public void TheStreamContractRejectsEveryReadAndSeekOperation()
        {
            using MemoryStream inner = new();
            using var stream = new WriteOffloadStream(inner);

            Assert.False(stream.CanRead);
            Assert.False(stream.CanSeek);
            Assert.True(stream.CanWrite);
            Assert.Throws<NotSupportedException>(() => _ = stream.Length);
            Assert.Throws<NotSupportedException>(() => _ = stream.Position);
            Assert.Throws<NotSupportedException>(() => stream.Position = 0);
            Assert.Throws<NotSupportedException>(() => stream.Read(new byte[1], 0, 1));
            Assert.Throws<NotSupportedException>(() => stream.Seek(0, SeekOrigin.Begin));
            Assert.Throws<NotSupportedException>(() => stream.SetLength(0));
        }

        [Fact]
        public void SyncWritesReachTheInnerStreamInOrder()
        {
            MemoryStream inner = new();
            using (var stream = new WriteOffloadStream(inner))
            {
                stream.Write([], 0, 0);
                stream.Write("hello "u8.ToArray(), 0, 6);
                stream.Write("world"u8);
                stream.Flush();
            }

            Assert.Equal("hello world", Encoding.UTF8.GetString(inner.ToArray()));
        }

        [Fact]
        public async Task AsyncWritesReachTheInnerStreamInOrder()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            MemoryStream inner = new();
            await using (var stream = new WriteOffloadStream(inner))
            {
                await stream.WriteAsync([], 0, 0, ct);
                await stream.WriteAsync("hello "u8.ToArray(), 0, 6, ct);
                await stream.WriteAsync("world"u8.ToArray().AsMemory(), ct);
                await stream.FlushAsync(ct);
            }

            Assert.Equal("hello world", Encoding.UTF8.GetString(inner.ToArray()));
        }

        // EnqueueOwned(Async) takes over a BiffBuffer.Detach array instead of copying. A zero-length
        // one still has to go back to BiffBuffer's own pool rather than being dropped or, worse,
        // returned to ArrayPool<byte>.Shared.
        [Fact]
        public async Task EnqueueOwnedWritesTheDetachedBufferAndReturnsAnEmptyOneUnwritten()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            MemoryStream inner = new();
            await using (var stream = new WriteOffloadStream(inner))
            {
                stream.EnqueueOwned(Detach("sync "u8, out int syncLength), syncLength);
                stream.EnqueueOwned(Detach([], out _), 0);
                await stream.EnqueueOwnedAsync(Detach("async"u8, out int asyncLength), asyncLength, ct);
                await stream.EnqueueOwnedAsync(Detach([], out _), 0, ct);
                await stream.FlushAsync(ct);
            }

            Assert.Equal("sync async", Encoding.UTF8.GetString(inner.ToArray()));
        }

        // The consumer runs on a pool thread and swallows its own exceptions so the task never faults
        // unobserved; the producer must still see the original exception, with its original type, on
        // the next call that synchronizes with the consumer.
        [Fact]
        public void AConsumerFailureSurfacesOnTheProducerWithItsOriginalType()
        {
            var stream = new WriteOffloadStream(new ThrowingStream());
            stream.Write("boom"u8);

            Assert.Throws<TimeoutException>(stream.Flush);
            Assert.Throws<TimeoutException>(stream.Dispose);
        }

        // Flush completes the channel writer, so anything enqueued afterwards hits a closed channel.
        // Each enqueue path has its own catch that must hand the buffer back before rethrowing.
        [Fact]
        public async Task EveryEnqueuePathReportsAClosedChannelAfterFlush()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            MemoryStream inner = new();
            var stream = new WriteOffloadStream(inner);
            stream.Write("a"u8);
            await stream.FlushAsync(ct);

            Assert.Throws<ChannelClosedException>(() => stream.Write("b"u8));
            Assert.Throws<ChannelClosedException>(() => stream.EnqueueOwned(Detach("c"u8, out int owned), owned));
            await Assert.ThrowsAsync<ChannelClosedException>(async () =>
                await stream.WriteAsync("d"u8.ToArray().AsMemory(), ct));
            await Assert.ThrowsAsync<ChannelClosedException>(async () =>
                await stream.EnqueueOwnedAsync(Detach("e"u8, out int ownedAsync), ownedAsync, ct));

            stream.Dispose();
            Assert.Equal("a", Encoding.UTF8.GetString(inner.ToArray()));
        }

        [Fact]
        public async Task DisposeAsyncAfterDisposeDoesNotRunTeardownTwice()
        {
            MemoryStream inner = new();
            var stream = new WriteOffloadStream(inner);
            stream.Write("x"u8);
            stream.Dispose();

            await stream.DisposeAsync();

            Assert.Equal("x", Encoding.UTF8.GetString(inner.ToArray()));
        }

        private static byte[] Detach(ReadOnlySpan<byte> content, out int length)
        {
            using BiffBuffer buffer = new();
            buffer.Write(content);
            return buffer.Detach(out length);
        }

        private sealed class ThrowingStream : Stream
        {
            public override bool CanRead => false;

            public override bool CanSeek => false;

            public override bool CanWrite => true;

            public override long Length => throw new NotSupportedException();

            public override long Position
            {
                get => throw new NotSupportedException(); set => throw new NotSupportedException();
            }

            public override void Flush()
            {
            }

            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

            public override void SetLength(long value) => throw new NotSupportedException();

            public override void Write(byte[] buffer, int offset, int count) => throw new TimeoutException("offload consumer failed");

            public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
                => throw new TimeoutException("offload consumer failed");
        }

        private static async Task<byte[]> BuildXlsxAsync(bool prefetchWrite, CancellationToken ct)
        {
            MemoryStream ms = new();
            await using (XlsxWorkbookWriter wb = await XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true, prefetchWrite: prefetchWrite, ct: ct))
            {
                await wb.StartAsync(ct);
                XlsxSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(ct);
                for (int r = 0; r < Rows; r++)
                {
                    await using XlsxRowWriter row = sheet.StartRow(ct);
                    row.Write(r);
                    row.Write($"row {r} text");
                }
                await sheet.EndAsync(ct);
            }
            return ms.ToArray();
        }

        private static async Task<byte[]> BuildXlsbAsync(bool prefetchWrite, CancellationToken ct)
        {
            MemoryStream ms = new();
            await using (XlsbWorkbookWriter wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true, prefetchWrite: prefetchWrite, ct: ct))
            {
                await wb.StartAsync(ct);
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(ct);
                for (int r = 0; r < Rows; r++)
                {
                    await using XlsbRowWriter row = await sheet.StartRowAsync(ct);
                    row.Write(r);
                    row.Write($"row {r} text");
                }
                await sheet.EndAsync(ct);
            }
            return ms.ToArray();
        }
    }
}
