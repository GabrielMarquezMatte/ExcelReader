using System.Text;
using System.Threading.Channels;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Tests
{
    public class WriteOffloadTests
    {
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

        [Fact]
        public void AConsumerFailureSurfacesOnTheProducerWithItsOriginalType()
        {
            var stream = new WriteOffloadStream(new ThrowingStream());
            stream.Write("boom"u8);

            Assert.Throws<TimeoutException>(stream.Flush);
            Assert.Throws<TimeoutException>(stream.Dispose);
        }

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

            public override int Read(byte[] buffer, int offset, int count)
            {
                throw new NotSupportedException();
            }

            public override long Seek(long offset, SeekOrigin origin)
            {
                throw new NotSupportedException();
            }

            public override void SetLength(long value)
            {
                throw new NotSupportedException();
            }

            public override void Write(byte[] buffer, int offset, int count)
            {
                throw new TimeoutException("offload consumer failed");
            }

            public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
            {
                throw new TimeoutException("offload consumer failed");
            }
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
