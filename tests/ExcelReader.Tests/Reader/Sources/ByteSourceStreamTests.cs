using ExcelReader.Core.Reader.Sources;

namespace ExcelReader.Tests.Reader.Sources
{
    public sealed class ByteSourceStreamTests
    {
        private sealed class CountingSource(byte[] bytes) : ByteSource
        {
            internal int Reads { get; private set; }

            internal bool Disposed { get; private set; }

            internal override long Length => bytes.Length;

            internal override int Read(long offset, Span<byte> destination)
            {
                Reads++;
                if (offset >= bytes.Length)
                {
                    return 0;
                }
                int count = (int)Math.Min(destination.Length, bytes.Length - offset);
                bytes.AsSpan((int)offset, count).CopyTo(destination);
                return count;
            }

            internal override ValueTask<int> ReadAsync(long offset, Memory<byte> destination, CancellationToken ct)
            {
                return new ValueTask<int>(Read(offset, destination.Span));
            }

            public override void Dispose()
            {
                Disposed = true;
            }
        }

        private static byte[] Payload(int length)
        {
            byte[] bytes = new byte[length];
            for (int i = 0; i < length; i++)
            {
                bytes[i] = (byte)((i * 31) + (i >> 8));
            }
            return bytes;
        }

        [Fact]
        public void Reads_Exactly_Its_Range()
        {
            byte[] bytes = Payload(300_000);
            using CountingSource source = new(bytes);
            using ByteSourceStream stream = new(source, offset: 1000, length: 200_000);
            using MemoryStream copy = new();
            stream.CopyTo(copy);
            Assert.True(copy.ToArray().AsSpan().SequenceEqual(bytes.AsSpan(1000, 200_000)));
        }

        [Fact]
        public async Task ReadAsync_Reads_Exactly_Its_Range()
        {
            byte[] bytes = Payload(300_000);
            using CountingSource source = new(bytes);
            await using ByteSourceStream stream = new(source, offset: 1000, length: 200_000);
            using MemoryStream copy = new();
            await stream.CopyToAsync(copy, TestContext.Current.CancellationToken);
            Assert.True(copy.ToArray().AsSpan().SequenceEqual(bytes.AsSpan(1000, 200_000)));
        }

        [Fact]
        public void Small_Reads_Are_Served_From_One_Buffered_Fetch()
        {
            using CountingSource source = new(Payload(60_000));
            using ByteSourceStream stream = new(source, offset: 0, length: 60_000);
            byte[] chunk = new byte[100];
            for (int i = 0; i < 600; i++)
            {
                stream.ReadExactly(chunk);
            }
            Assert.Equal(1, source.Reads);
        }

        [Fact]
        public void A_Large_Read_Bypasses_The_Buffer()
        {
            byte[] bytes = Payload(300_000);
            using CountingSource source = new(bytes);
            using ByteSourceStream stream = new(source, offset: 0, length: 300_000);
            byte[] big = new byte[300_000];
            Assert.Equal(300_000, stream.Read(big));
            Assert.Equal(1, source.Reads);
            Assert.True(big.AsSpan().SequenceEqual(bytes));
        }

        [Fact]
        public void A_Source_Shorter_Than_The_Range_Ends_The_Stream()
        {
            using CountingSource source = new(Payload(100));
            using ByteSourceStream stream = new(source, offset: 50, length: 500);
            byte[] buffer = new byte[1000];
            Assert.Equal(50, stream.Read(buffer));
            Assert.Equal(0, stream.Read(buffer));
        }

        [Fact]
        public void Dispose_Does_Not_Dispose_The_Source()
        {
            using CountingSource source = new(Payload(10));
            new ByteSourceStream(source, 0, 10).Dispose();
            Assert.False(source.Disposed);
        }

        [Fact]
        public void Read_After_Dispose_Throws()
        {
            using CountingSource source = new(Payload(1000));
            ByteSourceStream stream = new(source, 0, 1000);
            Assert.Equal(10, stream.Read(new byte[10]));
            stream.Dispose();
            Assert.Throws<ObjectDisposedException>(() => stream.Read(new byte[10]));
        }

        [Fact]
        public async Task ReadAsync_After_Dispose_Throws()
        {
            using CountingSource source = new(Payload(1000));
            ByteSourceStream stream = new(source, 0, 1000);
            await stream.DisposeAsync();
            await Assert.ThrowsAsync<ObjectDisposedException>(
                async () => await stream.ReadAsync(new byte[10], TestContext.Current.CancellationToken));
        }

        [Fact]
        public async Task A_Large_Async_Read_Bypasses_The_Buffer()
        {
            byte[] bytes = Payload(300_000);
            using CountingSource source = new(bytes);
            await using ByteSourceStream stream = new(source, 0, 300_000);
            byte[] big = new byte[300_000];
            Assert.Equal(300_000, await stream.ReadAsync(big, TestContext.Current.CancellationToken));
            Assert.Equal(1, source.Reads);
            Assert.True(big.AsSpan().SequenceEqual(bytes));
        }
    }
}
