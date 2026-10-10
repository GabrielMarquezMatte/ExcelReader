using ExcelReader.Core.Reader.Sources;

namespace ExcelReader.Tests.Reader.Sources
{
    public sealed class CachedByteSourceTests
    {
        private static readonly byte[] Payload = [.. Enumerable.Range(0, 10_000).Select(i => (byte)(i * 7))];

        [Fact]
        public void ReadExactly_Across_Blocks_Returns_The_Inner_Bytes()
        {
            using CachedByteSource cache = new(new CountingSource(Payload), blockSize: 1024, capacityBytes: 1 << 20);
            byte[] read = new byte[3000];
            cache.ReadExactly(500, read);
            Assert.Equal(Payload[500..3500], read);
        }

        [Fact]
        public void Each_Block_Is_Fetched_Once_While_It_Fits()
        {
            CountingSource inner = new(Payload);
            using CachedByteSource cache = new(inner, blockSize: 1024, capacityBytes: 1 << 20);
            byte[] chunk = new byte[100];
            for (int pass = 0; pass < 2; pass++)
            {
                for (int offset = 0; offset < Payload.Length; offset += chunk.Length)
                {
                    cache.ReadExactly(offset, chunk);
                }
            }
            Assert.Equal(10, inner.Reads);
        }

        [Fact]
        public void Concurrent_Readers_Of_One_Block_Fetch_It_Once()
        {
            CountingSource inner = new(Payload) { Delay = TimeSpan.FromMilliseconds(50) };
            using CachedByteSource cache = new(inner, blockSize: 4096, capacityBytes: 1 << 20);
            Parallel.For(0, 16, _ =>
            {
                byte[] read = new byte[64];
                cache.ReadExactly(0, read);
            });
            Assert.Equal(1, inner.Reads);
        }

        [Fact]
        public void The_Least_Recently_Used_Block_Is_Evicted_Past_Capacity()
        {
            CountingSource inner = new(Payload);
            using CachedByteSource cache = new(inner, blockSize: 1024, capacityBytes: 2048);
            byte[] one = new byte[1];
            foreach (long offset in new long[] { 0, 1024, 2048, 1024, 0 })
            {
                cache.ReadExactly(offset, one);
            }
            Assert.Equal(4, inner.Reads);
        }

        [Fact]
        public void A_Read_At_Or_Past_The_End_Returns_Zero()
        {
            using CachedByteSource cache = new(new CountingSource(Payload), blockSize: 1024, capacityBytes: 1 << 20);
            Assert.Equal(0, cache.Read(Payload.Length, new byte[10]));
            Assert.Equal(0, cache.Read(Payload.Length + 5, new byte[10]));
        }

        [Fact]
        public void A_Failed_Fetch_Is_Tried_Again()
        {
            CountingSource inner = new(Payload) { FailReads = 1 };
            using CachedByteSource cache = new(inner, blockSize: 1024, capacityBytes: 1 << 20);
            byte[] read = new byte[10];
            Assert.Throws<IOException>(() => cache.ReadExactly(0, read));
            cache.ReadExactly(0, read);
            Assert.Equal(Payload[..10], read);
        }

        [Fact]
        public void Dispose_Disposes_The_Inner_Source_Once()
        {
            CountingSource inner = new(Payload);
            CachedByteSource cache = new(inner, blockSize: 1024, capacityBytes: 1 << 20);
            cache.Dispose();
            cache.Dispose();
            Assert.Equal(1, inner.Disposals);
        }
    }
}
