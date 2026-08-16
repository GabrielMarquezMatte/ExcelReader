using System.Runtime.InteropServices;
using ExcelReader.Native;

namespace ExcelReader.Tests
{
    // ChunkedBuffer's whole point is that its storage is discontiguous, so every test here pushes
    // past its first chunk (256 elements) and well past the point where chunk size hits its cap —
    // a bug in the chunk-boundary arithmetic is invisible below that.
    public sealed class ChunkedBufferTests
    {
        private const int Count = 30_000;

        [Fact]
        public void Add_Should_Preserve_Every_Element_Across_Chunk_Boundaries()
        {
            ChunkedBuffer<int> buffer = new();
            for (int i = 0; i < Count; i++)
            {
                buffer.Add(i);
            }

            Assert.Equal(Count, buffer.Count);
            Assert.Equal(Count * sizeof(int), buffer.ByteLength);

            int[] copied = Flatten(buffer);
            for (int i = 0; i < Count; i++)
            {
                Assert.Equal(i, copied[i]);
            }
        }

        [Fact]
        public void AddRange_Should_Split_A_Span_Across_Chunks_Without_Losing_Bytes()
        {
            ChunkedBuffer<byte> buffer = new();
            List<byte> expected = [];
            byte[] source = new byte[997]; // deliberately coprime with every chunk size
            for (int i = 0; i < source.Length; i++)
            {
                source[i] = (byte)(i * 7);
            }

            for (int round = 0; round < 100; round++)
            {
                int length = round % source.Length;
                buffer.AddRange(source.AsSpan(0, length));
                expected.AddRange(source.AsSpan(0, length).ToArray());
            }

            Assert.Equal(expected.Count, buffer.Count);
            Assert.Equal(expected, Flatten(buffer));
        }

        [Fact]
        public void AddRange_And_Add_Should_Interleave_Correctly()
        {
            ChunkedBuffer<byte> buffer = new();
            List<byte> expected = [];
            for (int i = 0; i < 5_000; i++)
            {
                buffer.Add((byte)i);
                expected.Add((byte)i);
                buffer.AddRange([1, 2, 3]);
                expected.AddRange([1, 2, 3]);
            }

            Assert.Equal(expected, Flatten(buffer));
        }

        [Fact]
        public void Last_Should_Address_The_Most_Recent_Element_After_A_Chunk_Rollover()
        {
            // The validity bitmap's read-modify-write goes through Last, so it has to keep working
            // on the first element of a freshly allocated chunk, not just mid-chunk.
            ChunkedBuffer<byte> buffer = new();
            for (int i = 0; i < Count; i++)
            {
                buffer.Add(0);
                buffer.Last |= 0x5A;
            }

            Assert.All(Flatten(buffer), value => Assert.Equal(0x5A, value));
        }

        [Fact]
        public void Empty_Buffer_Should_Report_Nothing_And_Copy_Nothing()
        {
            ChunkedBuffer<long> buffer = new();

            Assert.Equal(0, buffer.Count);
            Assert.Equal(0, buffer.ByteLength);
            buffer.CopyTo(Span<byte>.Empty); // must not throw
        }

        private static T[] Flatten<T>(ChunkedBuffer<T> buffer) where T : unmanaged
        {
            byte[] bytes = new byte[buffer.ByteLength];
            buffer.CopyTo(bytes);
            return MemoryMarshal.Cast<byte, T>(bytes).ToArray();
        }
    }
}
