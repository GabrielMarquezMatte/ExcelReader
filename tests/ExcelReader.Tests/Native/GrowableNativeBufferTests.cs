using System.Runtime.InteropServices;
using ExcelReader.Native;

namespace ExcelReader.Tests.Native
{
    public sealed class GrowableNativeBufferTests
    {
        private const int Count = 30_000;

        [Fact]
        public void Add_Should_Preserve_Every_Element_Across_Growth()
        {
            using GrowableNativeBuffer<int> buffer = new();
            for (int i = 0; i < Count; i++)
            {
                buffer.Add(i);
            }

            Assert.Equal(Count, buffer.Count);
            Assert.Equal(Count * sizeof(int), buffer.ByteLength);
            Assert.Equal(Enumerable.Range(0, Count), Flatten(buffer));
        }

        [Fact]
        public void AddRange_And_Add_Should_Interleave_Correctly()
        {
            using GrowableNativeBuffer<byte> buffer = new();
            List<byte> expected = [];
            byte[] source = new byte[997];
            for (int i = 0; i < source.Length; i++)
            {
                source[i] = (byte)(i * 7);
            }

            for (int round = 0; round < 300; round++)
            {
                buffer.Add((byte)round);
                expected.Add((byte)round);
                int length = round * 13 % source.Length;
                buffer.AddRange(source.AsSpan(0, length));
                expected.AddRange(source.AsSpan(0, length).ToArray());
            }

            Assert.Equal(expected, Flatten(buffer));
        }

        [Fact]
        public void Last_Should_Address_The_Most_Recent_Element_After_Growth()
        {
            using GrowableNativeBuffer<byte> buffer = new();
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
            using GrowableNativeBuffer<long> buffer = new();

            Assert.Equal(0, buffer.Count);
            Assert.Equal(0, buffer.ByteLength);
            buffer.CopyTo(Span<byte>.Empty);
        }

        [Fact]
        public unsafe void Detach_Should_Hand_Over_The_Block_And_Leave_The_Buffer_Empty()
        {
            using GrowableNativeBuffer<long> buffer = new();
            for (long i = 0; i < 1000; i++)
            {
                buffer.Add(i * 3);
            }

            IntPtr block = buffer.Detach();
            try
            {
                Assert.Equal(Enumerable.Range(0, 1000).Select(i => i * 3L), new ReadOnlySpan<long>((void*)block, 1000).ToArray());
                Assert.Equal(0, buffer.Count);
            }
            finally
            {
                Marshal.FreeHGlobal(block);
            }
        }

        [Fact]
        public void Dispose_Twice_Should_Be_Harmless()
        {
            GrowableNativeBuffer<int> buffer = new();
            buffer.Add(1);

            buffer.Dispose();
            buffer.Dispose();

            Assert.Equal(0, buffer.Count);
        }

        private static T[] Flatten<T>(GrowableNativeBuffer<T> buffer) where T : unmanaged
        {
            byte[] bytes = new byte[buffer.ByteLength];
            buffer.CopyTo(bytes);
            return MemoryMarshal.Cast<byte, T>(bytes).ToArray();
        }
    }
}
