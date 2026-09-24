using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Tests.Reader
{
    public class BufferedStreamCursorTests
    {
        [Fact]
        public void MemoryCtorAliasesAWholeArrayStartingAtZero()
        {
            byte[] content = [1, 2, 3, 4, 5];
            var cursor = new BufferedStreamCursor(content.AsMemory(), maxCellBytes: 0, limitName: "Test");

            Assert.Same(content, cursor.Buf);
            Assert.Equal(0, cursor.Pos);
            Assert.Equal(content.Length, cursor.Len);
            Assert.True(cursor.Eof);
        }

        [Fact]
        public void MemoryCtorPreservesTheOffsetOfASlicedArray()
        {
            byte[] backing = [0xAA, 0xAA, 10, 20, 30, 0xAA];
            ReadOnlyMemory<byte> slice = backing.AsMemory(2, 3);
            var cursor = new BufferedStreamCursor(slice, maxCellBytes: 0, limitName: "Test");

            Assert.Same(backing, cursor.Buf);
            Assert.Equal(2, cursor.Pos);
            Assert.Equal(5, cursor.Len);
            Assert.Equal([10, 20, 30], cursor.Buf.AsSpan(cursor.Pos, cursor.Len - cursor.Pos).ToArray());
        }

        [Fact]
        public void MemoryCtorCopiesWhenTheSourceIsNotArrayBacked()
        {
            byte[] payload = [7, 8, 9];
            var manager = new NonArrayMemoryManager(payload);
            var cursor = new BufferedStreamCursor(manager.Memory, maxCellBytes: 0, limitName: "Test");

            Assert.NotSame(payload, cursor.Buf);
            Assert.Equal(0, cursor.Pos);
            Assert.Equal(payload.Length, cursor.Len);
            Assert.Equal(payload, cursor.Buf.AsSpan(0, cursor.Len).ToArray());
        }

        [Fact]
        public void FillIsANoOpOnceEofIsSetAtConstruction()
        {
            byte[] content = [1, 2, 3];
            var cursor = new BufferedStreamCursor(content.AsMemory(), maxCellBytes: 0, limitName: "Test");

            cursor.Fill(source: null);
            Assert.Equal(3, cursor.Len);
            Assert.True(cursor.Eof);
        }

        [Fact]
        public void ReturnDoesNotPoolAMemoryBackedCursorsForeignArray()
        {
            byte[] foreign = new byte[64];
            var cursor = new BufferedStreamCursor(foreign.AsMemory(), maxCellBytes: 0, limitName: "Test");

            cursor.Return();
            Assert.Empty(cursor.Buf);

            var rented = new List<byte[]>();
            try
            {
                for (int i = 0; i < 8; i++)
                {
                    byte[] candidate = ArrayPool<byte>.Shared.Rent(64);
                    rented.Add(candidate);
                    Assert.NotSame(foreign, candidate);
                }
            }
            finally
            {
                foreach (ref readonly var candidate in CollectionsMarshal.AsSpan(rented))
                {
                    ArrayPool<byte>.Shared.Return(candidate);
                }
            }
        }

        [Fact]
        [SuppressMessage("Security", "CA5394:Do not use insecure randomness",
            Justification = "Needs a reproducible seeded PRNG for deterministic test content, not cryptographic randomness.")]
        public void PooledCursorReassemblesContentAcrossManySmallRefillsRegardlessOfCompactOrGrowPath()
        {
            byte[] expected = new byte[500];
            new Random(12345).NextBytes(expected);
            using var source = new SmallChunkStream(expected, chunkSize: 7);
            var cursor = new BufferedStreamCursor(maxCellBytes: 0, limitName: "Test", initialCapacity: 16);
            try
            {
                byte[] actual = new byte[expected.Length];
                int written = 0;
                while (written < actual.Length)
                {
                    int step = Math.Min(3, actual.Length - written);
                    cursor.Ensure(source, step);
                    Assert.True(cursor.Len - cursor.Pos >= step);
                    cursor.Buf.AsSpan(cursor.Pos, step).CopyTo(actual.AsSpan(written));
                    cursor.Pos += step;
                    written += step;
                }
                Assert.Equal(expected, actual);
            }
            finally
            {
                cursor.Return();
            }
        }

        private sealed class SmallChunkStream(byte[] data, int chunkSize) : MemoryStream(data)
        {
            public override int Read(byte[] buffer, int offset, int count)
            {
                return base.Read(buffer, offset, Math.Min(count, chunkSize));
            }
        }

        [Fact]
        public void BaseOffsetStartsAtZeroForAStreamCursor()
        {
            var cursor = new BufferedStreamCursor(maxCellBytes: 1 << 20, limitName: "Test", initialCapacity: 16);

            Assert.Equal(0, cursor.BaseOffset);
        }

        [Fact]
        public void BaseOffsetPlusPosTracksTheSourceOffsetAcrossCompaction()
        {
            byte[] content = new byte[64];
            for (int i = 0; i < content.Length; i++)
            {
                content[i] = (byte)i;
            }
            using var source = new MemoryStream(content, writable: false);
            var cursor = new BufferedStreamCursor(maxCellBytes: 1 << 20, limitName: "Test", initialCapacity: 16);

            cursor.Fill(source);
            cursor.Pos = cursor.Len - 1;
            long sourceOffsetOfNextByte = cursor.BaseOffset + cursor.Pos;
            cursor.Fill(source);

            Assert.Equal(sourceOffsetOfNextByte, cursor.BaseOffset + cursor.Pos);
            Assert.Equal(content[sourceOffsetOfNextByte], cursor.Buf[cursor.Pos]);
        }

        [Fact]
        public void BaseOffsetRebasesASlicedMemoryCursorSoOffsetsStayZeroBased()
        {
            byte[] backing = [0xAA, 0xAA, 10, 20, 30, 0xAA];
            var cursor = new BufferedStreamCursor(backing.AsMemory(2, 3), maxCellBytes: 0, limitName: "Test");

            Assert.Equal(0, cursor.BaseOffset + cursor.Pos);
            cursor.Pos += 2;
            Assert.Equal(2, cursor.BaseOffset + cursor.Pos);
        }
    }
}
