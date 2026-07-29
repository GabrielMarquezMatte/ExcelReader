using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    // The memory-backed BufferedStreamCursor ctor used by the in-memory ZIP path once a ZipPart is
    // fully decompressed. No source stream, no refills — these tests cover
    // the ctor's aliasing/offset behavior and the one sharp edge: Return() must never hand a
    // caller-owned or ZipPart-owned array back to ArrayPool.
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

            // Passing a null source mirrors PooledStreamRowEnumerator's memory-backed ctor, where
            // _source is null; Fill/FillAsync must not dereference it.
            cursor.Fill(source: null);
            Assert.Equal(3, cursor.Len);
            Assert.True(cursor.Eof);
        }

        // Return() on a memory-backed cursor must never send `foreign` back to ArrayPool: it was
        // never rented from there, and doing so would corrupt the pool with an array some other owner
        // (the caller, or a ZipPart) is still holding onto. ArrayPool.Shared's per-bucket stack is
        // LIFO, so if Return() had wrongly pooled `foreign`, the very next same-size Rent on this
        // thread would very likely hand it straight back out.
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

        private sealed class NonArrayMemoryManager : MemoryManager<byte>
        {
            private readonly byte[] _data;

            internal NonArrayMemoryManager(byte[] data)
            {
                _data = data;
            }

            public override Span<byte> GetSpan()
            {
                return _data;
            }

            public override MemoryHandle Pin(int elementIndex = 0)
            {
                throw new NotSupportedException();
            }

            public override void Unpin()
            {
                throw new NotSupportedException();
            }

            [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP010:Call base.Dispose(disposing)",
                Justification = "MemoryManager<byte>.Dispose(bool) is abstract — there is no base implementation to call.")]
            protected override void Dispose(bool disposing)
            {
            }
        }
    }
}
