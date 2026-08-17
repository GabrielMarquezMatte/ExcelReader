using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ExcelReader.Native
{
    /// <summary>
    /// Append-only accumulator, of a column's values or of a whole sheet's row blobs, that never
    /// copies what it already holds: values land in a chain of chunks, so growth costs one fresh
    /// chunk instead of reallocating and copying everything so far. <see cref="CopyTo"/> flattens
    /// the chain into the single destination block, which is the only copy any value ever pays.
    /// </summary>
    /// <remarks>
    /// This replaces the <see cref="List{T}"/> each <c>ColumnBuilder</c> field used to be, and the
    /// <see cref="Array.Resize"/>d <c>byte[]</c> <c>AccumulateAllRows</c> used to grow. Both double
    /// by allocating a bigger array and copying, so accumulating N elements allocates ~2N elements'
    /// worth and throws ~N away, with the tall ones landing on the large object heap. Chunks are
    /// capped below the LOH threshold, so nothing accumulated here reaches it at any row count.
    /// Measured on Data/65K_Records_Data.xlsb (Ryzen 7 5700X, .NET 10.0.10, --job Medium):
    /// NativeTypedParseBenchmark.ParseTyped 20.58 MB -> 11.73 MB, NativeRowReadBenchmark.ReadAllBlob
    /// 67.72 MB -> 21.07 MB (the remainder of which is the blob itself plus the caller's copy).
    /// </remarks>
    // ponytail: chunks are plain allocations, not pooled — a parse still allocates its column data
    // once, it just no longer allocates it repeatedly. Upgrade path if that last slice matters is
    // ArrayPool&lt;T&gt;.Shared plus an IDisposable ColumnBuilder returning its chunks in ParseTyped's
    // finally; deliberately not taken here because it trades a use-after-return hazard for bytes.
    internal sealed class ChunkedBuffer<T> where T : unmanaged
    {
        // 32 KiB per chunk once the buffer is warm: under the ~85 000-byte large object heap
        // threshold for every T this holds (8 bytes at most), and large enough that a 65K-row column
        // is a few dozen chunks rather than thousands of them.
        private const int MaxChunkBytes = 32 * 1024;
        // First chunk size, in elements. Small so a three-row sheet with 14 columns does not pay a
        // full-size chunk per column; chunks grow geometrically from here up to the cap, mirroring a
        // List's doubling without any of its copying.
        private const int InitialChunkLength = 256;
        private readonly List<T[]> _chunks = [];
        private T[] _current = [];
        private int _used; // elements written into _current
        internal int Count { get; private set; }
        /// <summary>Size of everything appended so far, in bytes — the exact size <see cref="CopyTo"/> needs.</summary>
        internal int ByteLength => checked(Count * Unsafe.SizeOf<T>());

        /// <summary>
        /// The most recently appended element, by reference, for the read-modify-write the validity
        /// bitmap does on the byte it is currently filling. Only valid after at least one
        /// <see cref="Add"/>.
        /// </summary>
        internal ref T Last => ref _current[_used - 1];

        internal void Add(T value)
        {
            if (_used == _current.Length)
            {
                Grow();
            }
            _current[_used++] = value;
            Count++;
        }

        internal void AddRange(ReadOnlySpan<T> values)
        {
            while (!values.IsEmpty)
            {
                if (_used == _current.Length)
                {
                    Grow();
                }
                int take = Math.Min(_current.Length - _used, values.Length);
                values[..take].CopyTo(_current.AsSpan(_used));
                _used += take;
                Count += take;
                values = values[take..];
            }
        }

        /// <summary>Writes every element appended so far into <paramref name="destination"/>, which
        /// must be at least <see cref="ByteLength"/> bytes.</summary>
        internal void CopyTo(Span<byte> destination)
        {
            int offset = 0;
            for (int index = 0; index < _chunks.Count; index++)
            {
                T[] chunk = _chunks[index];
                // Only the last chunk is partially filled — every earlier one was filled to capacity
                // before the next was allocated.
                int length = index == _chunks.Count - 1 ? _used : chunk.Length;
                ReadOnlySpan<byte> bytes = MemoryMarshal.AsBytes(chunk.AsSpan(0, length));
                bytes.CopyTo(destination[offset..]);
                offset += bytes.Length;
            }
        }

        // Chunk size climbs with the total accumulated so far, capped so no chunk reaches the LOH.
        // Already-written chunks are never touched, so this is growth without any copying.
        private void Grow()
        {
            int maxChunkLength = Math.Max(MaxChunkBytes / Unsafe.SizeOf<T>(), 1);
            int length = Math.Min(Math.Max(Count, InitialChunkLength), maxChunkLength);
            _current = new T[length];
            _chunks.Add(_current);
            _used = 0;
        }
    }
}
