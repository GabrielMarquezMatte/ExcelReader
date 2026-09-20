using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ExcelReader.Native
{
    // ponytail: chunks are plain allocations, not pooled — a parse still allocates its column data
    internal sealed class ChunkedBuffer<T> where T : unmanaged
    {
        private const int MaxChunkBytes = 32 * 1024;
        private const int InitialChunkLength = 256;
        private readonly List<T[]> _chunks = [];
        private T[] _current = [];
        private int _used;
        internal int Count { get; private set; }
        internal int ByteLength => checked(Count * Unsafe.SizeOf<T>());

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

        internal void CopyTo(Span<byte> destination)
        {
            int offset = 0;
            for (int index = 0; index < _chunks.Count; index++)
            {
                T[] chunk = _chunks[index];
                int length = index == _chunks.Count - 1 ? _used : chunk.Length;
                ReadOnlySpan<byte> bytes = MemoryMarshal.AsBytes(chunk.AsSpan(0, length));
                bytes.CopyTo(destination[offset..]);
                offset += bytes.Length;
            }
        }

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
