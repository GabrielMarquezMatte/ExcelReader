namespace ExcelReader.Core.Reader.Sources
{
    /// <summary>
    /// Serves reads from blocks of <c>blockSize</c> bytes fetched once from an inner source, keeping the
    /// most recently used blocks up to <c>capacityBytes</c>.
    /// </summary>
    internal sealed class CachedByteSource(ByteSource inner, int blockSize, long capacityBytes) : ByteSource
    {
        private readonly int _maxBlocks = (int)Math.Clamp(capacityBytes / blockSize, 1, int.MaxValue);
        private readonly Dictionary<long, Entry> _blocks = [];
        private readonly Lock _gate = new();
        private long _clock;
        private int _disposed;

        internal override long Length => inner.Length;

        internal override int Read(long offset, Span<byte> destination)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(offset);
            if (offset >= Length || destination.IsEmpty)
            {
                return 0;
            }
            long index = offset / blockSize;
            byte[] block = Block(index);
            int start = (int)(offset - (index * blockSize));
            int count = Math.Min(destination.Length, block.Length - start);
            block.AsSpan(start, count).CopyTo(destination);
            return count;
        }

        internal override ValueTask<int> ReadAsync(long offset, Memory<byte> destination, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            return new ValueTask<int>(Read(offset, destination.Span));
        }

        public override void Dispose()
        {
            if (Interlocked.Exchange(ref _disposed, 1) != 0)
            {
                return;
            }
            lock (_gate)
            {
                _blocks.Clear();
            }
            inner.Dispose();
        }

        private byte[] Block(long index)
        {
            Entry entry;
            lock (_gate)
            {
                if (!_blocks.TryGetValue(index, out entry!))
                {
                    entry = new Entry(new Lazy<byte[]>(() => Fetch(index), LazyThreadSafetyMode.ExecutionAndPublication));
                    _blocks[index] = entry;
                    EvictBeyondCapacity(index);
                }
                entry.LastUse = ++_clock;
            }
            try
            {
                return entry.Data.Value;
            }
            catch
            {
                lock (_gate)
                {
                    if (_blocks.TryGetValue(index, out Entry? current) && ReferenceEquals(current, entry))
                    {
                        _blocks.Remove(index);
                    }
                }
                throw;
            }
        }

        private byte[] Fetch(long index)
        {
            long start = index * blockSize;
            byte[] data = new byte[(int)Math.Min(blockSize, Length - start)];
            inner.ReadExactly(start, data);
            return data;
        }

        // ponytail: linear scan for the oldest block; fine for the default 16 blocks, a linked list if capacities grow.
        private void EvictBeyondCapacity(long keep)
        {
            while (_blocks.Count > _maxBlocks)
            {
                long oldest = -1;
                long oldestUse = long.MaxValue;
                foreach ((long key, Entry candidate) in _blocks)
                {
                    if (key != keep && candidate.LastUse < oldestUse)
                    {
                        oldest = key;
                        oldestUse = candidate.LastUse;
                    }
                }
                _blocks.Remove(oldest);
            }
        }

        private sealed class Entry(Lazy<byte[]> data)
        {
            internal Lazy<byte[]> Data { get; } = data;

            internal long LastUse;
        }
    }
}
