namespace ExcelReader.Core.Reader.Sources
{
    internal sealed class MemoryByteSource : ByteSource
    {
        private readonly ReadOnlyMemory<byte> _memory;
        private readonly IDisposable? _owner;

        internal MemoryByteSource(ReadOnlyMemory<byte> memory, IDisposable? owner)
        {
            _memory = memory;
            _owner = owner;
        }

        internal override long Length => _memory.Length;

        internal override int Read(long offset, Span<byte> destination)
        {
            if (offset < 0 || offset >= _memory.Length)
            {
                return 0;
            }
            int count = (int)Math.Min(destination.Length, _memory.Length - offset);
            _memory.Span.Slice((int)offset, count).CopyTo(destination);
            return count;
        }

        internal override ValueTask<int> ReadAsync(long offset, Memory<byte> destination, CancellationToken ct)
        {
            return new ValueTask<int>(Read(offset, destination.Span));
        }

        internal override bool TryGetMemory(out ReadOnlyMemory<byte> memory)
        {
            memory = _memory;
            return true;
        }

        public override void Dispose()
        {
            _owner?.Dispose();
        }
    }
}
