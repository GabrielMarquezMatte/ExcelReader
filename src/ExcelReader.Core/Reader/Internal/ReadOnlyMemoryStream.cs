namespace ExcelReader.Core.Reader.Internal
{
    /// <summary>A seekable read-only stream over a <see cref="ReadOnlyMemory{T}"/>, whatever backs it, without copying it.</summary>
    internal sealed class ReadOnlyMemoryStream(ReadOnlyMemory<byte> memory) : ReadOnlyStream
    {
        private long _position;

        public override bool CanSeek => true;

        public override long Length => memory.Length;

        public override long Position
        {
            get => _position;
            set
            {
                ArgumentOutOfRangeException.ThrowIfNegative(value);
                _position = value;
            }
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            ValidateBufferArguments(buffer, offset, count);
            return Read(buffer.AsSpan(offset, count));
        }

        public override int Read(Span<byte> buffer)
        {
            long remaining = memory.Length - _position;
            if (remaining <= 0)
            {
                return 0;
            }
            int count = (int)Math.Min(buffer.Length, remaining);
            memory.Span.Slice((int)_position, count).CopyTo(buffer);
            _position += count;
            return count;
        }

        public override int ReadByte()
        {
            if (_position >= memory.Length)
            {
                return -1;
            }
            return memory.Span[(int)_position++];
        }

        public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            if (cancellationToken.IsCancellationRequested)
            {
                return ValueTask.FromCanceled<int>(cancellationToken);
            }
            return new ValueTask<int>(Read(buffer.Span));
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            ValidateBufferArguments(buffer, offset, count);
            return ReadAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            long position = origin switch
            {
                SeekOrigin.Begin => offset,
                SeekOrigin.Current => _position + offset,
                SeekOrigin.End => memory.Length + offset,
                _ => throw new ArgumentOutOfRangeException(nameof(origin)),
            };
            if (position < 0)
            {
                throw new IOException("An attempt was made to move the position before the beginning of the stream.");
            }
            _position = position;
            return position;
        }
    }
}
