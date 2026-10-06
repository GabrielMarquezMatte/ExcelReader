using System.Buffers;

namespace ExcelReader.Core.Reader.Sources
{
    /// <summary>Forward-only view of one byte range of a <see cref="ByteSource"/>, with its own offset.</summary>
    internal sealed class ByteSourceStream : Stream
    {
        private const int BufferSize = 64 * 1024;

        private readonly ByteSource _source;
        private readonly long _end;
        private long _next;
        private byte[]? _buffer;
        private int _bufferPos;
        private int _bufferLen;

        internal ByteSourceStream(ByteSource source, long offset, long length)
        {
            _source = source;
            _next = offset;
            _end = offset + length;
        }

        public override bool CanRead => true;

        public override bool CanSeek => false;

        public override bool CanWrite => false;

        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return Read(buffer.AsSpan(offset, count));
        }

        public override int Read(Span<byte> destination)
        {
            if (destination.IsEmpty)
            {
                return 0;
            }
            if (_bufferPos == _bufferLen)
            {
                long remaining = _end - _next;
                if (remaining <= 0)
                {
                    return 0;
                }
                if (destination.Length >= BufferSize)
                {
                    int direct = _source.Read(_next, destination[..(int)Math.Min(destination.Length, remaining)]);
                    _next += direct;
                    return direct;
                }
                _buffer ??= ArrayPool<byte>.Shared.Rent(BufferSize);
                _bufferLen = _source.Read(_next, _buffer.AsSpan(0, (int)Math.Min(BufferSize, remaining)));
                _bufferPos = 0;
                _next += _bufferLen;
                if (_bufferLen == 0)
                {
                    return 0;
                }
            }
            return TakeBuffered(destination);
        }

        public override async ValueTask<int> ReadAsync(Memory<byte> destination, CancellationToken cancellationToken = default)
        {
            if (destination.IsEmpty)
            {
                return 0;
            }
            if (_bufferPos == _bufferLen)
            {
                long remaining = _end - _next;
                if (remaining <= 0)
                {
                    return 0;
                }
                if (destination.Length >= BufferSize)
                {
                    int direct = await _source.ReadAsync(
                        _next, destination[..(int)Math.Min(destination.Length, remaining)], cancellationToken).ConfigureAwait(false);
                    _next += direct;
                    return direct;
                }
                _buffer ??= ArrayPool<byte>.Shared.Rent(BufferSize);
                _bufferLen = await _source.ReadAsync(
                    _next, _buffer.AsMemory(0, (int)Math.Min(BufferSize, remaining)), cancellationToken).ConfigureAwait(false);
                _bufferPos = 0;
                _next += _bufferLen;
                if (_bufferLen == 0)
                {
                    return 0;
                }
            }
            return TakeBuffered(destination.Span);
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            return ReadAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();
        }

        private int TakeBuffered(Span<byte> destination)
        {
            int count = Math.Min(destination.Length, _bufferLen - _bufferPos);
            _buffer.AsSpan(_bufferPos, count).CopyTo(destination);
            _bufferPos += count;
            return count;
        }

        public override void Flush()
        {
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            throw new NotSupportedException();
        }

        public override void SetLength(long value)
        {
            throw new NotSupportedException();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            throw new NotSupportedException();
        }

        protected override void Dispose(bool disposing)
        {
            byte[]? buffer = Interlocked.Exchange(ref _buffer, null);
            if (buffer is not null)
            {
                ArrayPool<byte>.Shared.Return(buffer);
            }
            base.Dispose(disposing);
        }
    }
}
