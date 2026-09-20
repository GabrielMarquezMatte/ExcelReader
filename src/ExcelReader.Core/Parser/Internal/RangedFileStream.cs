using Microsoft.Win32.SafeHandles;

namespace ExcelReader.Core.Parser.Internal
{
    internal sealed class RangedFileStream : Stream
    {
        private readonly SafeFileHandle _handle;
        private long _offset;

        internal RangedFileStream(SafeFileHandle handle, long start)
        {
            ArgumentNullException.ThrowIfNull(handle);
            ArgumentOutOfRangeException.ThrowIfNegative(start);
            _handle = handle;
            _offset = start;
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

        public override int Read(Span<byte> buffer)
        {
            int n = RandomAccess.Read(_handle, buffer, _offset);
            _offset += n;
            return n;
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            ValidateBufferArguments(buffer, offset, count);
            return Read(buffer.AsSpan(offset, count));
        }

        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            int n = await RandomAccess.ReadAsync(_handle, buffer, _offset, cancellationToken).ConfigureAwait(false);
            _offset += n;
            return n;
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            ValidateBufferArguments(buffer, offset, count);
            return ReadAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();
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
    }
}
