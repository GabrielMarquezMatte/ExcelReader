using System.Diagnostics.CodeAnalysis;

namespace ExcelReader.Core.Reader.Sources
{
    /// <summary>
    /// A seekable stream over a whole <see cref="ByteSource"/>. <see cref="ByteSource.FromStream"/> takes the
    /// source back out, so the paths that open a workbook from a stream read the source positionally.
    /// </summary>
    [SuppressMessage("Usage", "CA2213", Justification = "The source is disposed through Interlocked.Exchange so a taken source is not.")]
    internal sealed class SourceStream(ByteSource source) : Stream
    {
        private readonly ByteSource _source = source;
        private ByteSource? _owned = source;
        private long _position;

        public override bool CanRead => true;

        public override bool CanSeek => true;

        public override bool CanWrite => false;

        public override long Length => _source.Length;

        public override long Position
        {
            get => _position;
            set
            {
                ArgumentOutOfRangeException.ThrowIfNegative(value);
                _position = value;
            }
        }

        internal bool TryTakeSource([NotNullWhen(true)] out ByteSource? taken)
        {
            taken = Interlocked.Exchange(ref _owned, null);
            return taken is not null;
        }

        public override int Read(Span<byte> buffer)
        {
            int read = _source.Read(_position, buffer);
            _position += read;
            return read;
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return Read(buffer.AsSpan(offset, count));
        }

        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            int read = await _source.ReadAsync(_position, buffer, cancellationToken).ConfigureAwait(false);
            _position += read;
            return read;
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            return ReadAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            Position = origin switch
            {
                SeekOrigin.Begin => offset,
                SeekOrigin.Current => _position + offset,
                SeekOrigin.End => _source.Length + offset,
                _ => throw new ArgumentOutOfRangeException(nameof(origin)),
            };
            return _position;
        }

        public override void Flush()
        {
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
            if (disposing)
            {
                Interlocked.Exchange(ref _owned, null)?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
