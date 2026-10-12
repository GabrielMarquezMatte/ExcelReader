namespace ExcelReader.Core.Reader.Csv
{
    /// <summary>
    /// Serves a forward-only stream to one pass at a time, and lets a later pass start over for as long
    /// as everything read so far still fits in <c>limit</c> bytes. Disposing ends a pass;
    /// <see cref="Release"/> closes the stream underneath.
    /// </summary>
    internal sealed class ReplayStream(Stream inner, bool leaveOpen, int limit) : Stream
    {
        private MemoryStream? _recorded = new();
        private int _position;
        private int _busy;

        public override bool CanRead => true;

        public override bool CanSeek => false;

        public override bool CanWrite => false;

        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        internal bool TryBeginPass()
        {
            if (Interlocked.CompareExchange(ref _busy, 1, 0) != 0)
            {
                return false;
            }
            if (_recorded is null)
            {
                Volatile.Write(ref _busy, 0);
                return false;
            }
            _position = 0;
            return true;
        }

        internal void Release()
        {
            _recorded = null;
            if (!leaveOpen)
            {
                inner.Dispose();
            }
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return Read(buffer.AsSpan(offset, count));
        }

        public override int Read(Span<byte> buffer)
        {
            if (TryReplay(buffer, out int replayed))
            {
                return replayed;
            }
            int read = inner.Read(buffer);
            Record(buffer[..read]);
            return read;
        }

        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            if (TryReplay(buffer.Span, out int replayed))
            {
                return replayed;
            }
            int read = await inner.ReadAsync(buffer, cancellationToken).ConfigureAwait(false);
            Record(buffer.Span[..read]);
            return read;
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            return ReadAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();
        }

        private bool TryReplay(Span<byte> buffer, out int count)
        {
            count = 0;
            if (_recorded is not { } recorded || _position >= recorded.Length || buffer.IsEmpty)
            {
                return false;
            }
            count = Math.Min(buffer.Length, (int)recorded.Length - _position);
            recorded.GetBuffer().AsSpan(_position, count).CopyTo(buffer);
            _position += count;
            return true;
        }

        private void Record(ReadOnlySpan<byte> data)
        {
            if (_recorded is not { } recorded)
            {
                return;
            }
            if (recorded.Length + data.Length > limit)
            {
                _recorded = null;
                return;
            }
            recorded.Write(data);
            _position += data.Length;
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
            Volatile.Write(ref _busy, 0);
            base.Dispose(disposing);
        }
    }
}
