namespace ExcelReader.Core.Reader.Zip
{
    internal sealed class DecompressedByteCounter
    {
        private readonly long _limit;
        private readonly string _limitName;
        private long _total;

        internal DecompressedByteCounter(long limit, string limitName = nameof(ExcelReaderOptions.MaxTotalDecompressedBytes))
        {
            _limit = limit;
            _limitName = limitName;
        }

        internal void Add(long bytes)
        {
            if (bytes <= 0 || _limit <= 0)
            {
                return;
            }
            long total = Interlocked.Add(ref _total, bytes);
            if (total > _limit || total < 0)
            {
                throw new ExcelLimitExceededException(_limitName, _limit, total < 0 ? long.MaxValue : total);
            }
        }

        internal long Remaining => _limit <= 0 ? long.MaxValue : _limit - Interlocked.Read(ref _total);
    }

    internal sealed class LimitedReadStream : Stream
    {
        private readonly Stream _inner;
        private readonly DecompressedByteCounter? _totalCounter;
        private readonly DecompressedByteCounter? _entryCounter;
        private readonly long _declaredLength;
        private long _produced;

        // declaredLength is the entry's size from the central directory. Neither raw deflate nor a stored
        // entry can tell a cut-short stream from a complete one, so the length is the only end-of-data check.
        internal LimitedReadStream(
            Stream inner,
            long declaredLength,
            DecompressedByteCounter? totalCounter,
            string entryLimitName = "",
            long entryLimit = 0)
        {
            _inner = inner;
            _declaredLength = declaredLength;
            _totalCounter = totalCounter;
            _entryCounter = entryLimit > 0 ? new DecompressedByteCounter(entryLimit, entryLimitName) : null;
        }

        public override bool CanRead => _inner.CanRead;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush()
        {
            _inner.Flush();
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return Count(_inner.Read(buffer, offset, count), count);
        }

        public override int Read(Span<byte> buffer)
        {
            return Count(_inner.Read(buffer), buffer.Length);
        }

        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            return Count(await _inner.ReadAsync(buffer, cancellationToken).ConfigureAwait(false), buffer.Length);
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            return ReadAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();
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
            if (disposing)
            {
                _inner.Dispose();
            }
            base.Dispose(disposing);
        }

        public override async ValueTask DisposeAsync()
        {
            await _inner.DisposeAsync().ConfigureAwait(false);
            await base.DisposeAsync().ConfigureAwait(false);
        }

        private int Count(int read, int requested)
        {
            if (read <= 0)
            {
                if (requested > 0 && _produced != _declaredLength)
                {
                    throw new InvalidDataException("The ZIP entry produced less data than its declared uncompressed size.");
                }
                return read;
            }
            _produced += read;
            if (_produced > _declaredLength)
            {
                throw new InvalidDataException("The ZIP entry produced more data than its declared uncompressed size.");
            }
            _totalCounter?.Add(read);
            _entryCounter?.Add(read);
            return read;
        }
    }
}
