namespace ExcelReader.Core.Reader
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
            long total = checked(_total + bytes);
            if (total > _limit)
            {
                throw new ExcelLimitExceededException(_limitName, _limit, total);
            }
            _total = total;
        }

        // Budget left before this counter's own limit trips; used to bound a single entry's declared
        // size against what the workbook-wide counter has left, not just its absolute cap.
        internal long Remaining => _limit <= 0 ? long.MaxValue : _limit - _total;
    }

    internal sealed class LimitedReadStream : Stream
    {
        private readonly Stream _inner;
        private readonly DecompressedByteCounter? _totalCounter;
        private readonly DecompressedByteCounter? _entryCounter;

        internal LimitedReadStream(
            Stream inner,
            DecompressedByteCounter? totalCounter,
            string entryLimitName = "",
            long entryLimit = 0)
        {
            _inner = inner;
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
            int read = _inner.Read(buffer, offset, count);
            Count(read);
            return read;
        }

        public override int Read(Span<byte> buffer)
        {
            int read = _inner.Read(buffer);
            Count(read);
            return read;
        }

        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            int read = await _inner.ReadAsync(buffer, cancellationToken).ConfigureAwait(false);
            Count(read);
            return read;
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

        [System.Diagnostics.CodeAnalysis.SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "LimitedReadStream owns the wrapped entry stream for the duration of the decorated read.")]
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _inner.Dispose();
            }
            base.Dispose(disposing);
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "LimitedReadStream owns the wrapped entry stream for the duration of the decorated read.")]
        public override async ValueTask DisposeAsync()
        {
            await _inner.DisposeAsync().ConfigureAwait(false);
            await base.DisposeAsync().ConfigureAwait(false);
        }

        private void Count(int read)
        {
            if (read <= 0)
            {
                return;
            }
            _totalCounter?.Add(read);
            _entryCounter?.Add(read);
        }
    }
}
