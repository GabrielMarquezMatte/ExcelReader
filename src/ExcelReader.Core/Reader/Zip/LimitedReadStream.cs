using System.Buffers;
using ExcelReader.Core.Reader.Internal;
using ExcelReader.Core.Writer.Zip;

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

        // A sheet read counts from what the workbook's one-time loads (metadata, styles, shared strings)
        // inflated, so a single read stays within the limit while a re-read does not pay for earlier ones.
        internal DecompressedByteCounter ForRead()
        {
            return new DecompressedByteCounter(_limit, _limitName) { _total = Interlocked.Read(ref _total) };
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

    internal sealed class LimitedReadStream : ReadOnlyStream
    {
        internal const string ChecksumMismatch = "The ZIP entry's data does not match its CRC-32.";

        private readonly Stream _inner;
        private readonly DecompressedByteCounter? _totalCounter;
        private readonly DecompressedByteCounter? _entryCounter;
        private readonly long _declaredLength;
        private readonly uint? _expectedCrc;
        private long _produced;
        private uint _crc;

        // declaredLength is the entry's size from the central directory. Neither raw deflate nor a stored
        // entry can tell a cut-short stream from a complete one, so the length is the only end-of-data check.
        // expectedCrc, when given, is compared once the declared length has been produced.
        internal LimitedReadStream(
            Stream inner,
            long declaredLength,
            DecompressedByteCounter? totalCounter,
            string entryLimitName = "",
            long entryLimit = 0,
            uint? expectedCrc = null)
        {
            _inner = inner;
            _declaredLength = declaredLength;
            _totalCounter = totalCounter;
            _expectedCrc = expectedCrc;
            _entryCounter = entryLimit > 0 ? new DecompressedByteCounter(entryLimit, entryLimitName) : null;
        }

        public override bool CanRead => _inner.CanRead;

        public override void Flush()
        {
            _inner.Flush();
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return Count(_inner.Read(buffer, offset, count), buffer.AsSpan(offset, count));
        }

        public override int Read(Span<byte> buffer)
        {
            return Count(_inner.Read(buffer), buffer);
        }

        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            return Count(await _inner.ReadAsync(buffer, cancellationToken).ConfigureAwait(false), buffer.Span);
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            return ReadAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();
        }

        private bool VerifiesChecksum => _expectedCrc is not null || _inner is PrefetchStream { VerifiesChecksum: true };

        // A reader stops at the end of the sheet's rows; reading what follows still compares the entry's checksum.
        internal void ReadRemainderForChecksum()
        {
            if (!VerifiesChecksum || _produced >= _declaredLength)
            {
                return;
            }
            byte[] scratch = ArrayPool<byte>.Shared.Rent(16 * 1024);
            try
            {
                while (Read(scratch.AsSpan()) > 0)
                {
                }
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(scratch);
            }
        }

        internal async ValueTask ReadRemainderForChecksumAsync(CancellationToken ct)
        {
            if (!VerifiesChecksum || _produced >= _declaredLength)
            {
                return;
            }
            byte[] scratch = ArrayPool<byte>.Shared.Rent(16 * 1024);
            try
            {
                while (await ReadAsync(scratch.AsMemory(), ct).ConfigureAwait(false) > 0)
                {
                }
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(scratch);
            }
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

        private int Count(int read, ReadOnlySpan<byte> buffer)
        {
            if (read <= 0)
            {
                if (buffer.Length > 0 && _produced != _declaredLength)
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
            if (_expectedCrc is not uint expected)
            {
                return read;
            }
            _crc = Crc32.Append(_crc, buffer[..read]);
            if (_produced == _declaredLength && _crc != expected)
            {
                throw new InvalidDataException(ChecksumMismatch);
            }
            return read;
        }
    }
}
