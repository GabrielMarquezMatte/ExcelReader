using System.Runtime.CompilerServices;
using ExcelReader.Core.Reader.Internal;
using Microsoft.Win32.SafeHandles;

namespace ExcelReader.Core.Parser.ParallelCsv
{
    internal sealed class RangedFileStream : ReadOnlyStream
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

        // Pooled: a state machine box per chunk read was the parallel parse's largest allocation after the rows.
        [AsyncMethodBuilder(typeof(PoolingAsyncValueTaskMethodBuilder<>))]
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
    }
}
