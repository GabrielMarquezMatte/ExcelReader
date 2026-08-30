using Microsoft.Win32.SafeHandles;

namespace ExcelReader.Core.Parser.Internal
{
    // A forward-only view of a file starting at an arbitrary offset, backed by positional
    // RandomAccess reads rather than the handle's own file pointer. That is what lets N of these
    // coexist over one SafeFileHandle with no seek contention and no per-worker file open:
    // RandomAccess.Read is thread-safe and never mutates the handle's position.
    //
    // Reads run to the end of the file, not to the end of the caller's chunk. A chunk owns the
    // records that *start* inside it, and the record starting just before its end runs past that
    // end; a hard cut here would hand the parser a truncated final record, which it would happily
    // emit because EOF terminates a record. Bounding is therefore the worker's job (CsvChunkWorker),
    // which stops once a record starts at or after its chunk end.
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
            // Read-only stream: nothing is buffered on the write side.
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
