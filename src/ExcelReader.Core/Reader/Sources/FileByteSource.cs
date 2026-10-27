using Microsoft.Win32.SafeHandles;

namespace ExcelReader.Core.Reader.Sources
{
    internal sealed class FileByteSource : ByteSource
    {
        private readonly SafeFileHandle _handle;
        private readonly IDisposable? _owner;
        private readonly long _start;

        internal FileByteSource(SafeFileHandle handle, long length, IDisposable? owner, long start = 0)
        {
            _handle = handle;
            Length = length;
            _owner = owner;
            _start = start;
        }

        internal override long Length { get; }

        internal override int Read(long offset, Span<byte> destination)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(offset);
            return offset >= Length ? 0 : RandomAccess.Read(_handle, destination, _start + offset);
        }

        internal override ValueTask<int> ReadAsync(long offset, Memory<byte> destination, CancellationToken ct)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(offset);
            return offset >= Length ? new ValueTask<int>(0) : RandomAccess.ReadAsync(_handle, destination, _start + offset, ct);
        }

        public override void Dispose()
        {
            _owner?.Dispose();
        }
    }
}
