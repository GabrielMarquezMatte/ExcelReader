using Microsoft.Win32.SafeHandles;

namespace ExcelReader.Core.Reader.Sources
{
    internal sealed class FileByteSource : ByteSource
    {
        private readonly SafeFileHandle _handle;
        private readonly IDisposable? _owner;

        internal FileByteSource(SafeFileHandle handle, long length, IDisposable? owner)
        {
            _handle = handle;
            Length = length;
            _owner = owner;
        }

        internal override long Length { get; }

        internal override int Read(long offset, Span<byte> destination)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(offset);
            return offset >= Length ? 0 : RandomAccess.Read(_handle, destination, offset);
        }

        internal override ValueTask<int> ReadAsync(long offset, Memory<byte> destination, CancellationToken ct)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(offset);
            return offset >= Length ? new ValueTask<int>(0) : RandomAccess.ReadAsync(_handle, destination, offset, ct);
        }

        public override void Dispose()
        {
            _owner?.Dispose();
        }
    }
}
