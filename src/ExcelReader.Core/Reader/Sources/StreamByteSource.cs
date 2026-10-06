namespace ExcelReader.Core.Reader.Sources
{
    /// <summary>
    /// A seekable stream has one cursor, so each read takes the gate for its seek and read only;
    /// decompression and parsing happen outside it.
    /// </summary>
    internal sealed class StreamByteSource : ByteSource
    {
        private readonly Stream _stream;
        private readonly bool _leaveOpen;
        private readonly SemaphoreSlim _gate = new(1, 1);

        internal StreamByteSource(Stream stream, bool leaveOpen)
        {
            _stream = stream;
            _leaveOpen = leaveOpen;
            Length = stream.Length;
        }

        internal override long Length { get; }

        internal override int Read(long offset, Span<byte> destination)
        {
            if (offset >= Length)
            {
                return 0;
            }
            _gate.Wait();
            try
            {
                _stream.Position = offset;
                return _stream.Read(destination);
            }
            finally
            {
                _gate.Release();
            }
        }

        internal override async ValueTask<int> ReadAsync(long offset, Memory<byte> destination, CancellationToken ct)
        {
            if (offset >= Length)
            {
                return 0;
            }
            await _gate.WaitAsync(ct).ConfigureAwait(false);
            try
            {
                _stream.Position = offset;
                return await _stream.ReadAsync(destination, ct).ConfigureAwait(false);
            }
            finally
            {
                _gate.Release();
            }
        }

        public override void Dispose()
        {
            if (!_leaveOpen)
            {
                _stream.Dispose();
            }
            _gate.Dispose();
        }
    }
}
