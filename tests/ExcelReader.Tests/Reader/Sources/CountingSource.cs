using ExcelReader.Core.Reader.Sources;

namespace ExcelReader.Tests.Reader.Sources
{
    internal sealed class CountingSource(byte[] bytes) : ByteSource
    {
        internal int Reads;
        internal int Disposals;
        internal int FailReads;
        internal TimeSpan Delay { get; init; }

        internal override long Length => bytes.Length;

        internal override int Read(long offset, Span<byte> destination)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(offset);
            Interlocked.Increment(ref Reads);
            if (Delay > TimeSpan.Zero)
            {
                Thread.Sleep(Delay);
            }
            if (Interlocked.Decrement(ref FailReads) >= 0)
            {
                throw new IOException("transient");
            }
            if (offset >= bytes.Length)
            {
                return 0;
            }
            int count = (int)Math.Min(destination.Length, bytes.Length - offset);
            bytes.AsSpan((int)offset, count).CopyTo(destination);
            return count;
        }

        internal override ValueTask<int> ReadAsync(long offset, Memory<byte> destination, CancellationToken ct)
        {
            return new ValueTask<int>(Read(offset, destination.Span));
        }

        public override void Dispose()
        {
            Interlocked.Increment(ref Disposals);
        }
    }
}
