namespace ExcelReader.Core.Reader.Internal
{
    /// <summary>
    /// A read-only stream, forward-only unless a subclass overrides <see cref="CanSeek"/>,
    /// <see cref="Length"/>, <see cref="Position"/> and <see cref="Seek"/>. Subclasses supply the reads.
    /// </summary>
    internal abstract class ReadOnlyStream : Stream
    {
        public override bool CanRead => true;

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
