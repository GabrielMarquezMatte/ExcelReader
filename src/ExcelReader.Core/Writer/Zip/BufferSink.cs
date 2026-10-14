using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Zip
{
    /// <summary>
    /// A write-only stream that appends to a <see cref="BiffBuffer"/> it does not own. While
    /// <see cref="Discard"/> is set it drops what it is given, which is how the final block a
    /// <c>DeflateStream</c> writes on dispose is kept out of a stream of segments.
    /// </summary>
    internal sealed class BufferSink : Stream
    {
        private readonly BiffBuffer _target;

        internal BufferSink(BiffBuffer target)
        {
            _target = target;
        }

        internal bool Discard { get; set; }

        public override bool CanRead
        {
            get
            {
                return false;
            }
        }

        public override bool CanSeek
        {
            get
            {
                return false;
            }
        }

        public override bool CanWrite
        {
            get
            {
                return true;
            }
        }

        public override long Length
        {
            get
            {
                throw new NotSupportedException();
            }
        }

        public override long Position
        {
            get
            {
                throw new NotSupportedException();
            }
            set
            {
                throw new NotSupportedException();
            }
        }

        public override void Flush()
        {
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            throw new NotSupportedException();
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
            Write(buffer.AsSpan(offset, count));
        }

        public override void Write(ReadOnlySpan<byte> buffer)
        {
            if (!Discard)
            {
                _target.Write(buffer);
            }
        }
    }
}
